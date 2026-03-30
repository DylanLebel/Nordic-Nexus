# ==============================================================================
#  HubService.ps1  v2.10 - Nordic Minesteel Technologies
#  Background Service Engine & API Controller (with System Tray)
# ==============================================================================

param(
    [switch]$TestMode,
    [int]$Port = 8000
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

$scriptDir  = Split-Path $PSCommandPath -Parent
$script:hubScriptPath = if ($PSCommandPath) { $PSCommandPath } else { $MyInvocation.MyCommand.Path }
Push-Location $scriptDir

# --- Load Config ---
$script:configFile = if ($TestMode) { "config.test.json" } else { "config.json" }
$configPath = Join-Path $scriptDir $script:configFile
$script:indexDir             = "C:\Users\dlebel\Documents\PDFIndex"
$emailIntervalMinutes = 30
if (Test-Path $configPath) {
    try {
        $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
        if ($cfg.indexFolder)                { $script:indexDir             = $cfg.indexFolder }
        if ($cfg.emailCheckIntervalMinutes)  { $emailIntervalMinutes = [int]$cfg.emailCheckIntervalMinutes }
    } catch {
        Write-Host "[Hub] WARNING: Failed to load config from $configPath : $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "[Hub] WARNING: Config file not found at $configPath - using defaults" -ForegroundColor Yellow
}

# --- Hub Logging ---
$script:hubLogFile = Join-Path $script:indexDir ("hub_service_{0}.log" -f (Get-Date -Format "yyyy-MM-dd"))
$script:hubSettingsPath = Join-Path $script:indexDir "hub_settings.json"
function Write-HubLog {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$ts [$Level] $Message"
    try { Add-Content -Path $script:hubLogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
    $fg = switch ($Level) {
        "ERROR" { "Red" }
        "WARN"  { "Yellow" }
        default { "DarkGray" }
    }
    Write-Host $line -ForegroundColor $fg
}

function Get-HubDispatchMode {
    param([bool]$IsTestMode)
    $defaultMode = if ($IsTestMode) { "Manual" } else { "Auto" }
    if (Test-Path $script:hubSettingsPath) {
        try {
            $settings = Get-Content $script:hubSettingsPath -Raw | ConvertFrom-Json
            $mode = [string]$settings.DispatchMode
            if (@("Auto","Manual","Hold") -contains $mode) { return $mode }
        } catch {}
    }
    return $defaultMode
}

function Save-HubDispatchMode {
    param([string]$Mode)
    try {
        @{ DispatchMode = $Mode } | ConvertTo-Json | Set-Content -Path $script:hubSettingsPath -Encoding UTF8 -Force
    } catch {
        Write-HubLog "Failed to persist dispatch mode: $($_.Exception.Message)" "WARN"
    }
}

function Send-OutlookDraftByEntryId {
    param([string]$EntryId)
    if ([string]::IsNullOrWhiteSpace($EntryId)) { throw "Draft entry id is blank." }
    $outlook = $null
    $ns = $null
    $item = $null
    try {
        try {
            $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        } catch {
            $outlook = New-Object -ComObject Outlook.Application
        }
        $ns = $outlook.GetNamespace("MAPI")
        $item = $ns.GetItemFromID($EntryId)
        if ($null -eq $item) { throw "Draft not found in Outlook Drafts." }
        $item.Send()
    } finally {
        foreach ($obj in @($item, $ns, $outlook)) {
            if ($null -ne $obj) {
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch {}
            }
        }
    }
}

# --- Restore last known state from disk ---
$persistedLastRebuild = "Never"
$rebuildSummaryPath   = Join-Path $script:indexDir "last_rebuild.json"
if (Test-Path $rebuildSummaryPath) {
    try {
        $rs = Get-Content $rebuildSummaryPath -Raw | ConvertFrom-Json
        if ($rs.Timestamp) { $persistedLastRebuild = $rs.Timestamp }
    } catch {
        Write-Host "[Hub] WARNING: Failed to read last rebuild state: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# --- Synchronized State ---
$Global:HubState = [hashtable]::Synchronized(@{
    Status                = "Starting"
    LastEmailCheck        = "Never"
    LastIndexRebuild      = $persistedLastRebuild
    ActiveTask            = "None"
    PendingTask           = $null
    ActiveJobId           = $null
    ActiveJobName         = $null
    UpTime                = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    TestMode              = $TestMode.IsPresent
    Mode                  = if ($TestMode) { "TEST" } else { "Production" }
    Port                  = $Port
    IndexFolder           = $script:indexDir
    EmailIntervalMinutes  = $emailIntervalMinutes
    NextEmailCheck        = "Pending"
    DispatchMode          = (Get-HubDispatchMode -IsTestMode $TestMode.IsPresent)
    Errors                = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
})

# --- API Server Script ---
$apiScript = {
    param($SharedState, $Dir)

    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$($SharedState.Port)/")

    try {
        $listener.Start()
        $SharedState.Status = "Running"
    } catch {
        $SharedState.Status = "Error: Port $($SharedState.Port) / Access Denied"
        return
    }

    while ($listener.IsListening) {
        # Check for shutdown request between requests
        if ($SharedState.Status -eq "Stopping") {
            try { $listener.Stop() } catch {}
            break
        }

        # Use async GetContext with timeout so we can check for shutdown
        $asyncResult = $listener.BeginGetContext($null, $null)
        while (-not $asyncResult.AsyncWaitHandle.WaitOne(500)) {
            if ($SharedState.Status -eq "Stopping") {
                try { $listener.Stop() } catch {}
                break
            }
        }
        if ($SharedState.Status -eq "Stopping") { break }

        try {
            $context = $listener.EndGetContext($asyncResult)
        } catch { break }

        $request  = $context.Request
        $response = $context.Response
        $path     = $request.Url.AbsolutePath

        $response.AddHeader("Access-Control-Allow-Origin",  "*")
        $response.AddHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        $response.AddHeader("Access-Control-Allow-Headers", "Content-Type")

        if ($request.HttpMethod -eq "OPTIONS") {
            $response.StatusCode = 200
            $response.Close()
            continue
        }

        if ($path -eq "/" -or $path -eq "/index.html") {
            $htmlPath = Join-Path $Dir "index.html"
            if (Test-Path $htmlPath) {
                $buffer = [System.Text.Encoding]::UTF8.GetBytes((Get-Content $htmlPath -Raw))
                $response.ContentType      = "text/html; charset=utf-8"
                $response.ContentLength64  = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
            }
            $response.Close()
            continue
        }

        $responseData = @{ error = "Not Found" }
        $statusCode   = 404

        if ($path -eq "/status") {
            $progress = $null
            if ($SharedState.ActiveTask -ne "None") {
                $pPath = Join-Path $SharedState.IndexFolder "progress.json"
                if (Test-Path $pPath) {
                    try {
                        $fs = [System.IO.File]::Open($pPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                        $sr = [System.IO.StreamReader]::new($fs)
                        $progress = $sr.ReadToEnd() | ConvertFrom-Json
                        $sr.Dispose(); $fs.Dispose()
                    } catch {}
                }
            }

            $emailSummary = $null
            $esPath = Join-Path $SharedState.IndexFolder "last_email_summary.json"
            if (Test-Path $esPath) {
                try {
                    $fs = [System.IO.File]::Open($esPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                    $sr = [System.IO.StreamReader]::new($fs)
                    $emailSummary = $sr.ReadToEnd() | ConvertFrom-Json
                    $sr.Dispose(); $fs.Dispose()
                } catch {}
            }

            $history = @()
            $hPath = Join-Path $SharedState.IndexFolder "transmittal_history.json"
            if (Test-Path $hPath) {
                try {
                    $fs = [System.IO.File]::Open($hPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                    $sr = [System.IO.StreamReader]::new($fs)
                    $history = $sr.ReadToEnd() | ConvertFrom-Json
                    $sr.Dispose(); $fs.Dispose()
                } catch {}
            }

            $emailProgress = $null
            $epPath = Join-Path $SharedState.IndexFolder "email_progress.json"
            if (Test-Path $epPath) {
                try {
                    $fs = [System.IO.File]::Open($epPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                    $sr = [System.IO.StreamReader]::new($fs)
                    $emailProgress = $sr.ReadToEnd() | ConvertFrom-Json
                    $sr.Dispose(); $fs.Dispose()
                } catch {}
            }

            $recentErrors = @()
            if ($SharedState.Errors -and $SharedState.Errors.Count -gt 0) {
                $recentErrors = @($SharedState.Errors | Select-Object -Last 10)
            }

            $responseData = @{
                Status               = $SharedState.Status
                LastEmailCheck       = $SharedState.LastEmailCheck
                LastIndexRebuild     = $SharedState.LastIndexRebuild
                ActiveTask           = $SharedState.ActiveTask
                UpTime               = $SharedState.UpTime
                TestMode             = $SharedState.TestMode
                Mode                 = $SharedState.Mode
                DispatchMode         = $SharedState.DispatchMode
                Port                 = $SharedState.Port
                EmailIntervalMinutes = $SharedState.EmailIntervalMinutes
                NextEmailCheck       = $SharedState.NextEmailCheck
                Progress             = $progress
                LastEmailSummary     = $emailSummary
                History              = $history
                EmailProgress        = $emailProgress
                RecentErrors         = $recentErrors
            }
            $statusCode = 200
        }
        else {
            switch ($path) {
                "/trigger-crawl" {
                    if ($SharedState.ActiveTask -ne "None") {
                        $responseData = @{ error = "Busy"; message = "Another task is already running: $($SharedState.ActiveTask)" }; $statusCode = 409
                    } else {
                        $SharedState.ActiveTask  = "Index Rebuild"
                        $SharedState.PendingTask = "crawl"
                        Write-HubLog "Received API request: trigger-crawl"
                        $responseData = @{ message = "Crawl Started" }; $statusCode = 202
                    }
                }
                "/check-emails" {
                    if ($SharedState.ActiveTask -ne "None") {
                        $responseData = @{ error = "Busy"; message = "Another task is already running: $($SharedState.ActiveTask)" }; $statusCode = 409
                    } else {
                        $SharedState.ActiveTask  = "Email Check"
                        $SharedState.PendingTask = "email"
                        $SharedState.NextEmailCheck = "Running now..."
                        Write-HubLog "Received API request: check-emails"
                        $responseData = @{ message = "Email Check Started" }; $statusCode = 202
                    }
                }
                "/set-dispatch-mode" {
                    $mode = [string]$request.QueryString["mode"]
                    if (@("Auto","Manual","Hold") -contains $mode) {
                        $SharedState.DispatchMode = $mode
                        Save-HubDispatchMode -Mode $mode
                        Write-HubLog "Dispatch mode changed to $mode"
                        $responseData = @{ message = "Dispatch mode set to $mode"; DispatchMode = $mode }; $statusCode = 200
                    } else {
                        $responseData = @{ error = "Bad mode"; message = "Mode must be Auto, Manual, or Hold." }; $statusCode = 400
                    }
                }
                "/send-last-draft" {
                    $esPath = Join-Path $SharedState.IndexFolder "last_email_summary.json"
                    if (-not (Test-Path $esPath)) {
                        $responseData = @{ error = "Not Found"; message = "No last summary found." }; $statusCode = 404
                    } else {
                        try {
                            $summary = Get-Content $esPath -Raw | ConvertFrom-Json
                            $entryId = [string]$summary.DraftEntryId
                            if ([string]::IsNullOrWhiteSpace($entryId)) {
                                $responseData = @{ error = "No draft"; message = "No saved draft is available for the last result." }; $statusCode = 400
                            } else {
                                Send-OutlookDraftByEntryId -EntryId $entryId
                                if ($summary.PSObject.Properties.Name -contains "DraftCreated") { $summary.DraftCreated = $false }
                                if ($summary.PSObject.Properties.Name -contains "TransmittalSent") { $summary.TransmittalSent = $true }
                                if ($summary.PSObject.Properties.Name -contains "Status") { $summary.Status = "Transmittal Sent" }
                                if ($summary.PSObject.Properties.Name -contains "DispatchState") { $summary.DispatchState = "Sent" }
                                if ($summary.PSObject.Properties.Name -contains "DispatchMode") { $summary.DispatchMode = "Auto" }
                                $summary | ConvertTo-Json -Depth 8 | Set-Content -Path $esPath -Encoding UTF8 -Force
                                Write-HubLog "Sent Outlook draft from dashboard (EntryId=$entryId)"
                                $responseData = @{ message = "Draft sent" }; $statusCode = 200
                            }
                        } catch {
                            Write-HubLog "send-last-draft failed: $($_.Exception.Message)" "ERROR"
                            $responseData = @{ error = "Send failed"; message = $_.Exception.Message }; $statusCode = 500
                        }
                    }
                }
                "/clean-index" {
                    if ($SharedState.ActiveTask -ne "None") {
                        $responseData = @{ error = "Busy"; message = "Another task is already running: $($SharedState.ActiveTask)" }; $statusCode = 409
                    } else {
                        $SharedState.ActiveTask  = "Cleaning Index"
                        $SharedState.PendingTask = "clean"
                        Write-HubLog "Received API request: clean-index"
                        $responseData = @{ message = "Index Cleaning Started" }; $statusCode = 202
                    }
                }
                "/clear-last-summary" {
                    $esPath = Join-Path $SharedState.IndexFolder "last_email_summary.json"
                    if (Test-Path $esPath) { Remove-Item $esPath -Force }
                    $responseData = @{ message = "Last summary cleared" }; $statusCode = 200
                }
                "/clear-history-all" {
                    $hPath = Join-Path $SharedState.IndexFolder "transmittal_history.json"
                    if (Test-Path $hPath) { Set-Content $hPath "[]" -Force }
                    $responseData = @{ message = "History cleared" }; $statusCode = 200
                }
                "/clear-history-item" {
                    $idxStr = $request.QueryString["index"]
                    if ($null -ne $idxStr) {
                        $targetIdx = [int]$idxStr
                        $hPath = Join-Path $SharedState.IndexFolder "transmittal_history.json"
                        if (Test-Path $hPath) {
                            $history = Get-Content $hPath -Raw | ConvertFrom-Json
                            if ($targetIdx -ge 0 -and $targetIdx -lt $history.Count) {
                                # Remove the item
                                if ($history.Count -eq 1) { $history = @() }
                                else {
                                    $newList = [System.Collections.Generic.List[object]]::new()
                                    for ($i=0; $i -lt $history.Count; $i++) {
                                        if ($i -ne $targetIdx) { $newList.Add($history[$i]) }
                                    }
                                    $history = $newList
                                }
                                $history | ConvertTo-Json -Depth 5 | Set-Content $hPath -Force
                                $responseData = @{ message = "Item cleared" }; $statusCode = 200
                            }
                        }
                    }
                }
                "/search" {
                    $q = $request.QueryString["q"]
                    if ($q) {
                        $csvPath = Join-Path $SharedState.IndexFolder "pdf_index_clean.csv"
                        if (Test-Path $csvPath) {
                            try {
                                # Use Import-Csv for proper quoted-field handling
                                $allRows = Import-Csv -Path $csvPath
                                $matched = $allRows | Where-Object {
                                    $_.BasePart -like "*$q*" -or $_.FileName -like "*$q*" -or $_.FullPath -like "*$q*"
                                } | Select-Object -First 50
                                $responseData = @{
                                    query   = $q
                                    results = @($matched | ForEach-Object {
                                        @{ PartNumber = $_.BasePart; Revision = $_.Rev; Description = $_.FileName; FullPath = $_.FullPath }
                                    })
                                }
                            } catch {
                                $responseData = @{ error = "Search failed: $($_.Exception.Message)" }
                            }
                            $statusCode = 200
                        } else {
                            $responseData = @{ error = "Index not found" }; $statusCode = 404
                        }
                    } else {
                        $responseData = @{ error = "No query" }; $statusCode = 400
                    }
                }
            }
        }

        try {
            $json   = $responseData | ConvertTo-Json -Depth 5
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
            $response.ContentLength64 = $buffer.Length
            $response.ContentType     = "application/json"
            $response.StatusCode      = $statusCode
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
        } catch {
            try { $response.StatusCode = 500 } catch {}
        } finally {
            try { $response.Close() } catch {}
        }
    }
}

# --- App Tray Icon ---
function New-AppTrayIcon {
    try {
        $bmp = New-Object System.Drawing.Bitmap 16,16
        $g   = [System.Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $g.Clear([System.Drawing.Color]::Transparent)
        $bgBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(26, 115, 232))
        $g.FillEllipse($bgBrush, 0, 0, 15, 15)
        $font = New-Object System.Drawing.Font("Segoe UI", 7.5, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
        $fg   = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
        $sf   = New-Object System.Drawing.StringFormat
        $sf.Alignment = [System.Drawing.StringAlignment]::Center
        $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
        $g.DrawString("N", $font, $fg, (New-Object System.Drawing.RectangleF(0, 0, 16, 16)), $sf)
        $g.Dispose(); $font.Dispose(); $bgBrush.Dispose(); $fg.Dispose(); $sf.Dispose()
        $icon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
        $bmp.Dispose()
        return $icon
    } catch {
        return [System.Drawing.SystemIcons]::Application
    }
}
# --- System Tray Setup ---
$notifyIcon         = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Icon    = New-AppTrayIcon
$notifyIcon.Text    = if ($TestMode) { "NMT Hub v2.10 [TEST]" } else { "NMT Drawing Hub" }
$notifyIcon.Visible = $true


$contextMenu = New-Object System.Windows.Forms.ContextMenu
$item1       = $contextMenu.MenuItems.Add("Open Dashboard")
$item1.add_Click({ Start-Process "http://localhost:$($Global:HubState.Port)" })

$contextMenu.MenuItems.Add("-") | Out-Null

$itemRestart = $contextMenu.MenuItems.Add("Restart Hub")
$itemRestart.add_Click({
    $script:pendingRestart = $true
    $Global:HubState.Status = "Stopping"
})

$itemExit = $contextMenu.MenuItems.Add("Exit Hub")
$itemExit.add_Click({
    $Global:HubState.Status = "Stopping"
})
$script:lastNotifyFolder = ""
$notifyIcon.add_BalloonTipClicked({
    if ($script:lastNotifyFolder -and (Test-Path $script:lastNotifyFolder)) {
        Start-Process "explorer.exe" $script:lastNotifyFolder
    }
})
$notifyIcon.ContextMenu = $contextMenu

# --- Main Service Function ---
function Start-HubService {
    $script:apiRunspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
    $script:apiRunspace.ApartmentState = [System.Threading.ApartmentState]::MTA
    $script:apiRunspace.Open()

    $script:apiPS          = [System.Management.Automation.PowerShell]::Create()
    $script:apiPS.Runspace = $script:apiRunspace
    $script:apiPS.AddScript($apiScript).AddArgument($Global:HubState).AddArgument($scriptDir) | Out-Null
    $script:apiHandle = $script:apiPS.BeginInvoke()

    $modeLabel = if ($TestMode) { "TEST MODE" } else { "Production" }
    Write-Host "[Hub] Service Started ($modeLabel)" -ForegroundColor Cyan

    $timeout = (Get-Date).AddSeconds(5)
    while ($Global:HubState.Status -eq "Starting" -and (Get-Date) -lt $timeout) {
        Start-Sleep -Milliseconds 100
    }

    $lastEmailCheckTime = [datetime]::MinValue

    # Clear stale per-session files so dashboard starts fresh on restart
    Remove-Item (Join-Path $script:indexDir "email_progress.json") -Force -ErrorAction SilentlyContinue

    while ($Global:HubState.Status -ne "Stopping") {
        $now = Get-Date

        # --- Auto email check timer ---
        if ($Global:HubState.Status -eq "Running" -and $Global:HubState.ActiveTask -eq "None") {
            $elapsed = ($now - $lastEmailCheckTime).TotalMinutes
            if ($elapsed -ge $emailIntervalMinutes) {
                $Global:HubState.ActiveTask  = "Email Check"
                $Global:HubState.PendingTask = "email"
                $Global:HubState.NextEmailCheck = "Running now..."
            } else {
                $secsLeft = [math]::Ceiling(($emailIntervalMinutes - $elapsed) * 60)
                $minsLeft = [math]::Floor($secsLeft / 60)
                $sLeft    = $secsLeft % 60
                $Global:HubState.NextEmailCheck = "in ${minsLeft}m ${sLeft}s"
            }
        }

        if ($Global:HubState.Status -eq "Running" -and $Global:HubState.PendingTask) {
            $task = $Global:HubState.PendingTask
            $Global:HubState.PendingTask = $null
            Write-HubLog "Dispatching pending task '$task'"
             
            $rebuildPath = Join-Path $scriptDir "IndexRebuild.ps1"
            $cleanPath   = Join-Path $scriptDir "IndexClean.ps1"
            $monitorPath = Join-Path $scriptDir "EmailOrderMonitor.ps1"
            $job = $null
            $proc = $null

            switch ($task) {
                "crawl" {
                    Start-Job -Name "Hub_Crawl" -ScriptBlock {
                        param($rebuildPath, $configFile)
                        & $rebuildPath -Config $configFile
                    } -ArgumentList $rebuildPath, $script:configFile | Out-Null
                    Write-HubLog "Dispatched job Hub_Crawl using $rebuildPath -Config $script:configFile"
                }
                "email" {
                    $tm = $Global:HubState.TestMode
                    # Clear stale email_progress.json so dashboard doesn't show old state
                    Remove-Item (Join-Path $script:indexDir "email_progress.json") -Force -ErrorAction SilentlyContinue
                    $cfgPath = Join-Path $scriptDir $script:configFile
                    try {
                        $procArgs = @("-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-File", $monitorPath, "-Config", $cfgPath, "-DispatchMode", $Global:HubState.DispatchMode)
                        if ($tm) { $procArgs += "-TestMode" }
                        $proc = Start-Process -FilePath "powershell.exe" -ArgumentList $procArgs -WindowStyle Hidden -PassThru -ErrorAction Stop
                    } catch {
                        Write-HubLog "Start-Process Hub_Email failed: $($_.Exception.Message)" "ERROR"
                    }
                }
                "clean" {
                    Start-Job -Name "Hub_Clean" -ScriptBlock {
                        param($cleanPath, $configFile)
                        & $cleanPath -Config $configFile -Clean
                    } -ArgumentList $cleanPath, $script:configFile | Out-Null
                    Write-HubLog "Dispatched job Hub_Clean using $cleanPath -Config $script:configFile -Clean"
                }
            }

            if ($job) {
                $Global:HubState.ActiveJobId   = $job.Id
                $Global:HubState.ActiveJobName = $job.Name
                Write-HubLog ("Dispatched job {0} (Id={1}) for task '{2}'" -f $job.Name, $job.Id, $task)
            } elseif ($proc) {
                $script:activeEmailProcess = $proc
                $Global:HubState.ActiveJobId   = $null
                $Global:HubState.ActiveJobName = "Hub_EmailProc"
                Write-HubLog ("Dispatched process Hub_Email (Pid={0}) for task '{1}'" -f $proc.Id, $task)
            } elseif ($task -eq "email") {
                Write-HubLog "Task '$task' failed to dispatch (no process object returned)." "ERROR"
                $Global:HubState.ActiveTask    = "None"
                $Global:HubState.ActiveJobId   = $null
                $Global:HubState.ActiveJobName = $null
            }
        }

        Get-Job | Where-Object { $_.Name -like "Hub_*" -and $_.State -ne "Running" } | ForEach-Object {
            $jobName  = $_.Name
            $jobState = $_.State

            # Capture any job output/errors before removing
            $jobErrors = @()
            try {
                if ($_.State -eq "Failed") {
                    $jobErrors = @($_.ChildJobs | ForEach-Object { $_.Error } | ForEach-Object { $_.ToString() })
                }
                Receive-Job $_ -ErrorAction SilentlyContinue | Out-Null
            } catch {}

            if ($jobState -eq "Failed") {
                $errMsg = "[$(Get-Date -Format 'HH:mm:ss')] $jobName FAILED"
                if ($jobErrors.Count -gt 0) { $errMsg += ": $($jobErrors[0])" }
                Write-Host "[Hub] $errMsg" -ForegroundColor Red
                Write-HubLog $errMsg "ERROR"
                $Global:HubState.Errors.Add($errMsg) | Out-Null
                # Cap error log at 50 entries
                while ($Global:HubState.Errors.Count -gt 50) { $Global:HubState.Errors.RemoveAt(0) }
            } else {
                Write-HubLog "$jobName completed with state $jobState"
            }

            if ($jobName -eq "Hub_Crawl") {
                $Global:HubState.LastIndexRebuild = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                # Clear stale progress.json so it can't appear during unrelated tasks
                Remove-Item (Join-Path $script:indexDir "progress.json") -Force -ErrorAction SilentlyContinue
            }
            if ($jobName -eq "Hub_Clean") {
                Remove-Item (Join-Path $script:indexDir "progress.json") -Force -ErrorAction SilentlyContinue
            }
            if ($jobName -eq "Hub_Email") {
                $Global:HubState.LastEmailCheck = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                $lastEmailCheckTime = Get-Date
                $Global:HubState.NextEmailCheck = "in ${emailIntervalMinutes}m 0s"
            }
            $Global:HubState.ActiveTask = "None"
            $Global:HubState.ActiveJobId = $null
            $Global:HubState.ActiveJobName = $null
            Remove-Job $_ -Force
        }

        if ($script:activeEmailProcess) {
            $procExited = $false
            $procId = $null
            $procExitCode = $null
            try { $procId = $script:activeEmailProcess.Id } catch { }
            try {
                $script:activeEmailProcess.Refresh()
                $procExited = $script:activeEmailProcess.HasExited
                if ($procExited) { $procExitCode = [int]$script:activeEmailProcess.ExitCode }
            } catch {
                if ($procId) {
                    $stillRunning = Get-Process -Id $procId -ErrorAction SilentlyContinue
                    $procExited = ($null -eq $stillRunning)
                } else {
                    $procExited = $true
                }
            }

            if ($procExited) {
                Write-HubLog ("Active email process exited (Pid={0}, ExitCode={1})" -f $procId, $procExitCode)
                $Global:HubState.LastEmailCheck = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                $lastEmailCheckTime = Get-Date
                $Global:HubState.NextEmailCheck = "in ${emailIntervalMinutes}m 0s"

                if ($null -eq $procExitCode -or $procExitCode -ne 0) {
                    $reason = if ($null -eq $procExitCode) {
                        "Email monitor process ended unexpectedly (unknown exit code)."
                    } else {
                        "Email monitor process exited with code $procExitCode. Check monitor log for details."
                    }
                    try {
                        @{
                            Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                            Step      = "failed"
                            Order     = ""
                            Detail    = $reason
                        } | ConvertTo-Json | Set-Content -Path (Join-Path $script:indexDir "email_progress.json") -Encoding UTF8 -Force
                    } catch { }
                    Write-HubLog ("Hub_Email process failure reason: " + $reason) "ERROR"
                }

                $Global:HubState.ActiveTask    = "None"
                $Global:HubState.ActiveJobId   = $null
                $Global:HubState.ActiveJobName = $null
                $script:activeEmailProcess = $null
            }
        }

        # --- Drain notification queue dropped by EmailOrderMonitor ---
        try {
            $nFiles = @(Get-Item -Path (Join-Path $script:indexDir "notify_*.json") -ErrorAction SilentlyContinue)
            foreach ($nf in ($nFiles | Sort-Object Name)) {
                try {
                    $n = Get-Content $nf.FullName -Raw | ConvertFrom-Json
                    $notifyIcon.ShowBalloonTip(6000, [string]$n.Title, [string]$n.Message,
                        [System.Windows.Forms.ToolTipIcon]::Info)
                    if ($n.FolderPath) { $script:lastNotifyFolder = [string]$n.FolderPath }
                } catch {
                    # ShowBalloonTip or JSON parse failed - still remove the file to avoid infinite retry
                } finally {
                    Remove-Item $nf.FullName -Force -ErrorAction SilentlyContinue
                }
            }
        } catch { }

        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 500
    }

    # --- Cleanup ---
    Write-Host "[Hub] Shutting down..." -ForegroundColor Yellow
    $notifyIcon.Visible = $false
    $notifyIcon.Dispose()

    # Wait briefly for the API listener to notice the Stopping status and exit cleanly
    Start-Sleep -Milliseconds 1500

    if ($script:apiHandle -and $script:apiPS) {
        try { $script:apiPS.Stop(); $script:apiPS.Dispose() } catch {}
    }
    if ($script:apiRunspace) {
        try { $script:apiRunspace.Close(); $script:apiRunspace.Dispose() } catch {}
    }
    Get-Job | Where-Object { $_.Name -like "Hub_*" } | ForEach-Object {
        try { Stop-Job $_ -ErrorAction SilentlyContinue } catch {}
        Remove-Job $_ -Force -ErrorAction SilentlyContinue
    }
    Write-Host "[Hub] Service stopped." -ForegroundColor Cyan
}

Start-HubService
if ($script:pendingRestart) {
    Start-Sleep -Seconds 2
    $extraArgs = ""
    if ($TestMode)      { $extraArgs += " -TestMode" }
    if ($Port -ne 8000) { $extraArgs += " -Port $Port" }
    try {
        $vbsPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), "hub_relaunch.vbs")
        $vbsLine1 = 'Set o = CreateObject("WScript.Shell")'
        $vbsLine2 = 'o.Run "powershell.exe -ExecutionPolicy Bypass -Sta -File " & Chr(34) & "' + $script:hubScriptPath + '" & Chr(34) & "' + $extraArgs + '", 0, False'
        [System.IO.File]::WriteAllLines($vbsPath, @($vbsLine1, $vbsLine2), [System.Text.Encoding]::ASCII)
        Start-Process "wscript.exe" -ArgumentList "`"$vbsPath`""
    } catch {
        Write-Host "[Hub] ERROR: Failed to relaunch Hub: $($_.Exception.Message)" -ForegroundColor Red
    }
}
exit
