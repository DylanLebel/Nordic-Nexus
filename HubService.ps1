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

function Resolve-ConfiguredPath {
    param(
        [string]$PathValue,
        [string]$BasePath,
        [string]$DefaultValue = ""
    )

    $candidate = if (-not [string]::IsNullOrWhiteSpace($PathValue)) { $PathValue } else { $DefaultValue }
    if ([string]::IsNullOrWhiteSpace($candidate)) { return "" }
    if ([System.IO.Path]::IsPathRooted($candidate)) { return [System.IO.Path]::GetFullPath($candidate) }
    if ([string]::IsNullOrWhiteSpace($BasePath)) { return [System.IO.Path]::GetFullPath($candidate) }
    return [System.IO.Path]::GetFullPath((Join-Path $BasePath $candidate))
}

function Ensure-DirectoryExists {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path $Path)) {
        [void][System.IO.Directory]::CreateDirectory($Path)
    }
}

# --- Load Config ---
$script:configFile = if ($TestMode) { "config.test.json" } else { "config.json" }
$configPath = Join-Path $scriptDir $script:configFile
$script:indexDir             = "C:\Users\dlebel\Documents\PDFIndex"
$emailIntervalMinutes = 30
if (Test-Path $configPath) {
    try {
        $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
        $configBaseDir = Split-Path $configPath -Parent
        if ($cfg.indexFolder)                { $script:indexDir             = Resolve-ConfiguredPath -PathValue $cfg.indexFolder -BasePath $configBaseDir -DefaultValue $script:indexDir }
        if ($cfg.emailCheckIntervalMinutes)  { $emailIntervalMinutes = [int]$cfg.emailCheckIntervalMinutes }
    } catch {
        Write-Host "[Hub] WARNING: Failed to load config from $configPath : $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "[Hub] WARNING: Config file not found at $configPath - using defaults" -ForegroundColor Yellow
}
Ensure-DirectoryExists -Path $script:indexDir

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

function Get-HubSettings {
    if (Test-Path $script:hubSettingsPath) {
        try {
            $raw = Get-Content $script:hubSettingsPath -Raw
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                return ($raw | ConvertFrom-Json)
            }
        } catch {
            Write-HubLog "Failed to read hub settings: $($_.Exception.Message)" "WARN"
        }
    }
    return [pscustomobject]@{}
}

function Save-HubSettings {
    param([hashtable]$Patch)

    try {
        $current = Get-HubSettings
        $merged = [ordered]@{}
        foreach ($prop in $current.PSObject.Properties) {
            $merged[$prop.Name] = $prop.Value
        }
        foreach ($key in $Patch.Keys) {
            $merged[$key] = $Patch[$key]
        }
        $merged | ConvertTo-Json | Set-Content -Path $script:hubSettingsPath -Encoding UTF8 -Force
    } catch {
        Write-HubLog "Failed to persist hub settings: $($_.Exception.Message)" "WARN"
    }
}

function Get-HubDispatchMode {
    param([bool]$IsTestMode)
    $defaultMode = if ($IsTestMode) { "Manual" } else { "Auto" }
    $settings = Get-HubSettings
    $mode = [string]$settings.DispatchMode
    if (@("Auto","Manual","Hold") -contains $mode) { return $mode }
    return $defaultMode
}

function Save-HubDispatchMode {
    param([string]$Mode)
    Save-HubSettings @{ DispatchMode = $Mode }
}

function Get-HubEmailInterval {
    param([int]$DefaultValue)

    $settings = Get-HubSettings
    $parsed = 0
    if ([int]::TryParse([string]$settings.EmailIntervalMinutes, [ref]$parsed) -and $parsed -ge 1 -and $parsed -le 120) {
        return $parsed
    }
    return $DefaultValue
}

function Save-HubEmailInterval {
    param([int]$Minutes)
    Save-HubSettings @{ EmailIntervalMinutes = $Minutes }
}

$emailIntervalMinutes = Get-HubEmailInterval -DefaultValue $emailIntervalMinutes

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
    ActiveTaskStartedAt   = $null
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
    ReplayRequest         = $null
    Errors                = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
})

# --- API Server Script ---
$apiScript = {
    param($SharedState, $Dir)

    function Read-SharedJsonFile {
        param([string]$Path)
        if (-not (Test-Path $Path)) { return $null }
        $fs = $null
        $sr = $null
        try {
            $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $sr = [System.IO.StreamReader]::new($fs)
            return ($sr.ReadToEnd() | ConvertFrom-Json)
        } catch {
            return $null
        } finally {
            if ($sr) { try { $sr.Dispose() } catch {} }
            if ($fs) { try { $fs.Dispose() } catch {} }
        }
    }

    function Get-HistoryEntriesForApi {
        param($SharedState)
        $hPath = Join-Path $SharedState.IndexFolder "transmittal_history.json"
        $history = Read-SharedJsonFile -Path $hPath
        return @($history)
    }

    function New-ReplayRequestFromHistory {
        param(
            $SharedState,
            [string]$Dir,
            [int]$HistoryIndex,
            [string]$ModeLabel
        )

        $history = Get-HistoryEntriesForApi -SharedState $SharedState
        if ($HistoryIndex -lt 0 -or $HistoryIndex -ge $history.Count) {
            throw "History item not found."
        }

        $entry = $history[$HistoryIndex]
        $outputFolder = [string]$entry.OutputFolder
        if ([string]::IsNullOrWhiteSpace($outputFolder) -or -not (Test-Path $outputFolder)) {
            throw "History item does not have a valid output folder."
        }

        $bomFile = Join-Path $outputFolder "order_bom.txt"
        if (-not (Test-Path $bomFile)) {
            throw "Replay BOM not found for this history item."
        }

        $isTest = $false
        try { $isTest = [bool]$entry.TestMode } catch { $isTest = [bool]$SharedState.TestMode }
        $configFileName = if ($isTest) { "config.test.json" } else { "config.json" }
        $configPath = Join-Path $Dir $configFileName

        $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $safeMode = ($ModeLabel -replace '[^\w-]', '_').ToLowerInvariant()
        $replayRoot = Join-Path $outputFolder "_reruns"
        $targetOutputFolder = Join-Path $replayRoot ("{0}_{1}" -f $safeMode, $stamp)

        return [ordered]@{
            HistoryIndex      = $HistoryIndex
            ModeLabel         = $ModeLabel
            BomFile           = $bomFile
            ConfigFileName    = $configFileName
            ConfigPath        = $configPath
            SourceOutputFolder = $outputFolder
            TargetOutputFolder = $targetOutputFolder
            JobNumber         = [string]$entry.JobNumber
        }
    }

    function Open-SharedPath {
        param([string]$TargetPath)
        if ([string]::IsNullOrWhiteSpace($TargetPath)) { throw "Path is blank." }
        if (-not (Test-Path $TargetPath)) { throw "Path not found: $TargetPath" }
        Start-Process -FilePath $TargetPath | Out-Null
    }

    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$($SharedState.Port)/")
    try { $listener.Prefixes.Add("http://127.0.0.1:$($SharedState.Port)/") } catch {}
    try { $listener.Prefixes.Add("http://[::1]:$($SharedState.Port)/") } catch {}

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
                $htmlText = [System.IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($htmlText)
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
            $history = @($history)

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
                ActiveTaskStartedAt  = $SharedState.ActiveTaskStartedAt
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
                "/open-path" {
                    $targetPath = [string]$request.QueryString["path"]
                    try {
                        Open-SharedPath -TargetPath $targetPath
                        $responseData = @{ message = "Opened path"; Path = $targetPath }; $statusCode = 200
                    } catch {
                        $responseData = @{ error = "Open failed"; message = $_.Exception.Message }; $statusCode = 400
                    }
                }
                "/replay-history-item" {
                    if ($SharedState.ActiveTask -ne "None") {
                        $responseData = @{ error = "Busy"; message = "Another task is already running: $($SharedState.ActiveTask)" }; $statusCode = 409
                    } else {
                        try {
                            $targetIdx = [int]$request.QueryString["index"]
                            $replayRequest = New-ReplayRequestFromHistory -SharedState $SharedState -Dir $Dir -HistoryIndex $targetIdx -ModeLabel "Replay"
                            $SharedState.ReplayRequest = $replayRequest
                            $SharedState.ActiveTask = "Replay Order"
                            $SharedState.PendingTask = "replay"
                            $responseData = @{ message = "Replay started"; OutputFolder = $replayRequest.TargetOutputFolder; BomFile = $replayRequest.BomFile }; $statusCode = 202
                        } catch {
                            $responseData = @{ error = "Replay failed"; message = $_.Exception.Message }; $statusCode = 400
                        }
                    }
                }
                "/recheck-history-item" {
                    if ($SharedState.ActiveTask -ne "None") {
                        $responseData = @{ error = "Busy"; message = "Another task is already running: $($SharedState.ActiveTask)" }; $statusCode = 409
                    } else {
                        try {
                            $targetIdx = [int]$request.QueryString["index"]
                            $replayRequest = New-ReplayRequestFromHistory -SharedState $SharedState -Dir $Dir -HistoryIndex $targetIdx -ModeLabel "Recheck"
                            $SharedState.ReplayRequest = $replayRequest
                            $SharedState.ActiveTask = "Recheck Order"
                            $SharedState.PendingTask = "recheck"
                            $responseData = @{ message = "Recheck started"; OutputFolder = $replayRequest.TargetOutputFolder; BomFile = $replayRequest.BomFile }; $statusCode = 202
                        } catch {
                            $responseData = @{ error = "Recheck failed"; message = $_.Exception.Message }; $statusCode = 400
                        }
                    }
                }
                "/rebuild-and-replay-history-item" {
                    if ($SharedState.ActiveTask -ne "None") {
                        $responseData = @{ error = "Busy"; message = "Another task is already running: $($SharedState.ActiveTask)" }; $statusCode = 409
                    } else {
                        try {
                            $targetIdx = [int]$request.QueryString["index"]
                            $replayRequest = New-ReplayRequestFromHistory -SharedState $SharedState -Dir $Dir -HistoryIndex $targetIdx -ModeLabel "RebuildReplay"
                            $SharedState.ReplayRequest = $replayRequest
                            $SharedState.ActiveTask = "Rebuild + Replay"
                            $SharedState.PendingTask = "rebuild-replay"
                            $responseData = @{ message = "Rebuild + replay started"; OutputFolder = $replayRequest.TargetOutputFolder; BomFile = $replayRequest.BomFile }; $statusCode = 202
                        } catch {
                            $responseData = @{ error = "Rebuild + replay failed"; message = $_.Exception.Message }; $statusCode = 400
                        }
                    }
                }
                "/set-email-interval" {
                    $mins = [string]$request.QueryString["minutes"]
                    $parsed = 0
                    if ([int]::TryParse($mins, [ref]$parsed) -and $parsed -ge 1 -and $parsed -le 120) {
                        $SharedState.EmailIntervalMinutes = $parsed
                        Save-HubEmailInterval -Minutes $parsed
                        Write-HubLog "Email check interval changed to ${parsed}m"
                        $responseData = @{ message = "Email check interval set to ${parsed} minutes"; EmailIntervalMinutes = $parsed }; $statusCode = 200
                    } else {
                        $responseData = @{ error = "Bad value"; message = "Interval must be between 1 and 120 minutes." }; $statusCode = 400
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

                                $hPath = Join-Path $SharedState.IndexFolder "transmittal_history.json"
                                if (Test-Path $hPath) {
                                    try {
                                        $historyParsed = Get-Content $hPath -Raw | ConvertFrom-Json
                                        $historyList = [System.Collections.Generic.List[object]]::new()
                                        if ($historyParsed -is [System.Array]) {
                                            foreach ($item in $historyParsed) { $historyList.Add($item) }
                                        } elseif ($null -ne $historyParsed) {
                                            $historyList.Add($historyParsed)
                                        }

                                        foreach ($item in $historyList) {
                                            if ($null -eq $item) { continue }
                                            $sameDraft = ([string]$item.DraftEntryId -eq $entryId)
                                            $sameRun = ([string]$item.JobNumber -eq [string]$summary.JobNumber) -and ([string]$item.TransmittalNo -eq [string]$summary.TransmittalNo)
                                            if (-not $sameDraft -and -not $sameRun) { continue }
                                            if ($item.PSObject.Properties.Name -contains "DraftCreated") { $item.DraftCreated = $false }
                                            if ($item.PSObject.Properties.Name -contains "TransmittalSent") { $item.TransmittalSent = $true }
                                            if ($item.PSObject.Properties.Name -contains "Status") { $item.Status = "Transmittal Sent" }
                                            if ($item.PSObject.Properties.Name -contains "DispatchState") { $item.DispatchState = "Sent" }
                                            if ($item.PSObject.Properties.Name -contains "DispatchMode") { $item.DispatchMode = "Auto" }
                                            break
                                        }

                                        @($historyList) | ConvertTo-Json -Depth 8 | Set-Content -Path $hPath -Encoding UTF8 -Force
                                    } catch {
                                        Write-HubLog "send-last-draft history sync failed: $($_.Exception.Message)" "WARN"
                                    }
                                }

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

. (Join-Path $scriptDir "HubApiScript.ps1")
$apiScript = $script:HubApiScript

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
            $currentInterval = $Global:HubState.EmailIntervalMinutes
            if ($elapsed -ge $currentInterval) {
                $Global:HubState.ActiveTask  = "Email Check"
                $Global:HubState.PendingTask = "email"
                $Global:HubState.NextEmailCheck = "Running now..."
            } else {
                $secsLeft = [math]::Ceiling(($currentInterval - $elapsed) * 60)
                $minsLeft = [math]::Floor($secsLeft / 60)
                $sLeft    = $secsLeft % 60
                $Global:HubState.NextEmailCheck = "in ${minsLeft}m ${sLeft}s"
            }
        }

        if ($Global:HubState.Status -eq "Running" -and $Global:HubState.PendingTask) {
            $task = $Global:HubState.PendingTask
            $Global:HubState.PendingTask = $null
            if ($Global:HubState.ActiveTask -ne "None") {
                $Global:HubState.ActiveTaskStartedAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            }
            Write-HubLog "Dispatching pending task '$task'"
             
            $rebuildPath = Join-Path $scriptDir "IndexRebuild.ps1"
            $cleanPath   = Join-Path $scriptDir "IndexClean.ps1"
            $monitorPath = Join-Path $scriptDir "EmailOrderMonitor.ps1"
            $replayPath  = Join-Path $scriptDir "Replay-OrderRun.ps1"
            $progressPath = Join-Path $script:indexDir "progress.json"
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
                        $proc = Start-Process -FilePath "powershell.exe" -ArgumentList $procArgs -WorkingDirectory $scriptDir -WindowStyle Hidden -PassThru -ErrorAction Stop
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
                "replay" {
                    $req = $Global:HubState.ReplayRequest
                    if ($req) {
                        $job = Start-Job -Name "Hub_Replay" -ScriptBlock {
                            param($replayPath, $bomFile, $configPath, $outputFolder, $progressFile, $jobNumber)
                            try {
                                @{
                                    Message   = ("Replaying job {0}..." -f $jobNumber)
                                    Count     = 0
                                    Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                                } | ConvertTo-Json | Set-Content -Path $progressFile -Encoding UTF8 -Force -ErrorAction SilentlyContinue
                            } catch {}
                            & $replayPath -BomFile $bomFile -ConfigPath $configPath -OutputFolder $outputFolder -Quiet -HubProgressFile $progressFile
                        } -ArgumentList $replayPath, $req.BomFile, $req.ConfigPath, $req.TargetOutputFolder, $progressPath, $req.JobNumber
                        Write-HubLog "Dispatched replay for job $($req.JobNumber) -> $($req.TargetOutputFolder)"
                    } else {
                        Write-HubLog "Replay request missing from shared state." "ERROR"
                    }
                }
                "recheck" {
                    $req = $Global:HubState.ReplayRequest
                    if ($req) {
                        $job = Start-Job -Name "Hub_Recheck" -ScriptBlock {
                            param($replayPath, $bomFile, $configPath, $outputFolder, $progressFile, $jobNumber)
                            try {
                                @{
                                    Message   = ("Rechecking job {0}..." -f $jobNumber)
                                    Count     = 0
                                    Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                                } | ConvertTo-Json | Set-Content -Path $progressFile -Encoding UTF8 -Force -ErrorAction SilentlyContinue
                            } catch {}
                            & $replayPath -BomFile $bomFile -ConfigPath $configPath -OutputFolder $outputFolder -Quiet -HubProgressFile $progressFile
                        } -ArgumentList $replayPath, $req.BomFile, $req.ConfigPath, $req.TargetOutputFolder, $progressPath, $req.JobNumber
                        Write-HubLog "Dispatched recheck for job $($req.JobNumber) -> $($req.TargetOutputFolder)"
                    } else {
                        Write-HubLog "Recheck request missing from shared state." "ERROR"
                    }
                }
                "rebuild-replay" {
                    $req = $Global:HubState.ReplayRequest
                    if ($req) {
                        $job = Start-Job -Name "Hub_RebuildReplay" -ScriptBlock {
                            param($rebuildPath, $configFile, $replayPath, $bomFile, $configPath, $outputFolder, $progressFile, $jobNumber)
                            try {
                                @{
                                    Message   = ("Step 1 of 2: rebuilding index for job {0}..." -f $jobNumber)
                                    Count     = 0
                                    Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                                } | ConvertTo-Json | Set-Content -Path $progressFile -Encoding UTF8 -Force -ErrorAction SilentlyContinue
                            } catch {}
                            & $rebuildPath -Config $configFile
                            try {
                                @{
                                    Message   = ("Step 2 of 2: replaying job {0}..." -f $jobNumber)
                                    Count     = 0
                                    Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                                } | ConvertTo-Json | Set-Content -Path $progressFile -Encoding UTF8 -Force -ErrorAction SilentlyContinue
                            } catch {}
                            & $replayPath -BomFile $bomFile -ConfigPath $configPath -OutputFolder $outputFolder -Quiet -HubProgressFile $progressFile
                        } -ArgumentList $rebuildPath, $script:configFile, $replayPath, $req.BomFile, $req.ConfigPath, $req.TargetOutputFolder, $progressPath, $req.JobNumber
                        Write-HubLog "Dispatched rebuild + replay for job $($req.JobNumber) -> $($req.TargetOutputFolder)"
                    } else {
                        Write-HubLog "Rebuild + replay request missing from shared state." "ERROR"
                    }
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
            } elseif ($task -in @("email","replay","recheck","rebuild-replay")) {
                Write-HubLog "Task '$task' failed to dispatch (no process object returned)." "ERROR"
                $Global:HubState.ActiveTask    = "None"
                $Global:HubState.ActiveTaskStartedAt = $null
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

            if ($jobName -in @("Hub_Crawl","Hub_RebuildReplay")) {
                try {
                    $rebuildSummaryPath = Join-Path $script:indexDir "last_rebuild.json"
                    if (Test-Path $rebuildSummaryPath) {
                        $rs = Get-Content $rebuildSummaryPath -Raw | ConvertFrom-Json
                        if ($rs.Timestamp) { $Global:HubState.LastIndexRebuild = [string]$rs.Timestamp }
                    }
                } catch {}
                Remove-Item (Join-Path $script:indexDir "progress.json") -Force -ErrorAction SilentlyContinue
            }
            if ($jobName -in @("Hub_Clean","Hub_Replay","Hub_Recheck")) {
                Remove-Item (Join-Path $script:indexDir "progress.json") -Force -ErrorAction SilentlyContinue
            }
            if ($jobName -eq "Hub_Email") {
                $Global:HubState.LastEmailCheck = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                $lastEmailCheckTime = Get-Date
                $ei = $Global:HubState.EmailIntervalMinutes
                $Global:HubState.NextEmailCheck = "in ${ei}m 0s"
            }
            if ($jobName -in @("Hub_Replay","Hub_Recheck","Hub_RebuildReplay")) {
                $Global:HubState.ReplayRequest = $null
            }
            $Global:HubState.ActiveTask = "None"
            $Global:HubState.ActiveTaskStartedAt = $null
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
                $ei2 = $Global:HubState.EmailIntervalMinutes
                $Global:HubState.NextEmailCheck = "in ${ei2}m 0s"

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
                $Global:HubState.ActiveTaskStartedAt = $null
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
