$script:HubApiScript = {
    param($SharedState, $Dir)

    function Write-ApiLog {
        param(
            [string]$Message,
            [string]$Level = "INFO"
        )

        $logDir = [string]$SharedState.IndexFolder
        if ([string]::IsNullOrWhiteSpace($logDir)) { return }

        try {
            if (-not (Test-Path $logDir)) {
                [void][System.IO.Directory]::CreateDirectory($logDir)
            }

            $logPath = Join-Path $logDir ("hub_service_{0}.log" -f (Get-Date -Format "yyyy-MM-dd"))
            $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $line = "$ts [$Level] $Message"
            Add-Content -Path $logPath -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue

            $fg = switch ($Level) {
                "ERROR" { "Red" }
                "WARN"  { "Yellow" }
                default { "DarkGray" }
            }
            Write-Host $line -ForegroundColor $fg
        } catch {}
    }

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

    function Save-ApiSetting {
        param([hashtable]$Patch)

        $settingsDir = [string]$SharedState.IndexFolder
        $settingsPath = Join-Path $settingsDir "hub_settings.json"
        try {
            if (-not (Test-Path $settingsDir)) {
                [void][System.IO.Directory]::CreateDirectory($settingsDir)
            }

            $merged = [ordered]@{}
            if (Test-Path $settingsPath) {
                $raw = Get-Content $settingsPath -Raw
                if (-not [string]::IsNullOrWhiteSpace($raw)) {
                    $current = $raw | ConvertFrom-Json
                    foreach ($prop in $current.PSObject.Properties) {
                        $merged[$prop.Name] = $prop.Value
                    }
                }
            }

            foreach ($key in $Patch.Keys) {
                $merged[$key] = $Patch[$key]
            }

            $merged | ConvertTo-Json | Set-Content -Path $settingsPath -Encoding UTF8 -Force
        } catch {
            Write-ApiLog "Failed to persist hub settings from API: $($_.Exception.Message)" "WARN"
        }
    }

    function Save-ApiEmailInterval {
        param([int]$Minutes)
        Save-ApiSetting @{ EmailIntervalMinutes = $Minutes }
    }

    function Save-ApiDispatchMode {
        param([string]$Mode)
        Save-ApiSetting @{ DispatchMode = $Mode }
    }

    function Send-OutlookDraftByEntryIdApi {
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

    function Get-HistoryEntriesForApi {
        param($SharedState)
        $hPath = Join-Path $SharedState.IndexFolder "transmittal_history.json"
        $history = Read-SharedJsonFile -Path $hPath
        if ($null -eq $history) { return @() }
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
            HistoryIndex       = $HistoryIndex
            ModeLabel          = $ModeLabel
            BomFile            = $bomFile
            ConfigFileName     = $configFileName
            ConfigPath         = $configPath
            SourceOutputFolder = $outputFolder
            TargetOutputFolder = $targetOutputFolder
            JobNumber          = [string]$entry.JobNumber
        }
    }

    function Open-SharedPath {
        param([string]$TargetPath)
        if ([string]::IsNullOrWhiteSpace($TargetPath)) { throw "Path is blank." }
        if (-not (Test-Path $TargetPath)) { throw "Path not found: $TargetPath" }
        Start-Process -FilePath $TargetPath | Out-Null
    }

    function Get-QueryValue {
        param(
            [hashtable]$QueryParams,
            [string]$Name
        )

        if ($null -eq $QueryParams) { return $null }
        if ($QueryParams.ContainsKey($Name)) { return [string]$QueryParams[$Name] }
        return $null
    }

    function Parse-QueryString {
        param([string]$RawQuery)

        $values = @{}
        if ([string]::IsNullOrWhiteSpace($RawQuery)) { return $values }

        foreach ($pair in ($RawQuery -split '&')) {
            if ([string]::IsNullOrWhiteSpace($pair)) { continue }
            $kv = $pair -split '=', 2
            $key = [System.Uri]::UnescapeDataString(($kv[0] -replace '\+', ' '))
            $value = if ($kv.Count -gt 1) {
                [System.Uri]::UnescapeDataString(($kv[1] -replace '\+', ' '))
            } else {
                ""
            }
            $values[$key] = $value
        }

        return $values
    }

    function Get-StatusDescription {
        param([int]$StatusCode)

        switch ($StatusCode) {
            200 { return "OK" }
            202 { return "Accepted" }
            400 { return "Bad Request" }
            404 { return "Not Found" }
            409 { return "Conflict" }
            500 { return "Internal Server Error" }
            default { return "OK" }
        }
    }

    function New-ByteResponse {
        param(
            [int]$StatusCode,
            [string]$ContentType,
            [byte[]]$BodyBytes
        )

        if ($null -eq $BodyBytes) { $BodyBytes = [byte[]]@() }

        return [pscustomobject]@{
            StatusCode  = $StatusCode
            ContentType = $ContentType
            BodyBytes   = $BodyBytes
        }
    }

    function New-JsonResponse {
        param(
            [object]$Body,
            [int]$StatusCode = 200
        )

        $json = $Body | ConvertTo-Json -Depth 5
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
        return New-ByteResponse -StatusCode $StatusCode -ContentType "application/json; charset=utf-8" -BodyBytes $bytes
    }

    function Invoke-HubApiRequest {
        param(
            [string]$Method,
            [string]$Path,
            [hashtable]$QueryParams,
            $SharedState,
            [string]$Dir
        )

        if ([string]::IsNullOrWhiteSpace($Method)) { $Method = "GET" }
        if ([string]::IsNullOrWhiteSpace($Path)) { $Path = "/" }
        $methodUpper = $Method.ToUpperInvariant()

        if ($methodUpper -eq "OPTIONS") {
            return New-ByteResponse -StatusCode 200 -ContentType "text/plain; charset=utf-8" -BodyBytes ([byte[]]@())
        }

        if ($Path -eq "/" -or $Path -eq "/index.html") {
            $htmlPath = Join-Path $Dir "index.html"
            if (-not (Test-Path $htmlPath)) {
                return New-JsonResponse -Body @{ error = "Dashboard not found" } -StatusCode 404
            }

            $htmlText = [System.IO.File]::ReadAllText($htmlPath, [System.Text.Encoding]::UTF8)
            $htmlBytes = [System.Text.Encoding]::UTF8.GetBytes($htmlText)
            return New-ByteResponse -StatusCode 200 -ContentType "text/html; charset=utf-8" -BodyBytes $htmlBytes
        }

        $responseData = @{ error = "Not Found" }
        $statusCode = 404

        if ($Path -eq "/status") {
            $progress = $null
            if ($SharedState.ActiveTask -ne "None") {
                $progress = Read-SharedJsonFile -Path (Join-Path $SharedState.IndexFolder "progress.json")
            }

            $emailSummary = Read-SharedJsonFile -Path (Join-Path $SharedState.IndexFolder "last_email_summary.json")

            $history = Read-SharedJsonFile -Path (Join-Path $SharedState.IndexFolder "transmittal_history.json")
            if ($null -eq $history) { $history = @() } else { $history = @($history) }

            $emailProgress = Read-SharedJsonFile -Path (Join-Path $SharedState.IndexFolder "email_progress.json")

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
            return New-JsonResponse -Body $responseData -StatusCode 200
        }

        switch ($Path) {
            "/trigger-crawl" {
                if ($SharedState.ActiveTask -ne "None") {
                    $responseData = @{ error = "Busy"; message = "Another task is already running: $($SharedState.ActiveTask)" }
                    $statusCode = 409
                } else {
                    $SharedState.ActiveTask = "Index Rebuild"
                    $SharedState.PendingTask = "crawl"
                    Write-ApiLog "Received API request: trigger-crawl"
                    $responseData = @{ message = "Crawl Started" }
                    $statusCode = 202
                }
            }
            "/check-emails" {
                if ($SharedState.ActiveTask -ne "None") {
                    $responseData = @{ error = "Busy"; message = "Another task is already running: $($SharedState.ActiveTask)" }
                    $statusCode = 409
                } else {
                    $SharedState.ActiveTask = "Email Check"
                    $SharedState.PendingTask = "email"
                    $SharedState.NextEmailCheck = "Running now..."
                    Write-ApiLog "Received API request: check-emails"
                    $responseData = @{ message = "Email Check Started" }
                    $statusCode = 202
                }
            }
            "/open-path" {
                $targetPath = Get-QueryValue -QueryParams $QueryParams -Name "path"
                try {
                    Open-SharedPath -TargetPath $targetPath
                    $responseData = @{ message = "Opened path"; Path = $targetPath }
                    $statusCode = 200
                } catch {
                    $responseData = @{ error = "Open failed"; message = $_.Exception.Message }
                    $statusCode = 400
                }
            }
            "/replay-history-item" {
                if ($SharedState.ActiveTask -ne "None") {
                    $responseData = @{ error = "Busy"; message = "Another task is already running: $($SharedState.ActiveTask)" }
                    $statusCode = 409
                } else {
                    try {
                        $targetIdx = [int](Get-QueryValue -QueryParams $QueryParams -Name "index")
                        $replayRequest = New-ReplayRequestFromHistory -SharedState $SharedState -Dir $Dir -HistoryIndex $targetIdx -ModeLabel "Replay"
                        $SharedState.ReplayRequest = $replayRequest
                        $SharedState.ActiveTask = "Replay Order"
                        $SharedState.PendingTask = "replay"
                        $responseData = @{ message = "Replay started"; OutputFolder = $replayRequest.TargetOutputFolder; BomFile = $replayRequest.BomFile }
                        $statusCode = 202
                    } catch {
                        $responseData = @{ error = "Replay failed"; message = $_.Exception.Message }
                        $statusCode = 400
                    }
                }
            }
            "/recheck-history-item" {
                if ($SharedState.ActiveTask -ne "None") {
                    $responseData = @{ error = "Busy"; message = "Another task is already running: $($SharedState.ActiveTask)" }
                    $statusCode = 409
                } else {
                    try {
                        $targetIdx = [int](Get-QueryValue -QueryParams $QueryParams -Name "index")
                        $replayRequest = New-ReplayRequestFromHistory -SharedState $SharedState -Dir $Dir -HistoryIndex $targetIdx -ModeLabel "Recheck"
                        $SharedState.ReplayRequest = $replayRequest
                        $SharedState.ActiveTask = "Recheck Order"
                        $SharedState.PendingTask = "recheck"
                        $responseData = @{ message = "Recheck started"; OutputFolder = $replayRequest.TargetOutputFolder; BomFile = $replayRequest.BomFile }
                        $statusCode = 202
                    } catch {
                        $responseData = @{ error = "Recheck failed"; message = $_.Exception.Message }
                        $statusCode = 400
                    }
                }
            }
            "/rebuild-and-replay-history-item" {
                if ($SharedState.ActiveTask -ne "None") {
                    $responseData = @{ error = "Busy"; message = "Another task is already running: $($SharedState.ActiveTask)" }
                    $statusCode = 409
                } else {
                    try {
                        $targetIdx = [int](Get-QueryValue -QueryParams $QueryParams -Name "index")
                        $replayRequest = New-ReplayRequestFromHistory -SharedState $SharedState -Dir $Dir -HistoryIndex $targetIdx -ModeLabel "RebuildReplay"
                        $SharedState.ReplayRequest = $replayRequest
                        $SharedState.ActiveTask = "Rebuild + Replay"
                        $SharedState.PendingTask = "rebuild-replay"
                        $responseData = @{ message = "Rebuild + replay started"; OutputFolder = $replayRequest.TargetOutputFolder; BomFile = $replayRequest.BomFile }
                        $statusCode = 202
                    } catch {
                        $responseData = @{ error = "Rebuild + replay failed"; message = $_.Exception.Message }
                        $statusCode = 400
                    }
                }
            }
            "/set-email-interval" {
                $mins = Get-QueryValue -QueryParams $QueryParams -Name "minutes"
                $parsed = 0
                if ([int]::TryParse([string]$mins, [ref]$parsed) -and $parsed -ge 1 -and $parsed -le 120) {
                    $SharedState.EmailIntervalMinutes = $parsed
                    Save-ApiEmailInterval -Minutes $parsed
                    Write-ApiLog "Email check interval changed to ${parsed}m"
                    $responseData = @{ message = "Email check interval set to ${parsed} minutes"; EmailIntervalMinutes = $parsed }
                    $statusCode = 200
                } else {
                    $responseData = @{ error = "Bad value"; message = "Interval must be between 1 and 120 minutes." }
                    $statusCode = 400
                }
            }
            "/set-dispatch-mode" {
                $mode = Get-QueryValue -QueryParams $QueryParams -Name "mode"
                if (@("Auto","Manual","Hold") -contains $mode) {
                    $SharedState.DispatchMode = $mode
                    Save-ApiDispatchMode -Mode $mode
                    Write-ApiLog "Dispatch mode changed to $mode"
                    $responseData = @{ message = "Dispatch mode set to $mode"; DispatchMode = $mode }
                    $statusCode = 200
                } else {
                    $responseData = @{ error = "Bad mode"; message = "Mode must be Auto, Manual, or Hold." }
                    $statusCode = 400
                }
            }
            "/send-last-draft" {
                $esPath = Join-Path $SharedState.IndexFolder "last_email_summary.json"
                if (-not (Test-Path $esPath)) {
                    $responseData = @{ error = "Not Found"; message = "No last summary found." }
                    $statusCode = 404
                } else {
                    try {
                        $summary = Get-Content $esPath -Raw | ConvertFrom-Json
                        $entryId = [string]$summary.DraftEntryId
                        if ([string]::IsNullOrWhiteSpace($entryId)) {
                            $responseData = @{ error = "No draft"; message = "No saved draft is available for the last result." }
                            $statusCode = 400
                        } else {
                            Send-OutlookDraftByEntryIdApi -EntryId $entryId
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
                                    Write-ApiLog "send-last-draft history sync failed: $($_.Exception.Message)" "WARN"
                                }
                            }

                            Write-ApiLog "Sent Outlook draft from dashboard (EntryId=$entryId)"
                            $responseData = @{ message = "Draft sent" }
                            $statusCode = 200
                        }
                    } catch {
                        Write-ApiLog "send-last-draft failed: $($_.Exception.Message)" "ERROR"
                        $responseData = @{ error = "Send failed"; message = $_.Exception.Message }
                        $statusCode = 500
                    }
                }
            }
            "/clean-index" {
                if ($SharedState.ActiveTask -ne "None") {
                    $responseData = @{ error = "Busy"; message = "Another task is already running: $($SharedState.ActiveTask)" }
                    $statusCode = 409
                } else {
                    $SharedState.ActiveTask = "Cleaning Index"
                    $SharedState.PendingTask = "clean"
                    Write-ApiLog "Received API request: clean-index"
                    $responseData = @{ message = "Index Cleaning Started" }
                    $statusCode = 202
                }
            }
            "/clear-last-summary" {
                $esPath = Join-Path $SharedState.IndexFolder "last_email_summary.json"
                if (Test-Path $esPath) { Remove-Item $esPath -Force }
                $responseData = @{ message = "Last summary cleared" }
                $statusCode = 200
            }
            "/clear-history-all" {
                $hPath = Join-Path $SharedState.IndexFolder "transmittal_history.json"
                if (Test-Path $hPath) { Set-Content $hPath "[]" -Force }
                $responseData = @{ message = "History cleared" }
                $statusCode = 200
            }
            "/clear-history-item" {
                $idxStr = Get-QueryValue -QueryParams $QueryParams -Name "index"
                if ($null -ne $idxStr) {
                    $targetIdx = [int]$idxStr
                    $hPath = Join-Path $SharedState.IndexFolder "transmittal_history.json"
                    if (Test-Path $hPath) {
                        $history = Read-SharedJsonFile -Path $hPath
                        if ($null -eq $history) { $history = @() } else { $history = @($history) }
                        if ($targetIdx -ge 0 -and $targetIdx -lt $history.Count) {
                            if ($history.Count -eq 1) {
                                $history = @()
                            } else {
                                $newHistory = @()
                                for ($i = 0; $i -lt $history.Count; $i++) {
                                    if ($i -ne $targetIdx) { $newHistory += $history[$i] }
                                }
                                $history = $newHistory
                            }
                            @($history) | ConvertTo-Json -Depth 5 | Set-Content $hPath -Force
                            $responseData = @{ message = "Item cleared" }
                            $statusCode = 200
                        }
                    }
                }
            }
            "/search" {
                $q = Get-QueryValue -QueryParams $QueryParams -Name "q"
                if ($q) {
                    $csvPath = Join-Path $SharedState.IndexFolder "pdf_index_clean.csv"
                    if (Test-Path $csvPath) {
                        try {
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
                        $responseData = @{ error = "Index not found" }
                        $statusCode = 404
                    }
                } else {
                    $responseData = @{ error = "No query" }
                    $statusCode = 400
                }
            }
        }

        return New-JsonResponse -Body $responseData -StatusCode $statusCode
    }

    function Write-TcpResponse {
        param(
            $Client,
            [int]$StatusCode,
            [string]$ContentType,
            [byte[]]$BodyBytes
        )

        if ($null -eq $BodyBytes) { $BodyBytes = [byte[]]@() }

        $stream = $Client.GetStream()
        $writer = [System.IO.StreamWriter]::new($stream, [System.Text.Encoding]::ASCII, 1024, $true)
        try {
            $writer.NewLine = "`r`n"
            $writer.WriteLine("HTTP/1.1 $StatusCode $(Get-StatusDescription -StatusCode $StatusCode)")
            $writer.WriteLine("Content-Type: $ContentType")
            $writer.WriteLine("Content-Length: $($BodyBytes.Length)")
            $writer.WriteLine("Connection: close")
            $writer.WriteLine("Access-Control-Allow-Origin: *")
            $writer.WriteLine("Access-Control-Allow-Methods: GET, POST, OPTIONS")
            $writer.WriteLine("Access-Control-Allow-Headers: Content-Type")
            $writer.WriteLine("")
            $writer.Flush()

            if ($BodyBytes.Length -gt 0) {
                $stream.Write($BodyBytes, 0, $BodyBytes.Length)
                $stream.Flush()
            }
        } finally {
            try { $writer.Dispose() } catch {}
        }
    }

    function Handle-TcpClient {
        param(
            $Client,
            $SharedState,
            [string]$Dir
        )

        try {
            $stream = $Client.GetStream()
            $stream.ReadTimeout = 3000
            $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::ASCII, $false, 4096, $true)
            try {
                $requestLine = $reader.ReadLine()
                if ([string]::IsNullOrWhiteSpace($requestLine)) { return }

                while ($true) {
                    $headerLine = $reader.ReadLine()
                    if ($null -eq $headerLine -or $headerLine -eq "") { break }
                }

                $parts = $requestLine -split ' '
                if ($parts.Count -lt 2) {
                    $badBytes = [System.Text.Encoding]::UTF8.GetBytes("Bad Request")
                    Write-TcpResponse -Client $Client -StatusCode 400 -ContentType "text/plain; charset=utf-8" -BodyBytes $badBytes
                    return
                }

                $method = $parts[0].ToUpperInvariant()
                $rawTarget = $parts[1]
                $path = $rawTarget
                $queryParams = @{}
                $qIndex = $rawTarget.IndexOf('?')
                if ($qIndex -ge 0) {
                    $path = $rawTarget.Substring(0, $qIndex)
                    $queryParams = Parse-QueryString -RawQuery $rawTarget.Substring($qIndex + 1)
                }

                $result = Invoke-HubApiRequest -Method $method -Path $path -QueryParams $queryParams -SharedState $SharedState -Dir $Dir
                Write-TcpResponse -Client $Client -StatusCode $result.StatusCode -ContentType $result.ContentType -BodyBytes $result.BodyBytes
            } finally {
                if ($reader) { try { $reader.Dispose() } catch {} }
            }
        } catch {
            Write-ApiLog "TCP request handling failed: $($_.Exception.Message)" "ERROR"
            try {
                $errBytes = [System.Text.Encoding]::UTF8.GetBytes("Internal Server Error")
                Write-TcpResponse -Client $Client -StatusCode 500 -ContentType "text/plain; charset=utf-8" -BodyBytes $errBytes
            } catch {}
        } finally {
            try { $Client.Close() } catch {}
        }
    }

    function Start-TcpApiServer {
        param(
            $SharedState,
            [string]$Dir
        )

        $listeners = New-Object System.Collections.ArrayList
        try {
            foreach ($ip in @([System.Net.IPAddress]::Parse("127.0.0.1"), [System.Net.IPAddress]::IPv6Loopback)) {
                try {
                    $tcpListener = [System.Net.Sockets.TcpListener]::new($ip, [int]$SharedState.Port)
                    $tcpListener.Start()
                    [void]$listeners.Add($tcpListener)
                } catch {}
            }

            if ($listeners.Count -eq 0) {
                $SharedState.Status = "Error: Port $($SharedState.Port) unavailable"
                Write-ApiLog "Failed to start dashboard API on port $($SharedState.Port)." "ERROR"
                return
            }

            $SharedState.Status = "Running"
            Write-ApiLog "Hub API listening with TcpListener on port $($SharedState.Port)." "WARN"

            while ($SharedState.Status -ne "Stopping") {
                $handledRequest = $false

                foreach ($listener in @($listeners)) {
                    if ($SharedState.Status -eq "Stopping") { break }

                    if ($listener.Pending()) {
                        $client = $listener.AcceptTcpClient()
                        Handle-TcpClient -Client $client -SharedState $SharedState -Dir $Dir
                        $handledRequest = $true
                    }
                }

                if (-not $handledRequest) {
                    Start-Sleep -Milliseconds 100
                }
            }
        } finally {
            foreach ($listener in @($listeners)) {
                try { $listener.Stop() } catch {}
            }
        }
    }

    function Start-HttpApiServer {
        param(
            $SharedState,
            [string]$Dir
        )

        $listener = $null
        try {
            $listener = New-Object System.Net.HttpListener
            $listener.Prefixes.Add("http://localhost:$($SharedState.Port)/")
            try { $listener.Prefixes.Add("http://127.0.0.1:$($SharedState.Port)/") } catch {}
            try { $listener.Prefixes.Add("http://[::1]:$($SharedState.Port)/") } catch {}
            $listener.Start()
            $SharedState.Status = "Running"
            Write-ApiLog "Hub API listening with HttpListener on port $($SharedState.Port)."

            while ($listener.IsListening) {
                if ($SharedState.Status -eq "Stopping") {
                    try { $listener.Stop() } catch {}
                    break
                }

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
                } catch {
                    break
                }

                $request = $context.Request
                $response = $context.Response
                try {
                    $queryParams = @{}
                    foreach ($key in $request.QueryString.AllKeys) {
                        if ($null -ne $key) {
                            $queryParams[$key] = [string]$request.QueryString[$key]
                        }
                    }

                    $result = Invoke-HubApiRequest -Method $request.HttpMethod -Path $request.Url.AbsolutePath -QueryParams $queryParams -SharedState $SharedState -Dir $Dir
                    $response.AddHeader("Access-Control-Allow-Origin", "*")
                    $response.AddHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
                    $response.AddHeader("Access-Control-Allow-Headers", "Content-Type")
                    $response.ContentType = $result.ContentType
                    $response.ContentLength64 = $result.BodyBytes.Length
                    $response.StatusCode = $result.StatusCode
                    if ($result.BodyBytes.Length -gt 0) {
                        $response.OutputStream.Write($result.BodyBytes, 0, $result.BodyBytes.Length)
                    }
                } catch {
                    try { $response.StatusCode = 500 } catch {}
                } finally {
                    try { $response.Close() } catch {}
                }
            }

            return $true
        } catch {
            Write-ApiLog "HttpListener unavailable on port $($SharedState.Port): $($_.Exception.Message)" "WARN"
            return $false
        } finally {
            if ($listener) {
                try { $listener.Close() } catch {}
            }
        }
    }

    if (-not (Start-HttpApiServer -SharedState $SharedState -Dir $Dir)) {
        Start-TcpApiServer -SharedState $SharedState -Dir $Dir
    }
}
