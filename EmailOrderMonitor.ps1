# ==============================================================================
#  EmailOrderMonitor.ps1  v2.0 - Nordic Minesteel Technologies
#  Automated Spare Parts Order Processing (PDM Email -> Word Doc -> Collection)
# ==============================================================================
#  Monitors an Outlook mailbox folder for PDM state-change notifications.
#  When a new order email arrives, the monitor:
#    1. Extracts the .docx file path from the PDM notification
#    2. Opens the F80 order form (Word document) in the PDM vault
#    3. Reads part numbers from the Project Overview table
#    4. Collects all matching PDFs/DXFs via SimpleCollector
#    5. Sends a transmittal email summarizing results
#
#  Usage:
#    EmailOrderMonitor.ps1                    # Run once (check & process)
#    EmailOrderMonitor.ps1 -Watch             # Poll continuously
#    EmailOrderMonitor.ps1 -PollInterval 120  # Poll every 2 minutes
#
#  Configuration: Set email settings in config.json under "emailMonitor" key.
# ==============================================================================

param(
    [switch]$Watch,                  # Continuously poll for new orders
    [int]$PollInterval = 300,        # Seconds between polls (default 5 min)
    [switch]$TestMode,               # If set, don't move emails or send real transmittals
    [string]$Config = "config.json", # Path to config file
    [ValidateSet("Auto","Manual","Hold")]
    [string]$DispatchMode = ""
)

# --- Logging (Defined first so it's available for config errors) ---
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $prefix = switch ($Level) {
        "ERROR"   { "[!]" }
        "WARN"    { "[~]" }
        "SUCCESS" { "[+]" }
        default   { "[ ]" }
    }
    $entry = "$ts $prefix $Message"
    
    # Try to write to file if logFile is defined, otherwise just console
    if ($script:logFile) {
        Add-Content -Path $script:logFile -Value $entry -ErrorAction SilentlyContinue
    }
    
    $color = switch ($Level) {
        "ERROR"   { "Red" }
        "WARN"    { "Yellow" }
        "SUCCESS" { "Green" }
        default   { "Gray" }
    }
    Write-Host $entry -ForegroundColor $color
}

# --- Hub Notification (drops a timestamped JSON file; HubService picks it up via its tray loop) ---
function Push-HubNotification {
    param([string]$Title, [string]$Message, [string]$FolderPath = "")
    try {
        $stamp      = (Get-Date).ToString("yyyyMMdd_HHmmss_fff")
        $notifyFile = Join-Path $indexFolder "notify_${stamp}.json"
        @{ Title = $Title; Message = $Message; FolderPath = $FolderPath } |
            ConvertTo-Json | Set-Content $notifyFile -Encoding UTF8 -Force
    } catch { }
}

function Write-EmailProgress {
    param(
        [string]$Step,
        [string]$Order  = "",
        [string]$Detail = ""
    )
    try {
        @{
            Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Step      = $Step
            Order     = $Order
            Detail    = $Detail
        } | ConvertTo-Json | Set-Content -Path (Join-Path $indexFolder "email_progress.json") -Encoding UTF8 -Force
    } catch { }
}

# --- Load config ---
$scriptDir = Split-Path $PSCommandPath -Parent
$configPath = if ([System.IO.Path]::IsPathRooted($Config)) { $Config } else { Join-Path $scriptDir $Config }
Write-Log "Loading configuration from: $configPath" "INFO"
$cfg = @{}
if (Test-Path $configPath) {
    try { 
        $cfg = Get-Content $configPath -Raw | ConvertFrom-Json 
    } catch { 
        Write-Log "Error reading config at $configPath : $($_.Exception.Message)" "ERROR"
    }
}

$indexFolder  = if ($cfg.indexFolder)  { $cfg.indexFolder }  else { "C:\Users\dlebel\Documents\PDFIndex" }
$outputRoot   = if ($cfg.outputFolder) { $cfg.outputFolder } else { "C:\Users\dlebel\Documents\AssemblyPDFs" }
$logFolder    = if ($cfg.logFolder)    { $cfg.logFolder }    else { (Join-Path $outputRoot "MacroLogs") }

# PDM vault path - the email contains paths relative to this root
$pdmVaultPath = if ($cfg.pdmVaultPath) { $cfg.pdmVaultPath } else { "C:\NMT_PDM" }
$crawlRoots   = if ($cfg.crawlRoots)   { @($cfg.crawlRoots) } else { @($pdmVaultPath) }
$enableProjectIndexFallback = $false
if ($null -ne $cfg.enableProjectIndexFallback) {
    try { $enableProjectIndexFallback = [bool]$cfg.enableProjectIndexFallback } catch { $enableProjectIndexFallback = $false }
}

# Email monitor settings (from config.json emailMonitor section)
$emailCfg = if ($cfg.emailMonitor) { $cfg.emailMonitor } else { @{} }
$monitorFolder     = if ($emailCfg.inboxFolder)     { $emailCfg.inboxFolder }     else { "NMT_PDM" }
$processedFolder   = if ($emailCfg.processedFolder) { $emailCfg.processedFolder } else { "NMT_PDM\Processed" }
$collectMode       = if ($emailCfg.collectMode)     { $emailCfg.collectMode }     else { "BOTH" }
$transmittalTo     = if ($emailCfg.transmittalTo)   { $emailCfg.transmittalTo }   else { "dlebel@nmtech.com" }
$transmittalCc     = if ($emailCfg.transmittalCc)   { $emailCfg.transmittalCc }   else { "" }
$transmittalToProd = if ($emailCfg.transmittalToProd) { $emailCfg.transmittalToProd } else { $transmittalTo }
$transmittalCcProd = if ($emailCfg.transmittalCcProd) { $emailCfg.transmittalCcProd } else { $transmittalCc }
$transmittalToTest = if ($emailCfg.transmittalToTest) { $emailCfg.transmittalToTest } else { $transmittalTo }
$transmittalCcTest = if ($emailCfg.transmittalCcTest) { $emailCfg.transmittalCcTest } else { $transmittalCc }
$forceSendReceiveBeforeScan = $false
if ($null -ne $emailCfg.forceSendReceiveBeforeScan) {
    try { $forceSendReceiveBeforeScan = [bool]$emailCfg.forceSendReceiveBeforeScan } catch { $forceSendReceiveBeforeScan = $false }
}
$sendReceiveWaitSeconds = 4
if ($null -ne $emailCfg.sendReceiveWaitSeconds) {
    try { $sendReceiveWaitSeconds = [int]$emailCfg.sendReceiveWaitSeconds } catch { $sendReceiveWaitSeconds = 4 }
}
if ($sendReceiveWaitSeconds -lt 0) { $sendReceiveWaitSeconds = 0 }
if ($sendReceiveWaitSeconds -gt 30) { $sendReceiveWaitSeconds = 30 }
$enableAssemblyBomExpansion = $true
if ($null -ne $emailCfg.enableAssemblyBomExpansion) {
    try { $enableAssemblyBomExpansion = [bool]$emailCfg.enableAssemblyBomExpansion } catch { $enableAssemblyBomExpansion = $true }
}
$requireRunningSolidWorksForBomExpansion = $false
if ($null -ne $emailCfg.requireRunningSolidWorksForBomExpansion) {
    try { $requireRunningSolidWorksForBomExpansion = [bool]$emailCfg.requireRunningSolidWorksForBomExpansion } catch { $requireRunningSolidWorksForBomExpansion = $false }
}
$preferVbsMacroOnly = $false
if ($null -ne $emailCfg.preferVbsMacroOnly) {
    try { $preferVbsMacroOnly = [bool]$emailCfg.preferVbsMacroOnly } catch { $preferVbsMacroOnly = $false }
}
$allowRecursiveModelSearchFallback = $false
if ($null -ne $emailCfg.allowRecursiveModelSearchFallback) {
    try { $allowRecursiveModelSearchFallback = [bool]$emailCfg.allowRecursiveModelSearchFallback } catch { $allowRecursiveModelSearchFallback = $false }
}
$allowDirectComBomTraversalFallback = $false
if ($null -ne $emailCfg.allowDirectComBomTraversalFallback) {
    try { $allowDirectComBomTraversalFallback = [bool]$emailCfg.allowDirectComBomTraversalFallback } catch { $allowDirectComBomTraversalFallback = $false }
}
$enableVbsHelperFallback = $false
if ($null -ne $emailCfg.enableVbsHelperFallback) {
    try { $enableVbsHelperFallback = [bool]$emailCfg.enableVbsHelperFallback } catch { $enableVbsHelperFallback = $false }
}
$configuredDispatchMode = if ($emailCfg.dispatchMode) { [string]$emailCfg.dispatchMode } else { "" }
if ([string]::IsNullOrWhiteSpace($DispatchMode)) {
    $DispatchMode = $configuredDispatchMode
}
if ([string]::IsNullOrWhiteSpace($DispatchMode)) {
    $DispatchMode = if ($TestMode) { "Manual" } else { "Auto" }
}
if (@("Auto","Manual","Hold") -notcontains $DispatchMode) {
    $DispatchMode = if ($TestMode) { "Manual" } else { "Auto" }
}

function Get-StoredEpicorCredential {
    $credPath = Join-Path (Join-Path $env:LOCALAPPDATA "EpicorOrderMonitor") "epicor-creds.xml"
    if (-not (Test-Path $credPath)) { return $null }
    try {
        $stored = Import-Clixml -Path $credPath
        if ($stored.PSObject.Properties.Name -contains "Username" -and $stored.PSObject.Properties.Name -contains "Password") {
            $userSecure = ConvertTo-SecureString -String $stored.Username
            $passSecure = ConvertTo-SecureString -String $stored.Password
            $username = [System.Net.NetworkCredential]::new("", $userSecure).Password
            $password = [System.Net.NetworkCredential]::new("", $passSecure).Password
            return [pscustomobject]@{
                UserName = $username
                Password = $password
            }
        }
        if ($stored -is [pscredential]) {
            return [pscustomobject]@{
                UserName = $stored.UserName
                Password = $stored.GetNetworkCredential().Password
            }
        }
        return $null
    } catch {
        Write-Host "WARN: Could not load stored Epicor credentials from $credPath" -ForegroundColor Yellow
        return $null
    }
}

$epicorCfg          = if ($cfg.epicor)                      { $cfg.epicor }                      else { @{} }
$epicorApiUrl       = if ($epicorCfg.apiUrl)                { $epicorCfg.apiUrl.TrimEnd('/') }   else { "https://epicor.groupnmt.com/Prod" }
$epicorCompany      = if ($epicorCfg.company)               { $epicorCfg.company }               else { "NMT" }
$epicorUser         = if ($epicorCfg.username)              { $epicorCfg.username }              else { "" }
$epicorPass         = if ($epicorCfg.password)              { $epicorCfg.password }              else { "" }
$storedEpicorCred   = Get-StoredEpicorCredential
if ($storedEpicorCred) {
    if ([string]::IsNullOrWhiteSpace($epicorUser) -or ($epicorUser -in @("YOUR_EPICOR_USERNAME", "YOUR_USERNAME_HERE"))) {
        $epicorUser = $storedEpicorCred.UserName
    }
    if ([string]::IsNullOrWhiteSpace($epicorPass) -or ($epicorPass -in @("YOUR_EPICOR_PASSWORD", "YOUR_PASSWORD_HERE"))) {
        $epicorPass = $storedEpicorCred.Password
    }
}
$epicorUserIsPlaceholder = [string]::IsNullOrWhiteSpace($epicorUser) -or ($epicorUser -in @("YOUR_EPICOR_USERNAME", "YOUR_USERNAME_HERE"))
$epicorPassIsPlaceholder = [string]::IsNullOrWhiteSpace($epicorPass) -or ($epicorPass -in @("YOUR_EPICOR_PASSWORD", "YOUR_PASSWORD_HERE"))
$epicorEnabled      = ($epicorCfg.enabled -ne $false) -and (-not $epicorUserIsPlaceholder) -and (-not $epicorPassIsPlaceholder)
$script:epicorHeaders = $null
$script:historyPdfIndex = $null
$script:historyDxfIndex = $null

# --- Logging Directory Setup ---
$monitorLogDir = Join-Path $logFolder "EmailMonitor"
if (-not (Test-Path $monitorLogDir)) { New-Item -ItemType Directory -Path $monitorLogDir -Force | Out-Null }
$script:logFile    = Join-Path $monitorLogDir "monitor_$(Get-Date -Format 'yyyy-MM-dd').log"
$script:ocrDebugLog = Join-Path $monitorLogDir "ocr_debug_$(Get-Date -Format 'yyyy-MM-dd').txt"

# ==============================================================================
#  Automated Image OCR Fallback (Pure PowerShell WinRT Integration)
# ==============================================================================

# Reliable WinRT async helper  -  uses AsTask() so .NET's thread pool drives the
# completion rather than a polling loop (polling fails in PowerShell's MTA thread).
# System.Runtime.WindowsRuntime.dll must be loaded explicitly; it is not auto-loaded
# in PowerShell 5.1 even though it ships with .NET Framework 4.x.
$script:_asTaskGeneric = $null
function Invoke-WinRTAsync {
    param($AsyncOp, [Type]$TResult)
    if (-not $script:_asTaskGeneric) {
        $rtDll = Join-Path ([System.Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()) `
                           'System.Runtime.WindowsRuntime.dll'
        if (Test-Path $rtDll) { [void][System.Reflection.Assembly]::LoadFrom($rtDll) }
        $script:_asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() |
            Where-Object { $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 })[0]
    }
    $task = $script:_asTaskGeneric.MakeGenericMethod($TResult).Invoke($null, @($AsyncOp))
    if (-not $task.Wait(8000)) { throw "WinRT async timed out after 8s" }
    return $task.Result
}

function Get-OCRTextFromImage {
    param ([string]$ImagePath)
    # Uses Tesseract OCR with 3 passes for maximum accuracy on small table images.
    # Each PSM mode captures different parts correctly  -  we combine all results.
    $tesseractExe = "C:\Program Files\Tesseract-OCR\tesseract.exe"
    $tempUpscaled = $null
    try {
        if (Test-Path $tesseractExe) {
            Add-Type -AssemblyName System.Drawing
            $img = [System.Drawing.Image]::FromFile($ImagePath)
            $origW = $img.Width; $origH = $img.Height
            Write-Log "    Image: ${origW}x${origH}" "INFO"

            # Pass 1: Raw image, auto PSM (best for dash-style parts like 25347-A03)
            $rawText = & $tesseractExe $ImagePath stdout 2>$null
            $rawText = ($rawText -join "`n")
            Write-Log "    Pass 1 (raw auto): $($rawText.Length) chars" "INFO"

            # Smart upscale: larger scale for tiny images, smaller for big ones.
            # Upscaling helps Tesseract read small F80 table fonts.
            $scale = if ($origW -lt 400) { 8 } elseif ($origW -lt 800) { 4 } else { 2 }
            $nw = [int]($origW * $scale); $nh = [int]($origH * $scale)
            $bmp = New-Object System.Drawing.Bitmap($nw, $nh)
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $g.DrawImage($img, 0, 0, $nw, $nh)
            $g.Dispose(); $img.Dispose()
            $tempUpscaled = $ImagePath + '.upscaled.png'
            $bmp.Save($tempUpscaled, [System.Drawing.Imaging.ImageFormat]::Png)
            $bmp.Dispose()

            # Pass 2: Upscaled, auto PSM (best for 76359, McMaster digits)
            $upAutoText = & $tesseractExe $tempUpscaled stdout 2>$null
            $upAutoText = ($upAutoText -join "`n")
            Write-Log "    Pass 2 (upscale auto): $($upAutoText.Length) chars" "INFO"

            # Pass 3: Upscaled, PSM 6 (best for structured table rows)
            $upPsm6Text = & $tesseractExe $tempUpscaled stdout --psm 6 2> $null
            $upPsm6Text = ($upPsm6Text -join "`n")
            Write-Log "    Pass 3 (upscale psm6): $($upPsm6Text.Length) chars" "INFO"

            $allText = $rawText + "`n" + $upAutoText + "`n" + $upPsm6Text

            if (-not [string]::IsNullOrWhiteSpace($allText)) {
                Write-Log "    Tesseract total: $($allText.Length) chars from all passes" "SUCCESS"
            } else {
                Write-Log "    Tesseract returned empty" "WARN"
            }
            return $allText
        } else {
            Write-Log "    Tesseract not found, skipping OCR" "WARN"
            return ""
        }
    } catch {
        Write-Log "    OCR error: $($_.Exception.Message)" "WARN"
        return ""
    } finally {
        if ($tempUpscaled -and (Test-Path $tempUpscaled)) {
            Remove-Item $tempUpscaled -Force -ErrorAction SilentlyContinue
        }
    }
}

function Repair-OCRText {
    param([string]$Text)
    # Normalize em/en dashes
    $Text = $Text -replace '[\u2013\u2014\u2012]', '-'
    # Common OCR word misreads in Epicor F80 font
    $Text = $Text -replace 'Wanuactired', 'Manufactured'
    $Text = $Text -replace 'Wianuia', 'Manufactured'
    $Text = $Text -replace 'Manufa\b', 'Manufactured'
    $Text = $Text -replace 'Wianufactured', 'Manufactured'
    $Text = $Text -replace 'Manufacturedcture', 'Manufacture'
    # [ or ( misread as G in assembly refs: "[02-04-08" -> "G02-04-08"
    $Text = [regex]::Replace($Text, '(?m)^[\[\(]([A-Z0-9]\d{2}-\d{2}-\d{2,3})', 'G$1')
    # G-prefix OCR: "GO2-04-06" -> "G02-04-06" (O->0 after G in assy codes)
    $Text = [regex]::Replace($Text, '\bGO(\d)', 'G0$1')
    # G-prefix OCR: "GT1-20-08" -> "G11-20-08" (T->1 after G in assy codes)
    $Text = [regex]::Replace($Text, '\bGT(\d)', 'G1$1')
    # T->1 at start of NMT job number:  "T7141-10-P58-R" -> "17141-10-P58-R"
    $Text = [regex]::Replace($Text, '\bT(\d{4}-\d{2}-)', '1$1')
    # E->8 in job part suffix:  "17141-10-P5E-R" -> "17141-10-P58-R"
    $Text = [regex]::Replace($Text, '(?<=\d{5}-\d{2}-[A-Z]\d?)E(?=\d|-[A-Z]\b|\b)', '8')
    # ZP misread as 2P in hardware:  "-2P" at word boundary -> "-ZP"
    $Text = [regex]::Replace($Text, '(?<=[A-Z0-9])-2P\b', '-ZP', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    # O->"- 0." in hardware prefix:  "FWO625-" -> "FW-0.625-"
    $Text = [regex]::Replace($Text, '\b([A-Z]{2,5})O(\d{3})(?=[-x])', '$1-0.$2', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    # Missing decimal after dash:  "-0625-" -> "-0.625-"
    $Text = [regex]::Replace($Text, '(?<=-)\b0(\d{3})\b(?=-)', '0.$1')
    # Missing decimal after x in length dimension:  "x2250-" -> "x2.250-"
    $Text = [regex]::Replace($Text, '(?<=x)(\d)(\d{3})(?=-)', '$1.$2', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    # Missing dash before terminal 2-letter suffix:  "SAEZP" -> "SAE-ZP"
    $Text = [regex]::Replace($Text, '(?<=[A-Z]{3,})(?<!-)(ZP|ZY|SS|GS|PF)\b', '-$1', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    # O->0 and T->1 in short alpha-numeric codes:  "WOT0A" -> "W010A"
    $Text = [regex]::Replace($Text, '\b([A-Z])([O0ITi])([0-9OITi]{0,4})([A-Z])\b', {
        param($m)
        $inner = ($m.Groups[2].Value + $m.Groups[3].Value) -replace 'O','0' -replace 'T','1' -replace 'I','1' -replace 'i','1' -replace 'o','0'
        $m.Groups[1].Value + $inner + $m.Groups[4].Value
    }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    # U->L in "UINER" (Tesseract misreads "L" as "U" in Epicor table font: "LINER" -> "UINER")
    $Text = [regex]::Replace($Text, '\bUINER\b', 'LINER', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    
    # Backslash before letters in part numbers (OCR artifact): "1202-\WM-440" -> "1202-WM-440"
    $Text = $Text -replace '\\(?=[A-Z])', ''
    # OCR misread 6->5 in hardware decimal: "FHCS-0.626" -> "FHCS-0.625"
    $Text = $Text -replace '\b(FHCS|FW|HN)-0\.626\b', '$1-0.625'
    # OCR misread 0->O in known hardware prefixes: "FW-O.500" -> "FW-0.500"
    $Text = [regex]::Replace($Text, '\b(FW|HN|FHCS)-O\.', '$1-0.')
    # OCR misread W->V in hardware prefix: "FV-0.625" -> "FW-0.625"
    $Text = [regex]::Replace($Text, '\bFV-(\d)', 'FW-$1')
    # FHCS missing decimal in length after x: "x2250-" -> "x2.250-"
    $Text = [regex]::Replace($Text, '(?<=x)(\d)(\d{3})(?=-)', '$1.$2', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    # 2P->ZP without preceding dash: "2502P-" -> "250-ZP-" (OCR drops dash before ZP)
    $Text = [regex]::Replace($Text, '(\d)2P(?=-)', '$1-ZP')
    # NMT bearing OCR: "NMT1S64" -> "NMT TS64", "NMTTS64" -> "NMT TS64", "NMT 1S64" -> "NMT TS64"
    $Text = $Text -replace '\bNMT1S(\d)', 'NMT TS$1'
    $Text = $Text -replace '\bNMTTS(\d)', 'NMT TS$1'
    $Text = $Text -replace '\bNMT\s+1S(\d)', 'NMT TS$1'
    # Hardware decimal SO0 -> 500: "FW-0.SO0-SAE-ZP" -> "FW-0.500-SAE-ZP"
    $Text = [regex]::Replace($Text, '(?<=\b(?:FW|HN|FHCS)-0\.)SO0\b', '500')
    # HP10 trailing -8 -> -B (OCR misreads B as 8 in suffix)
    $Text = [regex]::Replace($Text, '\b(HP10-[0-9A-Z]+-[0-9A-Z]+-N)-8\b', '$1-B')
    # K1T -> KIT (OCR misreads I as 1)
    $Text = $Text -replace '\bK1T\b', 'KIT'
    # Known bearing OCR misreads
    $Text = $Text -replace 'SNIOSOS6C', 'SNLD 3056 G'
    $Text = $Text -replace 'SNI\s*D\s*3056\s*G', 'SNLD 3056 G'

    # Strip non-ASCII characters to remove mojibake
    $Text = $Text -replace '[^\x20-\x7E\r\n]', ' '

    # Clean up random standalone symbols
    $Text = $Text -replace '(?m)^\s*\|\s*$', ''
    
    # Fix OCR dropping decimal point on quantities (e.g., 200T EA -> 2.00 EA)
    $Text = [regex]::Replace($Text, '(?<=\s)(\d+)00(?=[A-Za-z]?\s*EA\b)', '$1.00', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    
    return $Text
}

function Test-HardwareLikePartNumber {
    param([string]$PartNumber)
    if ([string]::IsNullOrWhiteSpace($PartNumber)) { return $false }
    $pn = $PartNumber.Trim().ToUpperInvariant()
    if ($pn -match '^[A-Z]{2,6}-\d+[.,]\d{1,4}(?:[-X][A-Z0-9.]+){1,6}$') { return $true }
    if ($pn -match '\d+[.,]\d') { return $true }
    if ($pn -match '(?i)(?:^|-)GA\d{2,4}(?:HDG)?(?:-|$)') { return $true }
    if ($pn -match '(?i)(?:^|-)ZP(?:-|$)|(?:^|-)HDG(?:-|$)|(?:^|-)PLT(?:-|$)') { return $true }
    return $false
}

function Test-DrawingLikePartNumber {
    param([string]$PartNumber)
    if ([string]::IsNullOrWhiteSpace($PartNumber)) { return $false }
    $pn = $PartNumber.Trim().ToUpperInvariant()
    if (Test-HardwareLikePartNumber -PartNumber $pn) { return $false }
    # NMT job parts: 17141-10-P67, 20045-10-P23-R
    if ($pn -match '^\d{5}-\d{2}-[A-Z]\d{2,3}(?:-[LR])?$') { return $true }
    # NMT style with letters: 1202-WM-440, 25347-A02, 4823-A10
    if ($pn -match '^\d{4,6}[-_][A-Z0-9]{1,10}(?:[-_][A-Z0-9]{1,10})*$' -and $pn -match '[A-Z]') { return $true }
    # Family number format (all-digit with dash): 5035-253, 1035-018
    if ($pn -match '^\d{4}-\d{3}$') { return $true }
    # Alpha-dash chains: HP10-21A-0-N-B, BF12C11-P
    if ($pn -match '^[A-Z]{1,6}\d{1,4}(?:-[A-Z0-9]{1,6}){1,8}$') { return $true }
    # Multi-letter prefix + digits: PB23056-FS, HM3056-S
    if ($pn -match '^[A-Z]{2,4}\d{4,6}(?:-[A-Z]{1,3})?$') { return $true }
    # Assembly ref: G02-04-08, G11-20-08
    if ($pn -match '^[A-Z]\d{2}-\d{2}-\d{2,3}$') { return $true }
    # Digit-dash-digit with suffix: 02-05-00-C30-PU, 02-03-00-PU
    if ($pn -match '^\d{2}-\d{2}-\d{2,3}(?:-[A-Z0-9]{1,4}){0,3}$') { return $true }
    # Short alpha: MHB100EP, MFN014AB, W010A (letters + digits + letters, not hardware)
    if ($pn -match '^[A-Z]{1,4}\d{3,5}[A-Z]{1,2}$') { return $true }
    # Mixed alpha-digit with dash suffix: BF12C11-P
    if ($pn -match '^[A-Z]{1,3}\d{1,3}[A-Z]\d{1,3}(?:-[A-Z]{1,3})?$') { return $true }
    # Space-based bearing/pillow block nomenclature: SNLD 3056 G, MS 3056, NMT TS64-300
    if ($pn -match '^[A-Z]{2,4}\s+\d{3,5}(?:\s+[A-Z0-9]{1,4})?$') { return $true }
    if ($pn -match '^[A-Z]{2,4}\s+[A-Z0-9]{2,5}-\d{3}$') { return $true }
    if ($pn -match '^\d{5}\s+[A-Z]{2}\s+[A-Z0-9]{2,4}$') { return $true }
    # Plain 5-digit project/part number: 10569
    if ($pn -match '^\d{5}$') { return $true }
    return $false
}

function Normalize-Rev {
    param([string]$Rev)
    if ([string]::IsNullOrWhiteSpace($Rev)) { return "NA" }
    return ($Rev.Trim() -replace '^[Rr][Ee][Vv]', '').Trim().ToUpperInvariant()
}

function Compare-OrderRevToIndex {
    param([string]$OrderRev, [string]$IndexRev)
    $a = Normalize-Rev $OrderRev
    $b = Normalize-Rev $IndexRev
    if ($a -eq "NA" -or $b -eq "NA") { return "Unknown" }
    if ($a -ieq $b) { return "Match" }
    return "Mismatch"
}

function Get-RevisionComparison {
    param(
        [string]$OrderRev,
        [string]$EpicorRev,
        [string]$IndexRev
    )

    $order = Normalize-Rev $OrderRev
    $epicor = Normalize-Rev $EpicorRev
    $index = Normalize-Rev $IndexRev

    $epicorCompared = ($epicor -ne "NA")
    $indexCompared = ($index -ne "NA")
    $epicorMatches = ($epicorCompared -and $order -ne "NA" -and $order -ieq $epicor)
    $indexMatches = ($indexCompared -and $order -ne "NA" -and $order -ieq $index)
    $epicorMismatch = ($epicorCompared -and $order -ne "NA" -and -not $epicorMatches)
    $indexMismatch = ($indexCompared -and $order -ne "NA" -and -not $indexMatches)

    $status = "Unknown"
    $note = ""

    if ($order -eq "NA") {
        $status = "Unknown"
        $note = "Order rev missing"
    } elseif (-not $epicorCompared -and -not $indexCompared) {
        $status = "Unknown"
        $note = "No Epicor or index rev available"
    } elseif ($epicorMismatch -and $indexMismatch) {
        $status = "Epicor + Index Mismatch"
        $note = "Order rev differs from both Epicor and indexed drawing"
    } elseif ($epicorMismatch) {
        $status = "Epicor Mismatch"
        $note = "Epicor rev differs from the order"
    } elseif ($indexMismatch) {
        $status = "Index Mismatch"
        $note = "Indexed drawing rev differs from the order"
    } elseif ($epicorMatches -or $indexMatches) {
        $status = "Match"
        if ($epicorMatches -and $indexMatches) {
            $note = "Order, Epicor, and indexed drawing revs agree"
        } elseif ($epicorMatches) {
            $note = "Order matches Epicor"
        } else {
            $note = "Order matches indexed drawing"
        }
    }

    return [ordered]@{
        Status        = $status
        Note          = $note
        OrderRev      = if ($order -eq "NA") { "" } else { $order }
        EpicorRev     = if ($epicor -eq "NA") { "" } else { $epicor }
        IndexRev      = if ($index -eq "NA") { "" } else { $index }
        EpicorMatches = $epicorMatches
        IndexMatches  = $indexMatches
    }
}

function Get-EpicorHeaders {
    if ($script:epicorHeaders) { return $script:epicorHeaders }
    $b64 = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${epicorUser}:${epicorPass}"))
    $script:epicorHeaders = @{
        Authorization      = "Basic $b64"
        Accept             = "application/json"
        "x-epicor-company" = $epicorCompany
    }
    return $script:epicorHeaders
}

function Get-EpicorPartInfo {
    param([string]$PartNumber)
    if (-not $epicorEnabled) { return $null }
    try {
        $h = Get-EpicorHeaders
        $pnEnc = [Uri]::EscapeDataString($PartNumber)
        $url = "$epicorApiUrl/api/v1/Erp.BO.PartSvc/GetByID?partNum=$pnEnc"
        $resp = Invoke-RestMethod -Uri $url -Headers $h -Method Get -TimeoutSec 15 -ErrorAction Stop
        $partRows = $resp.returnObj.Part
        if (-not $partRows -or @($partRows).Count -eq 0) {
            return @{ Exists = $false; Revisions = @(); LatestRev = ""; Approved = $false; Description = "" }
        }
        $desc = $partRows[0].PartDescription
        $revRows = @($resp.returnObj.PartRev)
        if ($revRows.Count -eq 0) {
            return @{ Exists = $true; Revisions = @(); LatestRev = ""; Approved = $false; Description = $desc }
        }
        $approved = @($revRows | Where-Object { $_.Approved -eq $true })
        $pool = if ($approved.Count -gt 0) { $approved } else { $revRows }
        $latest = $pool | Sort-Object {
            $rv = ($_.RevisionNum -replace '[^0-9]', '')
            try { [int]$rv } catch { 0 }
        } -Descending | Select-Object -First 1
        return @{
            Exists      = $true
            Revisions   = $revRows
            LatestRev   = (Normalize-Rev $latest.RevisionNum)
            Approved    = [bool]$latest.Approved
            Description = $desc
        }
    } catch {
        return $null
    }
}

function Initialize-HistoryIndexes {
    if ($null -eq $script:historyPdfIndex) {
        $csvPath = Join-Path $indexFolder "pdf_index_clean.csv"
        if (Test-Path $csvPath) {
            try { $script:historyPdfIndex = @(Import-Csv $csvPath) } catch { $script:historyPdfIndex = @() }
        } else {
            $script:historyPdfIndex = @()
        }
    }
    if ($null -eq $script:historyDxfIndex) {
        $csvPath = Join-Path $indexFolder "dxf_index_clean.csv"
        if (Test-Path $csvPath) {
            try { $script:historyDxfIndex = @(Import-Csv $csvPath) } catch { $script:historyDxfIndex = @() }
        } else {
            $script:historyDxfIndex = @()
        }
    }
}

function Find-PartInHistoryIndex {
    param([string]$PartNumber, [switch]$Dxf)
    Initialize-HistoryIndexes
    $rows = if ($Dxf) { $script:historyDxfIndex } else { $script:historyPdfIndex }
    if (-not $rows) { return @() }
    $pn = $PartNumber.Trim().ToUpperInvariant()
    return @($rows | Where-Object { ([string]$_.BasePart).Trim().ToUpperInvariant() -eq $pn })
}

function Infer-DrawingPartFromIndexContext {
    param(
        [string]$ContextText,
        [string[]]$ExistingParts = @()
    )
    if ([string]::IsNullOrWhiteSpace($ContextText)) { return @() }
    if ([string]::IsNullOrWhiteSpace($indexFolder)) { return @() }

    $pdfIndexPath = Join-Path $indexFolder "pdf_index_clean.csv"
    $rows = @(Get-PdfIndexRowsCached -PdfIndexPath $pdfIndexPath)
    if ($rows.Count -eq 0) { return @() }

    $ctxU = $ContextText.ToUpperInvariant()
    $jobHint = ""
    $qtyHint = ""
    $clientTokens = @()
    $keywords = @()
    $keywordSeq = @()
    $keywordPairs = @()

    $mJob = [regex]::Match($ctxU, '(?i)\bJOB:\s*(\d{4,6})\b')
    if ($mJob.Success) { $jobHint = $mJob.Groups[1].Value }

    $mQty = [regex]::Match($ctxU, '(?i)\bQTY\s*[:\-]?\s*(\d{1,3})\b')
    if ($mQty.Success) { $qtyHint = $mQty.Groups[1].Value }

    $mClient = [regex]::Match($ctxU, '(?i)\bCLIENT:\s*([A-Z0-9 &\-_]{3,40})\b')
    if ($mClient.Success) {
        $clientTokens = @(
            [regex]::Matches($mClient.Groups[1].Value, '[A-Z0-9]{3,10}') |
            ForEach-Object { $_.Value } |
            Where-Object { $_ -notmatch '^(SPARE|PARTS|ORDER|CLIENT)$' } |
            Select-Object -Unique
        )
    }

    $stop = @(
        "JOB","CLIENT","FILE","ORDER","ORDERS","SPARE","PART","PARTS","PROJECT","DESIGN","DRAWINGS","DOCX",
        "F80","IN","ENGINEERING","THE","AND","FOR","WITH","QTY","QTY16","QTY4","QTY1"
    )
    $keywordSeq = @(
        [regex]::Matches($ctxU, '[A-Z]{4,12}') |
        ForEach-Object { $_.Value } |
        Where-Object { $stop -notcontains $_ }
    )
    $keywords = @($keywordSeq | Select-Object -Unique)
    if ($keywords.Count -gt 10) { $keywords = @($keywords | Select-Object -First 10) }
    if ($keywordSeq.Count -gt 1) {
        $pairs = [System.Collections.Generic.List[string]]::new()
        for ($i = 0; $i -lt ($keywordSeq.Count - 1); $i++) {
            $a = [string]$keywordSeq[$i]
            $b = [string]$keywordSeq[$i + 1]
            if ([string]::IsNullOrWhiteSpace($a) -or [string]::IsNullOrWhiteSpace($b)) { continue }
            if ($a -eq $b) { continue }
            $pair = "$a|$b"
            if (-not $pairs.Contains($pair)) { [void]$pairs.Add($pair) }
        }
        $keywordPairs = @($pairs | Select-Object -First 12)
    }

    $existingSet = @{}
    foreach ($ep in $ExistingParts) {
        $u = ([string]$ep).Trim().ToUpperInvariant()
        if (-not [string]::IsNullOrWhiteSpace($u)) { $existingSet[$u] = $true }
    }

    $bestByCore = @{}
    foreach ($row in $rows) {
        $ft = ""
        try { $ft = [string]$row.FileType } catch { }
        if ($ft -notmatch '(?i)^PDF$') { continue }

        $obs = ""
        try { $obs = [string]$row.IsObsolete } catch { }
        if ($obs -match '(?i)^(YES|TRUE|1)$') { continue }

        $fullPath = ""
        try { $fullPath = [string]$row.FullPath } catch { }
        if ([string]::IsNullOrWhiteSpace($fullPath)) { continue }
        $pathU = $fullPath.ToUpperInvariant()

        $core = ""
        $baseRaw = ""
        $fileName = ""
        try { $baseRaw = [string]$row.BasePart } catch { }
        try { $fileName = [string]$row.FileName } catch { }
        $mCore = [regex]::Match(([string]$baseRaw).ToUpperInvariant(), '^\s*(\d{4,6}-[A-Z]\d{1,3}(?:-[LR])?)\b')
        if (-not $mCore.Success) {
            $mCore = [regex]::Match(([string]$fileName).ToUpperInvariant(), '(\d{4,6}-[A-Z]\d{1,3}(?:-[LR])?)')
        }
        if (-not $mCore.Success) { continue }
        $core = $mCore.Groups[1].Value
        if ($existingSet.ContainsKey($core)) { continue }

        $score = 0
        $hits = 0

        if (-not [string]::IsNullOrWhiteSpace($jobHint) -and $pathU -match ("(?<!\\d)" + [regex]::Escape($jobHint) + "(?!\\d)")) {
            $score += 10
            $hits++
        }
        if (-not [string]::IsNullOrWhiteSpace($qtyHint)) {
            $qtyMatch = [regex]::Match($pathU, '(?i)\bQTY\s*[-_ ]*(\d{1,3})\b')
            if ($qtyMatch.Success) {
                $pathQty = [string]$qtyMatch.Groups[1].Value
                if ($pathQty -eq $qtyHint) {
                    $score += 8
                    $hits++
                } else {
                    $score -= 4
                }
            }
        }
        foreach ($ct in $clientTokens) {
            if ($pathU.Contains($ct)) {
                $score += 6
                $hits++
                break
            }
        }
        foreach ($kw in $keywords) {
            if ($pathU.Contains($kw)) {
                $score += 2
                $hits++
            }
        }
        foreach ($kp in $keywordPairs) {
            $bits = $kp.Split('|')
            if ($bits.Count -ne 2) { continue }
            if ($pathU.Contains($bits[0]) -and $pathU.Contains($bits[1])) {
                $score += 3
                $hits++
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($fileName)) {
            $fnU = $fileName.ToUpperInvariant()
            if ($fnU -match ('^' + [regex]::Escape($core) + '(?:[ _-]REV|\.)')) {
                $score += 2
            }
        }
        $baseU = ([string]$baseRaw).ToUpperInvariant().Trim()
        if (-not [string]::IsNullOrWhiteSpace($baseU)) {
            if ($baseU -eq $core) {
                $score += 4
                $hits++
            } elseif ($baseU.StartsWith($core + "-") -or $baseU.StartsWith($core + " ")) {
                $score -= 2
            } elseif ($baseU.Contains($core)) {
                $score -= 4
            }
            if ($baseU -match '(?i)(?:^|-)KIT(?:$|[^A-Z0-9])') { $score -= 5 }
            if ($baseU -match '(?i)^NMT-\d{3,5}\b') { $score -= 4 }
            if ($baseU -match '(?i)\bQA\b|DIM\s*CHECK') { $score -= 6 }
        }
        $fileU = ([string]$fileName).ToUpperInvariant()
        if ($fileU -match '(?i)\bQA\b|DIM\s*CHECK') { $score -= 4 }
        if ($pathU -match '(?i)\\20 - QA\\|\\QA\\|DIM CHECK') { $score -= 6 }
        if ($score -lt 8) { continue }

        $last = [datetime]::MinValue
        try { $last = [datetime]::Parse([string]$row.LastWriteTime) } catch { }

        if (-not $bestByCore.ContainsKey($core)) {
            $bestByCore[$core] = [pscustomobject]@{ Core = $core; Score = $score; Hits = $hits; LastWriteTime = $last; Path = $fullPath }
        } else {
            $cur = $bestByCore[$core]
            if ($score -gt $cur.Score -or ($score -eq $cur.Score -and $last -gt $cur.LastWriteTime)) {
                $bestByCore[$core] = [pscustomobject]@{ Core = $core; Score = $score; Hits = $hits; LastWriteTime = $last; Path = $fullPath }
            }
        }
    }

    $ranked = @($bestByCore.Values | Sort-Object -Property @{Expression='Score';Descending=$true}, @{Expression='Hits';Descending=$true}, @{Expression='LastWriteTime';Descending=$true}, @{Expression='Core';Descending=$false})
    if ($ranked.Count -eq 0) { return @() }

    $top = $ranked[0]
    if ($ranked.Count -gt 1) {
        $second = $ranked[1]
        if (($top.Score - $second.Score) -lt 2 -and $top.Score -lt 12) {
            Write-Log "  OCR fallback inference ambiguous (top '$($top.Core)' score=$($top.Score), second '$($second.Core)' score=$($second.Score)); skipping inference." "WARN"
            return @()
        }
    }

    Write-Log "  OCR fallback inferred drawing part '$($top.Core)' from index context (score=$($top.Score), path='$($top.Path)')" "SUCCESS"
    return @($top.Core)
}

function Get-PartNumbersFromImages {
    param ([string[]]$ImagePaths)
    $extractedParts = New-Object System.Collections.Generic.List[string]
    $familyTokenCandidates = New-Object System.Collections.Generic.List[string]
    $allOcrText = ""
    $script:ocrInferenceText = ""
    $patterns = @(
        '\b\d{5}-\d{2}-[A-Z]\d{2,3}(?:-[A-Z])?\b',              # NMT job part:  17141-10-P67, 17141-10-P58-R
        '\b[A-Z]{2,5}-\d+[.,]\d{1,3}(?:[-x][A-Z0-9.]+){1,4}\b', # Hardware:      FHCS-0.625-11x2.250-ZP-F
        '\b[A-Z]{1,4}\d{3,5}[A-Z]{1,2}\b',                        # Short multi:   MHB100EP, MFN014AB, W010A
        '\b[A-Z]{1,6}\d{1,4}(?:-[A-Z0-9]{1,6}){1,8}\b',          # Alpha-dash:    HP10-21A-0-N-B, BF12C11-P
        '\b\d{4,6}(?:[-_][A-Z0-9]{1,8}){1,4}\b',                 # NMT style:     25347-A02, 1202-WM-440, 4823-P8-HOUSING
        '\b\d{5}[A-Z]\d{3,4}\b',                                  # McMaster:      90107A030, 98164A268
        '\b[A-Z]{2,4}\d{4,6}(?:-[A-Z]{1,3})?\b',                  # Prefix+digits: PB23056-FS, HM3056-S
        '\b\d{2}-\d{2}-\d{2,3}(?:-[A-Z0-9]{1,4}){0,3}\b',        # DD-DD-DD:      02-05-00-C30-PU
        '\b\d{5}\b',                                               # Plain 5-digit: 76359, 10569
        '\b[A-Z]\d{2}-\d{2}-\d{2,3}\b',                           # Assembly ref:  G02-04-08, G11-20-08
        '\b[A-Z]{2,4}\s+\d{4,5}(?:\s+[A-Z]{1,2})?\b',             # Space-based:   SNLD 3056 G, MS 3056
        '\b\d{5}\s+[A-Z]{2}\s+[A-Z0-9]{2,4}\b',                  # Bearing:       24056 CC W33
        '\b[A-Z]{2,4}\s+[A-Z0-9]{2,5}-\d{3}\b',                   # NMT TS64-300
        '\b[A-Z]{1,3}\d{1,3}[A-Z]\d{1,3}(?:-[A-Z]{1,3})?\b',      # Mixed alpha-digit: BF12C11-P
        '\b[A-Z]{2,4}-\d{2,3}-[A-Z]\d{1,3}(?:-[A-Z]{1,3})?\b',    # Alpha-prefix assy: RG-14-P1, RG-14-P3-BR
        '\b[A-Z]{2,4}\d+\.\d+[Xx]\d+\.?\d*[A-Z]?\b'             # Tire/roller: FGW3.875X10P, SGW2.375X9.87
    )
    # Helper: extract Rev/Description/Qty context for a part from OCR text.
    # $anchor is the text to search for in the OCR string - usually the part number itself, but for
    # suffix-scan hits it is the RAW garbled OCR token (e.g. "7T4-10-P5EL") because that is what
    # still appears literally in the repaired OCR text.
    # Character class covers comma, parens, @, & etc. to handle descriptions like "LINER, BUCKET BACK (D)".
    if (-not $script:ocrPartContext) { $script:ocrPartContext = @{} }
    $script:ocrHasManufactured = $false   # reset per-order; set true if any OCR row is "Manufactured" type
    $trySetContext = {
        param([string]$pn, [string]$anchor, [string]$text)
        if ($script:ocrPartContext.ContainsKey($pn)) { return }
        $pat = [regex]::Escape($anchor) + '\s+[\[\(\{]?\s*([A-Z0-9_]{1,4})\s*(?:_+\s*)?[\]\)\}]?\s*(?:\||\s)\s*([A-Z(][A-Z0-9 ,&@()./\-]{1,100}?)\s+(\d+[,.]?\d*)\s*[\]\[\)\(\|]?\s*EA'
        $ctx = [regex]::Match($text, $pat, 'IgnoreCase')
        if (-not $ctx.Success) {
            # Fallback when OCR drops table delimiters around the Rev column.
            $pat2 = [regex]::Escape($anchor) + '\s+[\[\(\{]?\s*([A-Z0-9_]{1,4})\s*(?:_+\s*)?[\]\)\}]?\s+([A-Z(][A-Z0-9 ,&@()./\-]{1,100}?)\s+(\d+[,.]?\d*)\s*[\]\[\)\(\|]?\s*EA'
            $ctx = [regex]::Match($text, $pat2, 'IgnoreCase')
        }
        if ($ctx.Success) {
            $rawRev = $ctx.Groups[1].Value.Trim().ToUpperInvariant()
            $cRev  = ($rawRev -replace '[^A-Z0-9]', '')
            $cRev  = $cRev -replace 'T','1' -replace 'I','1' -replace 'L','1' -replace 'O','0'
            if ($cRev -match '\d+') { $cRev = [regex]::Match($cRev, '\d+').Value }
            if ([string]::IsNullOrWhiteSpace($cRev)) { $cRev = "1" }
            $cDesc = (Get-Culture).TextInfo.ToTitleCase($ctx.Groups[2].Value.Trim().ToLower())
            $cQty  = ($ctx.Groups[3].Value -replace ',','.') -replace '\.0+$','' -replace '(\d+)\.\d+$','$1'
            # Prefer the LAST "N EA" value from the OCR row (often the true order qty in F80).
            # This fixes rows like "72.00 EA 2.00 EA", where first value can be OCR noise.
            $rowMatch = [regex]::Match($text, [regex]::Escape($anchor) + '(?<row>[^\r\n]{0,240})', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($rowMatch.Success) {
                $eaMatches = [regex]::Matches([string]$rowMatch.Groups['row'].Value, '(\d+[,.]?\d*)\s*[\]\[\)\(\|]?\s*EA', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                if ($eaMatches.Count -gt 0) {
                    $lastEa = [string]$eaMatches[$eaMatches.Count - 1].Groups[1].Value
                    if (-not [string]::IsNullOrWhiteSpace($lastEa)) {
                        $cQty = ($lastEa -replace ',','.') -replace '\.0+$','' -replace '(\d+)\.\d+$','$1'
                    }
                }
            }
            # F80 Epicor table: Tesseract misreads "1" as "7" in Rev and Qty columns due to font rendering.
            # Detect F80 table context via its column header "MfgJobTyp"; correct isolated "7" ? "1".
            if ($text -match '(?i)\bMfgJobT(?:yp(?:e)?)?\b|\bMfgJob\b') {
                if ($cRev -eq '7') { $cRev = '1'; Write-Log "  OCR Corrected Rev 7?1 for $pn (F80 font misread)" "WARN" }
                if ($cQty -eq '7') { $cQty = '1'; Write-Log "  OCR Corrected Qty 7?1 for $pn (F80 font misread)" "WARN" }
            }
            $script:ocrPartContext[$pn] = "$pn Rev.$cRev $cDesc Qty: $cQty"
            Write-Log "  OCR Context: $($script:ocrPartContext[$pn])" "INFO"
        }
    }

    $i = 0

    foreach ($path in $ImagePaths) {
        $i++
        $leaf   = [System.IO.Path]::GetFileName($path)
        $sizeKB = [Math]::Round((Get-Item $path).Length / 1KB, 1)
        Write-Log "  OCR image $i/$($ImagePaths.Count): $leaf ($sizeKB KB)" "INFO"
        if (-not (Test-Path $path)) { Write-Log "    (not found, skipping)" "WARN"; continue }

        $ocrText = Get-OCRTextFromImage -ImagePath $path
        if ([string]::IsNullOrWhiteSpace($ocrText)) { continue }

        # Log first 300 chars of OCR output for debugging
        $preview = $ocrText.Substring(0, [Math]::Min(300, $ocrText.Length)) -replace '[\r\n]+', ' '
        Write-Log "    OCR text: $preview..." "INFO"

        # Dump FULL raw OCR text to dedicated debug log
        if ($script:ocrDebugLog) {
            $sep = "=" * 60
            $header = "`n$sep`nIMAGE: $path`nSIZE:  $sizeKB KB`nTIME:  $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`n$sep`n"
            Add-Content -Path $script:ocrDebugLog -Value ($header + $ocrText + "`n") -ErrorAction SilentlyContinue
            Write-Log "    OCR raw text written to: $($script:ocrDebugLog)" "INFO"
        }

        # Normalize and repair common OCR character substitutions
        $ocrText = Repair-OCRText -Text $ocrText
        if ($allOcrText.Length -lt 24000) {
            $chunk = $ocrText
            if ($chunk.Length -gt 3000) { $chunk = $chunk.Substring(0, 3000) }
            $allOcrText += "`n" + $chunk
        }
        try {
            $inferText = ($ocrText -replace '[\r\n]+', ' ')
            $inferText = [regex]::Replace($inferText, '\s+', ' ').Trim()
            if ($inferText.Length -gt 600) { $inferText = $inferText.Substring(0, 600) }
            if (-not [string]::IsNullOrWhiteSpace($inferText) -and $script:ocrInferenceText.Length -lt 5000) {
                $script:ocrInferenceText += " " + $inferText
            }
        } catch { }

        # Capture short family tokens even when OCR drops a leading digit (e.g. "036-054").
        foreach ($fm in [regex]::Matches($ocrText, '\b\d{3,4}[-_]\d{3}\b')) {
            $tok = ($fm.Value.ToUpperInvariant() -replace '_', '-')
            if (-not $familyTokenCandidates.Contains($tok)) {
                [void]$familyTokenCandidates.Add($tok)
            }
        }

        # Detect order type: F80 rows with "Manufactured" = new parts being fabricated (need Construction
        # drawings). "No Job" = reorder of existing spare parts (Procurement only, no Construction).
        if (-not $script:ocrHasManufactured -and $ocrText -match '(?i)\bManufactured\b') {
            $script:ocrHasManufactured = $true
            Write-Log "  OCR: 'Manufactured' job type detected - Construction checkbox will be checked" "INFO"
        }

        # --- Standard pattern matching ---
        foreach ($p in $patterns) {
            $regMatches = [regex]::Matches($ocrText, $p, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            foreach ($m in $regMatches) {
                $pn = $m.Value.ToUpper() -replace '\s+', ''
                if ($pn -match '^20\d{2}$|^19\d{2}$') { continue }
                if ($pn -match '^JOB') { continue }
                # Skip paint codes, unit-of-measure, description text, addresses
                if ($pn -match '^RAL\s|^HEAD\s|^EA\s|^INO\b|^CRATE\b|^PRIME|^GRADE|^ASSY\s|^GR\d|^T1X|^THRD\s|^V2-|^I2-') { continue }
                # Skip metric bolt specs (M20X80, M27X100)
                if ($pn -match '^M\d+X\d+$') { continue }
                if ($pn -match '^(TYPE|WINE|LAKE|MINE)\b') { continue }
                # Skip OCR artifacts: N010E (NOTE), C104S (C1045 grade), S020RT (quantity)
                if ($pn -match '^N010[A-Z]$|^C104[A-Z]$|^S020[A-Z]{2}$') { continue }
                # Skip leading-zero 5-digit (OCR quantity/gridline artifacts: 05200, 05503)
                if ($pn -match '^0\d{4}$') { continue }
                # Strip trailing description-word suffixes (e.g. -FACE, -LEFT, -RIGHT)
                $pn = $pn -replace '-(FACE|LEFT|RIGHT|FRONT|REAR|SIDE|BACK|TOP|BOTTOM)$', ''
                # Skip US/CA state+zip: "OH 44077", "ON L9E"
                if ($pn -match '^(OH|ON|MB|ID|MN)\s+\d') { continue }
                # Skip all-digit hyphenated strings UNLESS they match a known family format or bare 5-digit.
                if ($pn -match '^[\d\-_]+$' -and $pn -notmatch '^\d{4}-\d{3}$' -and $pn -notmatch '^\d{5}$') { continue }
                # Skip strings containing month abbreviations (OCR date-field artifacts e.g. DATE09-MAR-2026OR)
                if ($pn -match '(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)') { continue }
                # Skip bare 5-digit numbers (project/quote numbers, page refs) when
                # better-qualified drawing-like parts have already been found.
                if ($pn -match '^\d{5}$') {
                    $hasDrawing = $extractedParts | Where-Object { Test-DrawingLikePartNumber -PartNumber $_ }
                    if ($hasDrawing) { continue }
                }
                if (-not $extractedParts.Contains($pn)) {
                    [void]$extractedParts.Add($pn)
                    Write-Log "  OCR Found Part: $pn" "SUCCESS"
                    & $trySetContext $pn $pn $ocrText
                }
            }
        }

        # --- McMaster correction: OCR misreads 'A' in McMaster numbers ---
        # Two cases:
        #   Case 1: A read as digit  â†’ 98164A268 becomes 981648268 (9 digits) â†’ Substring(6) to skip the fake digit
        #   Case 2: A dropped entirely â†’ 98164A268 becomes 98164268  (8 digits) â†’ Substring(5) keeps all remaining
        foreach ($m in [regex]::Matches($ocrText, '\b\d{8,10}\b')) {
            $digits = $m.Value
            $len = $digits.Length
            $candidates = @()
            if ($len -ge 9 -and $len -le 10) {
                # Case 1: A was read as a digit (length preserved)
                $candidates += $digits.Substring(0,5) + 'A' + $digits.Substring(6)
            }
            if ($len -eq 8) {
                # Case 2: A was dropped entirely (one char shorter)
                $candidates += $digits.Substring(0,5) + 'A' + $digits.Substring(5)
            }
            foreach ($corrected in $candidates) {
                if ($corrected -match '^\d{5}A\d{2,4}$' -and -not $extractedParts.Contains($corrected)) {
                    [void]$extractedParts.Add($corrected)
                    Write-Log "  OCR Corrected: $digits -> $corrected (McMaster A fix)" "SUCCESS"
                }
            }
        }

        # --- Job-context suffix scan: recover garbled NMT plate numbers ---
        # When OCR corrupts the 5-digit job prefix but the sub-number and plate suffix
        # remain partially readable, reconstruct using an already-extracted known prefix
        # plus character corrections: S->5, T->7, E->8, O->0
        # Examples:  "7T4-10-P5EL" -> P58-L   "T14-10-PST-R" -> P57-R   "W7T4T1O-PSTL" -> P57-L
        $knownJobPrefixes = @($extractedParts | Where-Object { $_ -match '^\d{5}-\d{2}-[A-Z]' } |
            ForEach-Object { [regex]::Match($_, '^\d{5}-\d{2}').Value } | Sort-Object -Unique)

        if ($knownJobPrefixes.Count -gt 0) {
            # Load PDF index once for alt-correction validation
            $jcPdfRows = @()
            if (-not [string]::IsNullOrWhiteSpace($indexFolder)) {
                $jcPdfRows = Get-PdfIndexRowsCached -PdfIndexPath (Join-Path $indexFolder "pdf_index_clean.csv")
            }

            # Try primary (O->0, S->5) first. Only if primary produces nothing valid
            # (0-leading or no result) try the alternate map (O->5, S->9).
            # This prevents alt from generating false positives like P97 when P57 already works.
            $resolveSuffix = {
                param([string]$RawSfx, [string]$RawSide, [string]$JobPrefix)
                foreach ($cs in @(@{O='0';S='5'}, @{O='5';S='9'})) {
                    $sfx  = $RawSfx -replace 'T','7' -replace 'E','8' -replace 'O',$cs.O -replace 'S',$cs.S
                    $side = $RawSide
                    if ($side -eq '' -and $sfx -match '^(\d{2,3})([LR])$') { $sfx=$Matches[1]; $side='-'+$Matches[2] }
                    if ($sfx -notmatch '^\d{2,3}$' -or $sfx[0] -eq '0') { continue }
                    $pn = "$JobPrefix-P$sfx$side"
                    if ($pn -notmatch '^\d{5}-\d{2}-P\d{2,3}(-[LR])?$') { continue }
                    # Primary succeeded - validate against index if available, then return immediately
                    if ($jcPdfRows.Count -gt 0) {
                        $base = $pn -replace '-[LR]$',''
                        if (-not ($jcPdfRows | Where-Object { $_.BasePart -eq $base })) { continue }
                    }
                    return @($pn)   # return on first valid correction set - don't try alt
                }
                return @()
            }

            # Scan 1: garbled prefix with visible "-\d{2}-P" sub-number separator
            #   e.g. "7T4-10-P5EL", "714-10-PSL", "T714-10-POSR", "T14-10-PST-R"
            foreach ($m in [regex]::Matches($ocrText, '[A-Z0-9]{2,8}-\d{2}-\s*P([A-Z0-9]{1,3})(-[LR])?\b', 'IgnoreCase')) {
                $rawPfx = [regex]::Match($m.Value, '^[A-Z0-9]{2,8}-\d{2}').Value
                if ($knownJobPrefixes -contains $rawPfx) { continue }
                $rawSide = if ($m.Groups[2].Success) { $m.Groups[2].Value } else { '' }
                foreach ($kp in $knownJobPrefixes) {
                    foreach ($pn in (& $resolveSuffix $m.Groups[1].Value $rawSide $kp)) {
                        if (-not $extractedParts.Contains($pn)) {
                            [void]$extractedParts.Add($pn)
                            Write-Log "  OCR Job-Suffix (s1): '$($m.Value.Trim())' -> '$pn'" "SUCCESS"
                            & $trySetContext $pn $m.Value.Trim() $ocrText
                        }
                    }
                }
            }
            # Scan 2: sub-number merged/dropped into garbled prefix (no "-\d{2}-P" separator)
            #   e.g. "W7T4T1O-PSTL" where "17141-10" and "-" are all garbled together
            foreach ($m in [regex]::Matches($ocrText, '\b([A-Z][A-Z0-9]{3,8})-P([A-Z0-9]{2,3})\b', 'IgnoreCase')) {
                $pfx = $m.Groups[1].Value
                if (([regex]::Matches($pfx, '\d')).Count -lt 2) { continue }
                if ($knownJobPrefixes -contains $pfx) { continue }
                foreach ($kp in $knownJobPrefixes) {
                    foreach ($pn in (& $resolveSuffix $m.Groups[2].Value '' $kp)) {
                        if (-not $extractedParts.Contains($pn)) {
                            [void]$extractedParts.Add($pn)
                            Write-Log "  OCR Job-Suffix (s2): '$($m.Value.Trim())' -> '$pn'" "SUCCESS"
                            & $trySetContext $pn $m.Value.Trim() $ocrText
                        }
                    }
                }
            }
        }
    }

    # --- Post-processing: Remove bare 5-digit OCR noise ---
    # 1. Exact prefix: "25347" is prefix of found "25347-A02" â†’ assembly number, remove
    # 2. Fuzzy prefix: "26847" differs by <=2 digits from "25347" (prefix of "25347-A02") â†’ OCR misread, remove
    $dashParts = @($extractedParts | Where-Object { $_ -match '^\d{4,6}[-_]' })
    $dashPrefixes = @($dashParts | ForEach-Object { ($_ -split '[-_]')[0] })

    $toRemove = @()
    foreach ($pn in $extractedParts) {
        if ($pn -match '^\d{5}$') {
            # Check 1: exact prefix match
            $exactMatch = @($extractedParts | Where-Object { $_ -ne $pn -and $_.StartsWith($pn) })
            if ($exactMatch.Count -gt 0) {
                $toRemove += $pn
                Write-Log "  Filtering out '$pn' (assembly prefix of $($exactMatch -join ', '))" "WARN"
                continue
            }
            # Check 2: fuzzy match  -  within 2 digit differences of a dash-part prefix
            foreach ($prefix in $dashPrefixes) {
                if ($prefix.Length -eq $pn.Length) {
                    $diffs = 0
                    for ($ci = 0; $ci -lt $pn.Length; $ci++) {
                        if ($pn[$ci] -ne $prefix[$ci]) { $diffs++ }
                    }
                    if ($diffs -le 2) {
                        $toRemove += $pn
                        Write-Log "  Filtering out '$pn' (OCR misread of '$prefix', $diffs digit(s) off)" "WARN"
                        break
                    }
                }
            }
        }
    }
    foreach ($r in $toRemove) { [void]$extractedParts.Remove($r) }

    # --- Filter bare job-number prefix when full job parts are present ---
    # e.g. remove "17141-10" once "17141-10-P67" has been found
    $jobParts = @($extractedParts | Where-Object { $_ -match '^\d{5}-\d{2}-[A-Z]' })
    foreach ($jp in $jobParts) {
        $jobPrefix = [regex]::Match($jp, '^\d{5}-\d{2}').Value
        if ($extractedParts.Contains($jobPrefix)) {
            [void]$extractedParts.Remove($jobPrefix)
            Write-Log "  Filtering out '$jobPrefix' (job-number prefix; full parts like '$jp' found)" "WARN"
        }
    }

    # --- Normalize OCR-misread short family parts ---
    # Handles cases where the family prefix is misread or truncated:
    #   7036-018 -> 1035-018
    #   036-054  -> 1035-054
    # using dominant in-document family and PDF index validation.
    if (-not [string]::IsNullOrWhiteSpace($indexFolder)) {
        $normRows = Get-PdfIndexRowsCached -PdfIndexPath (Join-Path $indexFolder "pdf_index_clean.csv")
        if ($normRows.Count -gt 0) {
            $indexFamilySet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            foreach ($row in $normRows) {
                $bp = ""
                try { $bp = [string]$row.BasePart } catch { $bp = "" }
                if ([string]::IsNullOrWhiteSpace($bp)) { continue }
                $bp = $bp.Trim().ToUpperInvariant()
                if ($bp -match '^\d{4}-\d{3}$') { [void]$indexFamilySet.Add($bp) }
            }

            if ($indexFamilySet.Count -gt 0) {
                $familyPool = New-Object System.Collections.Generic.List[string]
                foreach ($pn in $extractedParts) {
                    if ($pn -match '^\d{4}-\d{3}$') { [void]$familyPool.Add($pn) }
                }
                foreach ($tok in $familyTokenCandidates) {
                    if ($tok -match '^\d{3,4}-\d{3}$') { [void]$familyPool.Add($tok) }
                }

                $prefixScores = @{}
                foreach ($fp in $familyPool) {
                    if ($fp -notmatch '^(?<pfx>\d{4})-(?<sfx>\d{3})$') { continue }
                    $pfx = [string]$Matches['pfx']
                    if (-not $prefixScores.ContainsKey($pfx)) { $prefixScores[$pfx] = 0 }
                    $score = if ($indexFamilySet.Contains($fp)) { 3 } else { 1 }
                    $prefixScores[$pfx] += $score
                }

                $dominantPrefix = ""
                if ($prefixScores.Count -gt 0) {
                    $dominantPrefix = (
                        $prefixScores.GetEnumerator() |
                        Sort-Object -Property Value -Descending |
                        Select-Object -First 1
                    ).Key
                }

                if (-not [string]::IsNullOrWhiteSpace($dominantPrefix)) {
                    $rawCandidates = New-Object System.Collections.Generic.List[string]
                    foreach ($pn in $extractedParts) {
                        if ($pn -match '^\d{3,4}-\d{3}$' -and -not $rawCandidates.Contains($pn)) { [void]$rawCandidates.Add($pn) }
                    }
                    foreach ($tok in $familyTokenCandidates) {
                        if ($tok -match '^\d{3,4}-\d{3}$' -and -not $rawCandidates.Contains($tok)) { [void]$rawCandidates.Add($tok) }
                    }

                    foreach ($raw in $rawCandidates) {
                        if ($raw -notmatch '^(?<rawPfx>\d{3,4})-(?<sfx>\d{3})$') { continue }
                        $rawPfx = [string]$Matches['rawPfx']
                        $sfx = [string]$Matches['sfx']
                        $fixed = "$dominantPrefix-$sfx"
                        if (-not $indexFamilySet.Contains($fixed)) { continue }

                        $isRawValid = ($rawPfx.Length -eq 4 -and $indexFamilySet.Contains($raw))
                        if ($fixed -eq $raw -and $isRawValid) { continue }

                        if (-not $extractedParts.Contains($fixed)) {
                            [void]$extractedParts.Add($fixed)
                            Write-Log "  OCR Family normalize: '$raw' -> '$fixed'" "SUCCESS"
                        }
                        if ($script:ocrPartContext -and $script:ocrPartContext.ContainsKey($raw) -and -not $script:ocrPartContext.ContainsKey($fixed)) {
                            $script:ocrPartContext[$fixed] = $script:ocrPartContext[$raw] -replace [regex]::Escape($raw), $fixed
                        } elseif ($script:ocrPartContext -and -not $script:ocrPartContext.ContainsKey($fixed) -and -not [string]::IsNullOrWhiteSpace($allOcrText)) {
                            & $trySetContext $fixed $raw $allOcrText
                            if (-not $script:ocrPartContext.ContainsKey($fixed)) {
                                & $trySetContext $fixed $fixed $allOcrText
                            }
                        }

                        if ($extractedParts.Contains($raw) -and -not $isRawValid) {
                            [void]$extractedParts.Remove($raw)
                            Write-Log "  OCR Family remove misread token '$raw' (normalized to '$fixed')" "WARN"
                        }
                    }
                }
            }
        }
    }

    # --- L/R pair completion ---
    # Liner plates come in L/R pairs from the same assembly.
    # If OCR found one side (e.g. P59-R recovered via alt correction), add the other side
    # if it exists in the PDF index. No guessing - just completing what the F80 has.
    if (-not [string]::IsNullOrWhiteSpace($indexFolder)) {
        $lrPdfRows = Get-PdfIndexRowsCached -PdfIndexPath (Join-Path $indexFolder "pdf_index_clean.csv")
        if ($lrPdfRows.Count -gt 0) {
            $lrParts = @($extractedParts | Where-Object { $_ -match '^\d{5}-\d{2}-P\d{2,3}-[LR]$' })
            foreach ($lrp in $lrParts) {
                $other = if ($lrp -match '-L$') { $lrp -replace '-L$','-R' } else { $lrp -replace '-R$','-L' }
                if ($extractedParts.Contains($other)) { continue }
                # Confirm the base plate is in the PDF index before adding the paired side
                $base = $lrp -replace '-[LR]$',''
                if ($lrPdfRows | Where-Object { $_.BasePart -eq $base }) {
                    [void]$extractedParts.Add($other)
                    Write-Log "  OCR L/R-Pair: '$other' added (paired with '$lrp' from F80)" "SUCCESS"
                    # Inherit description context from the found side (same part, other hand)
                    if ($script:ocrPartContext -and $script:ocrPartContext.ContainsKey($lrp) -and -not $script:ocrPartContext.ContainsKey($other)) {
                        $script:ocrPartContext[$other] = $script:ocrPartContext[$lrp] -replace [regex]::Escape($lrp), $other
                        Write-Log "  OCR Context (L/R inherit): $($script:ocrPartContext[$other])" "INFO"
                    }
                }
            }
        }
    }

    return $extractedParts.ToArray()
}

function Convert-VectorToPng {
    param([string]$VectorPath)
    try {
        Add-Type -AssemblyName System.Drawing
        $meta = New-Object System.Drawing.Imaging.Metafile($VectorPath)
        $w = [Math]::Max($meta.Width,  1)
        $h = [Math]::Max($meta.Height, 1)
        $bmp = New-Object System.Drawing.Bitmap($w, $h)
        $g   = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::White)
        $g.DrawImage($meta, 0, 0, $w, $h)
        $g.Dispose(); $meta.Dispose()
        $pngPath = [System.IO.Path]::ChangeExtension($VectorPath, '.converted.png')
        $bmp.Save($pngPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
        return $pngPath
    } catch { return $null }
}

function Get-PartNumbersFromDocxImages {
    param ([object]$doc, [string]$DocxPath)
    Write-Log "Pivoting to Word Image Extraction (ZIP method)..." "INFO"

    # A .docx is a ZIP archive. Images live in word/media/  -  extract them directly,
    # no Word SaveAs needed, no dialogs, no hangs.
    $tempDir = Join-Path $env:TEMP ("NMT_IMG_" + [guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    $convertedPngs = [System.Collections.Generic.List[string]]::new()
    $tempDocxCopy = ""
    $zip = $null

    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem

        $zipSourcePath = $DocxPath
        try {
            $zip = [System.IO.Compression.ZipFile]::OpenRead($zipSourcePath)
        } catch {
            $openMsg = $_.Exception.Message
            Write-Log "Primary DOCX ZIP open failed: $openMsg" "WARN"

            $tempDocxCopy = Join-Path $tempDir ("source_copy_" + [System.IO.Path]::GetFileName($DocxPath))
            $copied = $false
            try {
                Copy-Item -LiteralPath $DocxPath -Destination $tempDocxCopy -Force -ErrorAction Stop
                if (Test-Path $tempDocxCopy) {
                    $copied = $true
                    Write-Log "Using copied DOCX for OCR ZIP read: $tempDocxCopy" "INFO"
                }
            } catch {
                Write-Log "Direct DOCX copy failed: $($_.Exception.Message)" "WARN"
            }

            if (-not $copied -and $doc) {
                try {
                    # Save a shadow copy from the already-open Word document without altering source.
                    $doc.SaveCopyAs($tempDocxCopy)
                    if (Test-Path $tempDocxCopy) {
                        $copied = $true
                        Write-Log "Using Word SaveCopyAs DOCX for OCR ZIP read: $tempDocxCopy" "INFO"
                    }
                } catch {
                    Write-Log "Word SaveCopyAs fallback failed: $($_.Exception.Message)" "WARN"
                }
            }

            if (-not $copied) { throw }
            $zipSourcePath = $tempDocxCopy
            $zip = [System.IO.Compression.ZipFile]::OpenRead($zipSourcePath)
        }

        $mediaEntries = $zip.Entries | Where-Object {
            $_.FullName -match '^word/media/' -and
            $_.Name     -match '\.(png|jpg|jpeg|gif|bmp|emf|wmf)$'
        }

        $extracted = [System.Collections.Generic.List[string]]::new()
        foreach ($entry in $mediaEntries) {
            $dest = Join-Path $tempDir $entry.Name
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, $dest, $true)
            $extracted.Add($dest)
        }
        $zip.Dispose()
        $zip = $null

        Write-Log "Extracted $($extracted.Count) image(s) from DOCX archive." "INFO"

        # Raster files  -  OCR directly
        $rasterPaths = @($extracted | Where-Object { $_ -match '\.(png|jpg|jpeg|gif|bmp)$' })

        # Vector files  -  render to PNG via System.Drawing first
        $vectorPaths = @($extracted | Where-Object { $_ -match '\.(emf|wmf)$' })
        foreach ($v in $vectorPaths) {
            $png = Convert-VectorToPng -VectorPath $v
            if ($png) { $convertedPngs.Add($png) }
        }

        $allPaths = $rasterPaths + $convertedPngs.ToArray()
        Write-Log "Images ready for OCR: $($rasterPaths.Count) raster + $($vectorPaths.Count) vector ($($convertedPngs.Count) converted)." "INFO"

        if ($allPaths.Count -eq 0) {
            Write-Log "No images found in DOCX archive." "WARN"
            return @()
        }

        return (Get-PartNumbersFromImages -ImagePaths $allPaths)

    } catch {
        if ($zip) {
            try { $zip.Dispose() } catch { }
            $zip = $null
        }
        Write-Log "DOCX image extraction failed: $($_.Exception.Message)" "ERROR"
        return @()
    } finally {
        foreach ($p in $convertedPngs) {
            if (Test-Path $p) { Remove-Item $p -Force -ErrorAction SilentlyContinue }
        }
        if (-not [string]::IsNullOrWhiteSpace($tempDocxCopy) -and (Test-Path $tempDocxCopy)) {
            Remove-Item $tempDocxCopy -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

# ==============================================================================
#  PDM Email Parsing
# ==============================================================================

function Get-DocxPathFromEmail {
    param([string]$EmailBody, [string]$EmailSubject)

    $patterns = @(
        '(?mi)(?:file|document|path)\s*[:=]\s*(.+\.docx)',
        '(?mi)(\\[^\r\n]+\.docx)',
        '(?mi)([A-Z]:\\[^\r\n]+\.docx)',
        '(?mi)([\w].*?[- - ]\s*F80[a]?\.docx)',
        '(?mi)(?:"|''|`)(.*?\.docx)(?:"|''|`)'
    )

    foreach ($pattern in $patterns) {
        $match = [regex]::Match($EmailBody, $pattern)
        if ($match.Success) {
            $rawPath = $match.Groups[1].Value.Trim()
            $rawPath = $rawPath -replace '^["''`]+|["''`]+$', ''
            return $rawPath.Trim()
        }
    }

    $subjectMatch = [regex]::Match($EmailSubject, '(\\[^\r\n]+\.docx|[A-Z]:\\[^\r\n]+\.docx)')
    if ($subjectMatch.Success) {
        return $subjectMatch.Groups[1].Value.Trim()
    }

    return $null
}

function Sync-PdmFile {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    # Only try PDM if path is inside vault root
    if ($Path -notmatch [regex]::Escape($pdmVaultPath)) { return $false }
    
    try {
        $vault = New-Object -ComObject "ConisioLib.EdmVault"
        # NMT_PDM is the vault name from PDFIndexManager.ps1
        $vault.LoginAuto("NMT_PDM", 0)
        if ($vault.IsLoggedIn) {
            $fileName = Split-Path $Path -Leaf
            Write-Log "  PDM: Searching for filename '$fileName'..." "INFO"
            $search = $vault.CreateSearch()
            $search.FileName = $fileName
            $search.Recursive = $true
            $pos = $search.GetFirstResult()
            if (-not $pos) {
                Write-Log "  PDM: File '$fileName' not found in vault search." "WARN"
                return $false
            }
            while ($pos) {
                $vPath = $pos.Path
                if ($vPath -eq $Path -or $vPath.ToLower() -eq $Path.ToLower()) {
                    $file = $vault.GetObject(1, $pos.ID)
                    if ($file) {
                        Write-Log "  PDM: Match found! ID:$($pos.ID) Path:$vPath" "SUCCESS"
                        Write-Log "  PDM: Syncing (Get Latest)..." "INFO"
                        $file.GetFileCopy(0, "", 0, $null)
                        return $true
                    }
                } else {
                    Write-Log "  PDM: Found similar file but path mismatch: '$vPath' vs '$Path'" "WARN"
                }
                $pos = $search.GetNextResult()
            }
        } else {
            Write-Log "  PDM: Vault login failed (NMT_PDM)" "ERROR"
        }
    } catch {
        Write-Log "  PDM Sync error for '$Path': $($_.Exception.Message)" "WARN"
    }
    return $false
}

function Resolve-DocxPath {
    param([string]$RawPath)

    if ($RawPath -match '^[A-Z]:\\') {
        [void](Sync-PdmFile -Path $RawPath)
        return $RawPath
    }

    $fullPath = Join-Path $pdmVaultPath $RawPath.TrimStart('\')
    if (Test-Path $fullPath) { return $fullPath }
    
    # Try PDM Sync on the initial path before searching
    if (Sync-PdmFile -Path $fullPath) {
        if (Test-Path $fullPath) { return $fullPath }
    }

    $fileName = Split-Path $RawPath -Leaf
    $searchPaths = @(
        (Join-Path $pdmVaultPath "10 - Orders\$fileName"),
        (Join-Path $pdmVaultPath "Orders\$fileName"),
        (Join-Path $pdmVaultPath $fileName)
    )
    foreach ($sp in $searchPaths) {
        if (Test-Path $sp) { return $sp }
        if (Sync-PdmFile -Path $sp) {
            if (Test-Path $sp) { return $sp }
        }
    }

    # Recursive search fallback (bounded to 10 - Orders for speed)
    Write-Log "  Path not found on disk; searching for '$fileName' under vault roots..." "WARN"
    foreach ($root in @("10 - Orders", "Orders", "")) {
        $searchDir = Join-Path $pdmVaultPath $root
        if (Test-Path $searchDir) {
            $found = Get-ChildItem -Path $searchDir -Filter $fileName -Recurse -File -ErrorAction SilentlyContinue | 
                     Select-Object -First 1
            if ($found) {
                Write-Log "  File found via disk search: $($found.FullName)" "SUCCESS"
                return $found.FullName
            }
        }
    }

    # Final attempt: PDM Vault Search if local searches failed
    try {
        Write-Log "  File not found on disk; performing PDM Vault Search for '$fileName'..." "INFO"
        $vault = New-Object -ComObject "ConisioLib.EdmVault"
        $vault.LoginAuto("NMT_PDM", 0)
        if ($vault.IsLoggedIn) {
            $search = $vault.CreateSearch()
            $search.FileName = $fileName
            $search.Recursive = $true
            $search.FindFiles = $true
            $pos = $search.GetFirstResult()
            if ($pos) {
                $pdmPath = $pos.Path
                Write-Log "  PDM Search Found: $pdmPath" "SUCCESS"
                [void](Sync-PdmFile -Path $pdmPath)
                if (Test-Path $pdmPath) { return $pdmPath }
            }
        }
    } catch {
        Write-Log "  PDM Search failed: $($_.Exception.Message)" "WARN"
    }

    return $fullPath
}

# ==============================================================================
#  Word Document Parsing
# ==============================================================================

function Read-OrderFromWordDoc {
    param([string]$DocxPath)

    $result = @{
        PartNumbers  = [System.Collections.Generic.List[string]]::new()
        OrderLines   = [System.Collections.Generic.List[hashtable]]::new()
        ContractInfo = @{}
        JobNumber    = ""; ClientName = ""; Success = $false; Error = ""
    }

    if (-not (Test-Path $DocxPath)) {
        $result.Error = "Document not found: $DocxPath"
        return $result
    }

    $word = $null; $doc = $null
    try {
        try {
            $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
        } catch {
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
        }

        $doc = $word.Documents.Open($DocxPath, $false, $true)
        
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($DocxPath)
        if ($fileName -match '^(\d{4,6})') { $result.JobNumber = $Matches[1] }
        if ($fileName -match '^\d{4,6}\s*[- - ]\s*(.+?)\s*[- - ]') { $result.ClientName = $Matches[1].Trim() }

        $tableCount = $doc.Tables.Count
        Write-Log "Document contains $tableCount table(s)" "INFO"

        for ($t = $tableCount; $t -ge 1; $t--) {
            $table = $doc.Tables.Item($t)
            $rowCount = $table.Rows.Count
            if ($rowCount -lt 1) { continue }
            
            Write-Log "Scanning Table ${t} ($rowCount rows)..." "INFO"
            
            $partCol = -1; $startRow = -1
            $revCol  = -1; $descCol = -1; $qtyCol = -1   # for Rev / Description / Qty columns
            $maxHeaderScan = [math]::Min(10, $rowCount)

            for ($rIdx = 1; $rIdx -le $maxHeaderScan; $rIdx++) {
                $rowArr = @()
                try {
                    $cells = $table.Rows.Item($rIdx).Cells
                    for ($c = 1; $c -le $cells.Count; $c++) {
                        $rowArr += ($cells.Item($c).Range.Text -replace '[\r\n\a\x07]', '').Trim()
                    }
                } catch { }

                $joined = ($rowArr -join " | ").ToUpper()
                if ($rIdx -le 2) { Write-Log "  Row ${rIdx}: $joined" "INFO" }

                for ($c = 0; $c -lt $rowArr.Count; $c++) {
                    $h = [string]$rowArr[$c].ToUpper()
                    if     ($h -match 'PART$|PART\s*NO|PART\s*NUM|PART\s*#|NMT\s*PART') { $partCol = $c + 1 }
                    elseif ($h -match '^REV(ISION)?$|^REV\s*[#.]?$')                    { $revCol  = $c + 1 }
                    elseif ($h -match 'DESCRIPTION$|PART\s*DESC|^DESC$')                 { $descCol = $c + 1 }
                    elseif ($h -match '^QTY$|^QUANT|^ORDER\s*Q|^ORD\s*Q')                { $qtyCol  = $c + 1 }
                }
                if ($partCol -ge 1) { $startRow = $rIdx + 1; break }
            }

            # Brute Force
            if ($partCol -lt 0) {
                for ($r = 1; $r -le $maxHeaderScan; $r++) {
                    try {
                        $cells = $table.Rows.Item($r).Cells
                        for ($c = 1; $c -le $cells.Count; $c++) {
                            $cellTxt = ($cells.Item($c).Range.Text -replace '[\r\n\a\x07]','').Trim()
                            # Only test short cells - part numbers are never sentences
                            if ($cellTxt.Length -le 40 -and $cellTxt -match '\d{5}-\d{2}-[A-Z]\d{2,3}|\d{4,6}[-_][A-Z0-9]{1,6}|\d{5}[A-Z]\d{3,4}|[A-Z]{2,5}-\d+[.,]\d') {
                                $partCol = $c; $startRow = $r; break
                            }
                        }
                    } catch { }
                    if ($partCol -ge 1) { 
                        Write-Log "  Table ${t} identified via Pattern Match in Col $partCol" "SUCCESS"
                        break 
                    }
                }
            }

            if ($partCol -ge 1) {
                for ($r = $startRow; $r -le $rowCount; $r++) {
                    try {
                        $pn = ($table.Cell($r, $partCol).Range.Text -replace '[\r\n\a\x07]', '').Trim().ToUpper()
                        # Allow suffixes like -A or -1 in part numbers, and handle 4-6 digit prefixes
                        if ($pn.Length -le 40 -and $pn -match '\d{5}-\d{2}-[A-Z]\d{2,3}(?:[-_][A-Z0-9]+)?|\d{4,6}[-_][A-Z0-9]{1,10}(?:[-_][A-Z0-9]{1,10})*|[A-Z]{1,6}\d{1,4}(?:-[A-Z0-9]{1,6}){2,8}|\d{5}[A-Z]\d{3,4}|[A-Z]{2,5}-\d+[.,]\d|[A-Z]\d{3,5}[A-Z]|\d{5}' -and -not $result.PartNumbers.Contains($pn)) {
                            # Pull Rev / Description / Qty from their detected columns
                            $rev = ""; $desc = ""; $qty = ""
                            if ($revCol  -gt 0) { try { $rev  = ($table.Cell($r,$revCol ).Range.Text -replace '[\r\n\a\x07]','').Trim() } catch {} }
                            if ($descCol -gt 0) { try { $desc = ($table.Cell($r,$descCol).Range.Text -replace '[\r\n\a\x07]','').Trim() } catch {} }
                            if ($qtyCol  -gt 0) { try { $qty  = ($table.Cell($r,$qtyCol ).Range.Text -replace '[\r\n\a\x07]','').Trim() } catch {} }
                            # Build display line: "20045-10-A11 Rev.2 Rope Guide Qty: 4"
                            $lineDesc = $pn
                            if ($rev  -ne "" -and $rev -ne "0") { $lineDesc += " Rev.$rev" }
                            if ($desc -ne "") { $lineDesc += " $($desc -replace '\s+',' ')" }
                            if ($qty  -ne "") {
                                # Normalise "4,00 EA" -> "4", "24.00 EA" -> "24"
                                $qn = ($qty -replace ',','.') -replace '\.0+\s*(EA.*)?$','$1' -replace '\s*EA\s*$','' -replace '(\d+)\.\d+$','$1'
                                $lineDesc += " Qty: $qn"
                            }
                            $result.PartNumbers.Add($pn)
                            $result.OrderLines.Add(@{ Part=$pn; Description=$lineDesc; Rev=$rev; Qty=$qty })
                            Write-Log "  Found Part: $pn (Rev: $rev, Qty: $qty)" "SUCCESS"
                        }
                    } catch { }
                }
                if ($result.PartNumbers.Count -gt 0) { break }
            }
        }

        # ---- Deep scan: Project Overview free text & nested tables ----
        # ALWAYS run this scan - the Project Overview often has free-text part descriptions
        # that the structured table scan above misses (e.g. "1204-a141 - Brake Pad Assy").
        {
            # Part-number patterns safe for text extraction (no bare \d{5}  -  too many false positives)
            $pnPatterns = @(
                '\b\d{5}-\d{2}-[A-Z]\d{2,3}(?:-[A-Z])?\b',
                '\b[A-Z]{2,5}-\d+[.,]\d{1,3}(?:[-x][A-Z0-9.]+){1,4}\b',
                '\b[A-Z]{1,4}\d{3,5}[A-Z]{1,2}\b',
                '\b[A-Z]{1,6}\d{1,4}(?:-[A-Z0-9]{1,6}){1,8}\b',
                '\b\d{4,6}[-_][A-Z0-9]{1,8}(?:[-_][A-Z0-9]{1,8}){0,3}\b',
                '\b\d{5}[A-Z]\d{3,4}\b',
                '\b[A-Z]{2,4}\d{4,6}(?:-[A-Z]{1,3})?\b',
                '\b\d{2}-\d{2}-\d{2,3}(?:-[A-Z0-9]{1,4}){0,3}\b',
                '\b[A-Z]{1,3}\d{1,3}[A-Z]\d{1,3}(?:-[A-Z]{1,3})?\b',
                '\b[A-Z]{2,4}-\d{2,3}-[A-Z]\d{1,3}(?:-[A-Z]{1,3})?\b',
                '\b[A-Z]{2,4}\d+\.\d+[Xx]\d+\.?\d*[A-Z]?\b'
            )

            # 1. Scan Project Overview table's Range.Text (includes nested table text + free-text paragraphs)
            for ($t = 1; $t -le $tableCount; $t++) {
                try {
                    $tbl = $doc.Tables.Item($t)
                    $r1 = ($tbl.Range.Text -replace '[\r\n\a\x07]',' ').Trim().ToUpper()
                    if ($r1 -match 'PROJECT\s*OVERVIEW') {
                        $rangeText = $tbl.Range.Text -replace '[\r\n\a\x07]', ' '
                        Write-Log "Deep scan: Project Overview range ($($rangeText.Length) chars)..." "INFO"

                        # Check for nested tables (parts grid inside this table)
                        $nestedCount = $tbl.Range.Tables.Count
                        if ($nestedCount -gt 1) {
                            Write-Log "  Found $($nestedCount - 1) nested table(s) inside Project Overview" "INFO"
                            for ($nt = 2; $nt -le $nestedCount; $nt++) {
                                $nested = $tbl.Range.Tables.Item($nt)
                                $nRows = $nested.Rows.Count
                                Write-Log "  Scanning nested table ($nRows rows)..." "INFO"
                                for ($nr = 1; $nr -le $nRows; $nr++) {
                                    try {
                                        $nCells = $nested.Rows.Item($nr).Cells
                                        for ($nc = 1; $nc -le $nCells.Count; $nc++) {
                                            $ct = ($nCells.Item($nc).Range.Text -replace '[\r\n\a\x07]','').Trim().ToUpper()
                                            foreach ($pp in $pnPatterns) {
                                                if ($ct -match $pp) {
                                                    $pn = ([regex]::Match($ct, $pp)).Value
                                                    if (-not $result.PartNumbers.Contains($pn)) {
                                                        $result.PartNumbers.Add($pn)
                                                        $result.OrderLines.Add(@{ Part=$pn; Description="Nested Table Row $nr" })
                                                        Write-Log "  Found Part: $pn" "SUCCESS"
                                                    }
                                                }
                                            }
                                        }
                                    } catch { }
                                }
                            }
                        }

                        # Always regex the full range text to catch free-text part descriptions
                        foreach ($pp in $pnPatterns) {
                            foreach ($m in [regex]::Matches($rangeText, $pp, 'IgnoreCase')) {
                                $pn = $m.Value.ToUpper() -replace '\s+', ' '
                                if ($pn -match '^[\d\-_]+$' -and $pn -notmatch '^\d{4}-\d{3}$') { continue }
                                if ($pn -match '^RAL\s|^HEAD\s|^EA\s|^INO\b|^CRATE\b|^PRIME|^GRADE|^ASSY\s|^GR\d|^T1X|^THRD\s|^V2-|^I2-') { continue }
                                if (-not $result.PartNumbers.Contains($pn)) {
                                    $result.PartNumbers.Add($pn)
                                    $result.OrderLines.Add(@{ Part=$pn; Description="Project Overview Text" })
                                    Write-Log "  Found Part: $pn (Project Overview)" "SUCCESS"
                                }
                            }
                        }
                        break
                    }
                } catch { }
            }

            # 2. Scan shapes / text boxes (tables inside floating frames)
            if ($result.PartNumbers.Count -eq 0) {
                try {
                    foreach ($shape in $doc.Shapes) {
                        try {
                            if (-not $shape.TextFrame -or -not $shape.TextFrame.HasText) { continue }
                            $st = $shape.TextFrame.TextRange.Text
                            if ($st -notmatch '(?i)Part|MfgJobType') { continue }
                            Write-Log "Deep scan: Shape '$($shape.Name)' contains parts data" "INFO"
                            foreach ($pp in $pnPatterns) {
                                foreach ($m in [regex]::Matches($st, $pp)) {
                                    $pn = $m.Value.ToUpper()
                                    if (-not $result.PartNumbers.Contains($pn)) {
                                        $result.PartNumbers.Add($pn)
                                        $result.OrderLines.Add(@{ Part=$pn; Description="Shape" })
                                        Write-Log "  Found Part: $pn" "SUCCESS"
                                    }
                                }
                            }
                        } catch { }
                    }
                } catch { }
            }

            # 3. Catch-all: scan every table's full range text for part numbers
            if ($result.PartNumbers.Count -eq 0) {
                Write-Log "Deep scan catch-all: scanning all $tableCount table(s) range text..." "INFO"
                for ($tc = 1; $tc -le $tableCount; $tc++) {
                    try {
                        $rt = ($doc.Tables.Item($tc).Range.Text -replace '[\r\n\a\x07]', ' ')
                        foreach ($pp in $pnPatterns) {
                            foreach ($m in [regex]::Matches($rt, $pp)) {
                                $pn = $m.Value.ToUpper() -replace '\s+', ''
                                if ($pn -match '^[\d\-_]+$' -and $pn -notmatch '^\d{4}-\d{3}$') { continue }
                                if (-not $result.PartNumbers.Contains($pn)) {
                                    $result.PartNumbers.Add($pn)
                                    $result.OrderLines.Add(@{ Part=$pn; Description="Table $tc Text" })
                                    Write-Log "  Catch-all found: $pn (table $tc)" "SUCCESS"
                                }
                            }
                        }
                    } catch { }
                }
            }
        }

        # ==== EXTRACTION CHAIN: Tables (above) -> Image OCR -> Text Fallbacks ====

        # --- METHOD 1: OCR on embedded images (parts table is a picture) ---
        if ($result.PartNumbers.Count -eq 0) {
            Write-Log "=== IMAGE OCR EXTRACTION ===" "WARN"
            $ocrParts = Get-PartNumbersFromDocxImages -doc $doc -DocxPath $DocxPath
            foreach ($pn in $ocrParts) {
                if (-not $result.PartNumbers.Contains($pn)) {
                    $result.PartNumbers.Add($pn)
                    # Use full context (Part Rev.N Description Qty: N) if OCR extraction captured it
                    $ocrDesc = if ($script:ocrPartContext -and $script:ocrPartContext.ContainsKey($pn)) {
                        $script:ocrPartContext[$pn]
                    } else { $pn }
                    $result.OrderLines.Add(@{ Part=$pn; Description=$ocrDesc })
                }
            }

            $contextHint = "JOB:$($result.JobNumber) CLIENT:$($result.ClientName) FILE:$([System.IO.Path]::GetFileNameWithoutExtension($DocxPath)) OCR:$($script:ocrInferenceText)"
            $drawingParts = @($result.PartNumbers | Where-Object { Test-DrawingLikePartNumber -PartNumber $_ })
            $hardwareParts = @($result.PartNumbers | Where-Object { Test-HardwareLikePartNumber -PartNumber $_ })

            if ($drawingParts.Count -eq 0 -and $result.PartNumbers.Count -gt 0) {
                $inferred = @(Infer-DrawingPartFromIndexContext -ContextText $contextHint -ExistingParts $result.PartNumbers.ToArray())
                foreach ($ip in $inferred) {
                    if (-not $result.PartNumbers.Contains($ip)) {
                        $result.PartNumbers.Add($ip)
                        $result.OrderLines.Add(@{ Part=$ip; Description=$ip })
                    }
                }
            }

            $drawingParts = @($result.PartNumbers | Where-Object { Test-DrawingLikePartNumber -PartNumber $_ })
            $hardwareParts = @($result.PartNumbers | Where-Object { Test-HardwareLikePartNumber -PartNumber $_ })

            if ($drawingParts.Count -gt 0 -and $hardwareParts.Count -gt 0) {
                $filteredParts = [System.Collections.Generic.List[string]]::new()
                foreach ($pn in $result.PartNumbers) {
                    if (-not (Test-HardwareLikePartNumber -PartNumber $pn)) {
                        if (-not $filteredParts.Contains($pn)) { $filteredParts.Add($pn) }
                    }
                }
                $filteredLines = [System.Collections.Generic.List[hashtable]]::new()
                foreach ($ol in $result.OrderLines) {
                    $op = ""
                    try { $op = [string]$ol.Part } catch { }
                    if (-not (Test-HardwareLikePartNumber -PartNumber $op)) {
                        $filteredLines.Add($ol)
                    }
                }
                Write-Log "  OCR cleanup: removed $($hardwareParts.Count) hardware-like part(s); keeping $($filteredParts.Count) drawing-like part(s)." "WARN"
                $result.PartNumbers = $filteredParts
                $result.OrderLines = $filteredLines
            }

            $drawingParts = @($result.PartNumbers | Where-Object { Test-DrawingLikePartNumber -PartNumber $_ })
            $hardwareParts = @($result.PartNumbers | Where-Object { Test-HardwareLikePartNumber -PartNumber $_ })
            if ($drawingParts.Count -eq 0 -and $hardwareParts.Count -gt 0) {
                Write-Log "  OCR produced hardware-only parts and no reliable drawing match; suppressing to avoid bad transmittal." "WARN"
                $result.PartNumbers = [System.Collections.Generic.List[string]]::new()
                $result.OrderLines = [System.Collections.Generic.List[hashtable]]::new()
            }

            Write-Log "  OCR extraction: $($result.PartNumbers.Count) parts found" "INFO"
        }

        # --- METHOD 2: Full Text (fallback for documents with text-based parts) ---
        if ($result.PartNumbers.Count -eq 0) {
            Write-Log "=== FULL TEXT FALLBACK ===" "WARN"
            $rawText = $doc.Range().Text
            $stripped = ($rawText -replace '[\r\n\a\x07\x0B\x0C\x01\x13\x14\x15]', '')
            $stripped = ($stripped -replace '[\u2013\u2014\u2012]', '-').ToUpper()
            $txtPatterns = @(
                '\b\d{5}-\d{2}-[A-Z]\d{2,3}(?:-[A-Z])?\b',
                '\b[A-Z]{2,5}-\d+[.,]\d{1,3}(?:[-x][A-Z0-9.]+){1,4}\b',
                '\b[A-Z]{1,6}\d{1,4}(?:-[A-Z0-9]{1,6}){1,8}\b',
                '\b\d{4,6}[-_][A-Z0-9]{1,6}(?:[-_][A-Z0-9]{1,6}){0,3}\b',
                '\b\d{5}[A-Z]\d{3,4}\b',
                '\b[A-Z]{2,4}\d{3,5}[A-Z]{1,2}\b',
                '\b[A-Z]{2,4}\d{4,6}(?:-[A-Z]{1,3})?\b',
                '\b\d{2}-\d{2}-\d{2,3}(?:-[A-Z0-9]{1,4}){0,3}\b',
                '\b[A-Z]{1,3}\d{1,3}[A-Z]\d{1,3}(?:-[A-Z]{1,3})?\b'
            )
            foreach ($pp in $txtPatterns) {
                foreach ($m in [regex]::Matches($stripped, $pp)) {
                    $pn = $m.Value
                    if ($result.JobNumber -and $pn -eq $result.JobNumber) { continue }
                    if ($pn -match '^[\d\-_]+$' -and $pn -notmatch '^\d{4}-\d{3}$') { continue }
                    if (-not $result.PartNumbers.Contains($pn)) {
                        $result.PartNumbers.Add($pn)
                        $result.OrderLines.Add(@{ Part=$pn; Description="Full Text" })
                        Write-Log "  Found Part: $pn" "SUCCESS"
                    }
                }
            }
        }

        # Final cleanup: never allow hardware-only OCR/text hits to drive transmittal drawings.
        $finalContextHint = "JOB:$($result.JobNumber) CLIENT:$($result.ClientName) FILE:$([System.IO.Path]::GetFileNameWithoutExtension($DocxPath)) OCR:$($script:ocrInferenceText)"
        $finalDrawingParts = @($result.PartNumbers | Where-Object { Test-DrawingLikePartNumber -PartNumber $_ })
        $finalHardwareParts = @($result.PartNumbers | Where-Object { Test-HardwareLikePartNumber -PartNumber $_ })

        if ($finalDrawingParts.Count -eq 0 -and $result.PartNumbers.Count -gt 0) {
            $inferred2 = @(Infer-DrawingPartFromIndexContext -ContextText $finalContextHint -ExistingParts $result.PartNumbers.ToArray())
            foreach ($ip2 in $inferred2) {
                if (-not $result.PartNumbers.Contains($ip2)) {
                    $result.PartNumbers.Add($ip2)
                    $result.OrderLines.Add(@{ Part=$ip2; Description=$ip2 })
                }
            }
        }

        $finalDrawingParts = @($result.PartNumbers | Where-Object { Test-DrawingLikePartNumber -PartNumber $_ })
        $finalHardwareParts = @($result.PartNumbers | Where-Object { Test-HardwareLikePartNumber -PartNumber $_ })

        if ($finalDrawingParts.Count -gt 0 -and $finalHardwareParts.Count -gt 0) {
            $keepParts = [System.Collections.Generic.List[string]]::new()
            foreach ($pnKeep in $result.PartNumbers) {
                if (-not (Test-HardwareLikePartNumber -PartNumber $pnKeep)) {
                    if (-not $keepParts.Contains($pnKeep)) { $keepParts.Add($pnKeep) }
                }
            }
            $keepLines = [System.Collections.Generic.List[hashtable]]::new()
            foreach ($olKeep in $result.OrderLines) {
                $opKeep = ""
                try { $opKeep = [string]$olKeep.Part } catch { }
                if (-not (Test-HardwareLikePartNumber -PartNumber $opKeep)) {
                    $keepLines.Add($olKeep)
                }
            }
            Write-Log "Final cleanup: removed $($finalHardwareParts.Count) hardware-like part(s), kept $($keepParts.Count) drawing-like part(s)." "WARN"
            $result.PartNumbers = $keepParts
            $result.OrderLines = $keepLines
        }

        $finalDrawingParts = @($result.PartNumbers | Where-Object { Test-DrawingLikePartNumber -PartNumber $_ })
        $finalHardwareParts = @($result.PartNumbers | Where-Object { Test-HardwareLikePartNumber -PartNumber $_ })
        if ($finalDrawingParts.Count -eq 0 -and $finalHardwareParts.Count -gt 0) {
            Write-Log "Final cleanup: only hardware-like parts remain; suppressing parts list to avoid incorrect transmittal." "WARN"
            $result.PartNumbers = [System.Collections.Generic.List[string]]::new()
            $result.OrderLines = [System.Collections.Generic.List[hashtable]]::new()
        }

        $result.Success = $result.PartNumbers.Count -gt 0
        if (-not $result.Success) { Write-Log "No parts found by any method." "WARN" }
    } catch {
        $result.Error = $_.Exception.Message
        Write-Log "Word Error: $($result.Error)" "ERROR"
    } finally {
        if ($doc) { $doc.Close($false); [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null }
        if ($word) { [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
    }
    return $result
}

function Extract-PartNumbers {
    param([string]$EmailBody, [string]$EmailSubject)
    $parts = New-Object System.Collections.Generic.List[string]
    $patterns = @('\b\d{4,6}[-_][A-Z0-9]{1,6}\b', '[A-Z]{2,4}[-_]\d{4,6}')
    foreach ($p in $patterns) {
        foreach ($m in [regex]::Matches("$EmailSubject`n$EmailBody", $p)) {
            $pn = $m.Value.ToUpper()
            if (-not $parts.Contains($pn)) { [void]$parts.Add($pn) }
        }
    }
    return $parts.ToArray()
}

function Get-PartNumbersFromEmailAttachments {
    param([object]$MailItem)
    Write-Log "Scanning email attachments for images..." "INFO"
    $extractedParts = @()
    $tempDir = Join-Path $env:TEMP ("NMT_MAIL_" + [guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    try {
        $imgPaths = @()
        foreach ($att in $MailItem.Attachments) {
            if ($att.FileName -match '\.(png|jpg|jpeg|gif|bmp)$') {
                $savePath = Join-Path $tempDir $att.FileName
                $att.SaveAsFile($savePath)
                $imgPaths += $savePath
            }
        }
        if ($imgPaths.Count -gt 0) { $extractedParts = Get-PartNumbersFromImages -ImagePaths $imgPaths }
    } finally {
        if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
    return $extractedParts
}

# ==============================================================================
#  Assembly BOM Expansion - Find all sub-parts for NMT assemblies
# ==============================================================================

function Get-FileBaseNameUpper {
    param([string]$PathOrName)
    if ([string]::IsNullOrWhiteSpace($PathOrName)) { return "" }
    $leaf = $PathOrName
    try {
        if ($PathOrName.IndexOf('\') -ge 0 -or $PathOrName.IndexOf('/') -ge 0) {
            $leaf = Split-Path -Path $PathOrName -Leaf
        }
    } catch { }
    if ([string]::IsNullOrWhiteSpace($leaf)) { $leaf = $PathOrName }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
    if ([string]::IsNullOrWhiteSpace($base)) { $base = $leaf }
    # REMOVED: stripping instance suffix like -1 because it hits -A or -1 in real parts
    return $base.Trim().ToUpper()
}

function Get-SwComponentModelPath {
    param([object]$Component)
    $modelPath = ""
    if ($null -eq $Component) { return $modelPath }
    try { $modelPath = $Component.GetPathName() } catch { }
    if ([string]::IsNullOrWhiteSpace($modelPath)) {
        try {
            $childDoc = $Component.GetModelDoc2()
            if ($childDoc) { $modelPath = $childDoc.GetPathName() }
        } catch { }
    }
    if ([string]::IsNullOrWhiteSpace($modelPath)) {
        try { $modelPath = $Component.Name2 } catch { }
    }
    return $modelPath
}

function Get-SwComponentPartNumber {
    param([object]$Component)
    $partNum = ""
    if ($null -eq $Component) { return $partNum }

    # Try custom properties first.
    try {
        $swModelDoc = $Component.GetModelDoc2()
        if ($swModelDoc) {
            $custPropMgr = $null
            try { $custPropMgr = $swModelDoc.Extension.CustomPropertyManager("") } catch { }
            if ($custPropMgr) {
                $propNames = @("Part Number", "PartNumber", "PartNo", "Part No", "DrawingNo", "Drawing No", "Number")
                foreach ($propName in $propNames) {
                    $propValue = ""
                    $resolvedValue = ""
                    try { $null = $custPropMgr.Get4($propName, $false, [ref]$propValue, [ref]$resolvedValue) } catch { }
                    if (-not [string]::IsNullOrWhiteSpace($resolvedValue)) {
                        $partNum = $resolvedValue
                        break
                    }
                    if (-not [string]::IsNullOrWhiteSpace($propValue)) {
                        $partNum = $propValue
                        break
                    }
                }
            }
        }
    } catch { }

    # Fallback to model file name.
    if ([string]::IsNullOrWhiteSpace($partNum)) {
        $modelPath = Get-SwComponentModelPath -Component $Component
        if (-not [string]::IsNullOrWhiteSpace($modelPath)) {
            $partNum = Get-FileBaseNameUpper -PathOrName $modelPath
        }
    }

    return $partNum.Trim().ToUpper()
}

function Collect-SolidWorksPartNumbersFromComponent {
    param(
        [object]$Component,
        [hashtable]$PartSet,
        [hashtable]$SeenModelPaths
    )

    if ($null -eq $Component) { return }

    $isSuppressed = $false
    try { $isSuppressed = [bool]$Component.IsSuppressed() } catch { }
    if ($isSuppressed) { return }

    $modelPath = Get-SwComponentModelPath -Component $Component
    if (-not [string]::IsNullOrWhiteSpace($modelPath)) {
        $cfgName = ""
        try { $cfgName = [string]$Component.ReferencedConfiguration } catch { }
        $mpKey = ($modelPath + "|" + $cfgName).ToLowerInvariant()
        if ($SeenModelPaths.ContainsKey($mpKey)) { return }
        $SeenModelPaths[$mpKey] = $true
    }

    $partNum = Get-SwComponentPartNumber -Component $Component
    if (-not [string]::IsNullOrWhiteSpace($partNum)) {
        if ($partNum -notmatch '(?i)LOAD[\s_]?CERT|SCOPE|MANUAL') {
            $PartSet[$partNum] = $true
        }
    }

    # Resolve this component if it's a sub-assembly so GetChildren() works at depth
    try {
        $compDoc = $Component.GetModelDoc2()
        if ($compDoc -and $compDoc.GetType() -eq 2) {
            $null = $compDoc.ResolveAllLightWeightComponents($true)
        }
    } catch { }

    $children = $null
    try { $children = $Component.GetChildren() } catch { }
    if ($null -eq $children) { return }

    foreach ($child in @($children)) {
        if ($null -ne $child) {
            Collect-SolidWorksPartNumbersFromComponent -Component $child -PartSet $PartSet -SeenModelPaths $SeenModelPaths
        }
    }
}

function Get-JobFoldersFromContext {
    param(
        [string]$OrderDocPath,
        [string]$JobNumber,
        [string[]]$CrawlRoots
    )

    $candidates = @{}

    function Add-CandidatePath {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return }
        if (-not (Test-Path $Path)) { return }
        $norm = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
        if (-not $candidates.ContainsKey($norm)) { $candidates[$norm] = $true }
    }

    if (-not [string]::IsNullOrWhiteSpace($OrderDocPath)) {
        $fileName = Split-Path -Path $OrderDocPath -Leaf
        if ([string]::IsNullOrWhiteSpace($JobNumber) -and $fileName -match '^(\d{4,6})(\D|$)') {
            $JobNumber = $Matches[1]
        }

        $current = Split-Path -Path $OrderDocPath -Parent
        for ($i = 0; $i -lt 12; $i++) {
            if ([string]::IsNullOrWhiteSpace($current)) { break }
            $leaf = Split-Path -Path $current -Leaf
            if (-not [string]::IsNullOrWhiteSpace($JobNumber) -and
                $leaf -match ("^" + [regex]::Escape($JobNumber) + "(\D|$)")) {
                Add-CandidatePath -Path $current
            }
            $parent = Split-Path -Path $current -Parent
            if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
            $current = $parent
        }

        # If DOCX lives in the job folder already, keep it.
        $docFolder = Split-Path -Path $OrderDocPath -Parent
        if (Test-Path $docFolder) {
            $docLeaf = Split-Path -Path $docFolder -Leaf
            if ($docLeaf -notmatch '(?i)orders') {
                Add-CandidatePath -Path $docFolder
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($JobNumber)) { return @($candidates.Keys) }

    # Search under configured roots (focused + bounded BFS).
    foreach ($root in $CrawlRoots) {
        if ([string]::IsNullOrWhiteSpace($root) -or -not (Test-Path $root)) { continue }

        $starts = New-Object System.Collections.Generic.List[string]
        $starts.Add($root)
        foreach ($suffix in @("Orders", "Orders\Capital", "Projects", "10 - Orders")) {
            $p = Join-Path $root $suffix
            if (Test-Path $p) { $starts.Add($p) }
        }

        foreach ($start in ($starts | Select-Object -Unique)) {
            $queue = New-Object System.Collections.ArrayList
            [void]$queue.Add(@($start, 0))

            while ($queue.Count -gt 0) {
                $item = $queue[0]
                $queue.RemoveAt(0)
                $path = [string]$item[0]
                $depth = [int]$item[1]

                $children = @()
                try {
                    $children = @(Get-ChildItem -Path $path -ErrorAction SilentlyContinue |
                                  Where-Object { $_.PSIsContainer })
                } catch { }

                foreach ($child in $children) {
                    if ($child.Name -match ("^" + [regex]::Escape($JobNumber) + "(\D|$)")) {
                        Add-CandidatePath -Path $child.FullName
                    }
                    if ($depth -lt 5) {
                        [void]$queue.Add(@($child.FullName, $depth + 1))
                    }
                }
            }
        }
    }

    $ranked = @($candidates.Keys | ForEach-Object {
        $p = [string]$_
        $up = $p.ToUpperInvariant()
        $jobUp = $JobNumber.ToUpperInvariant()
        $score = 0
        if ($up.Contains("\" + $jobUp + " ")) { $score += 100 }
        if ($up.Contains("\" + $jobUp + "-")) { $score += 95 }
        if ($up.Contains("\" + $jobUp + "\")) { $score += 85 }
        if ($up.Contains("\EPICOR\ORDERS\CAPITAL\")) { $score += 50 }
        if ($up.Contains("\DRAWINGS\WORKING\")) { $score += 30 }
        if ($up.Contains("\WORKING\")) { $score += 20 }
        if ($up.Contains("\DRAWING")) { $score += 10 }
        if ($up.Contains("\OBSOLETE\") -or $up.Contains("\ARCHIVE\") -or $up.Contains("\OLD\")) { $score -= 40 }
        [pscustomobject]@{ Path = $p; Score = $score }
    } | Sort-Object -Property @{Expression='Score';Descending=$true}, @{Expression='Path';Descending=$false})

    return @($ranked | Select-Object -ExpandProperty Path)
}

function Get-PartNumbersFromSolidWorksAssembly {
    param(
        [object]$SwApp,
        [string]$AssemblyPath,
        [switch]$SkipHelperFallback
    )

    $partSet = @{}
    if ($null -eq $SwApp -or [string]::IsNullOrWhiteSpace($AssemblyPath)) { return @() }

    $swModel = $null
    $swModelOwnerApp = $SwApp
    $alreadyOpen = $false

    try {
        $existing = $SwApp.GetOpenDocumentByName($AssemblyPath)
        if ($existing) {
            $swModel = $existing
            $alreadyOpen = $true
        }
    } catch { }

    if ($null -eq $swModel) {
        $open = Open-SolidWorksAssemblyModel -SwApp $SwApp -AssemblyPath $AssemblyPath
        $swModel = $open.Model
        if ($open.App) { $swModelOwnerApp = $open.App }
        if ($swModel) { $AssemblyPath = [string]$open.Path }
        if ($null -eq $swModel) {
            Write-Log "  SolidWorks failed to open $AssemblyPath" "WARN"
            foreach ($msg in @($open.Errors | Select-Object -First 12)) {
                Write-Log "    $msg" "WARN"
            }
            if (-not $SkipHelperFallback) {
                $helperParts = @(Invoke-SolidWorksStaBomHelper -AssemblyPath $AssemblyPath)
                $assyBase = Get-FileBaseNameUpper -PathOrName $AssemblyPath
                $helperSubParts = @($helperParts | Where-Object { $_ -ne $assyBase })
                if ($helperSubParts.Count -gt 0) {
                    Write-Log "  STA helper returned $($helperParts.Count) unique part number(s)" "SUCCESS"
                    $script:swTraversalDisabled = $false
                    return @($helperParts | Sort-Object -Unique)
                }
                if ($helperParts.Count -gt 0) {
                    Write-Log "  STA helper returned only top-level assembly ($assyBase); treating as failed expansion." "WARN"
                }
            } else {
                Write-Log "  Helper fallback skipped for direct COM retry mode." "WARN"
            }
            return @()
        }
    }

    try { $null = $swModel.ResolveAllLightWeightComponents($true) } catch { }

    # Deep-resolve: force all nested sub-assemblies to fully resolve so
    # GetChildren() returns their contents at every nesting level.
    try {
        $allDeepComps = $swModel.GetComponents($false)
        if ($allDeepComps) {
            Write-Log "  Deep-resolving $(@($allDeepComps).Count) components..." "INFO"
            foreach ($dc in @($allDeepComps)) {
                if ($null -ne $dc) {
                    try {
                        $dcDoc = $dc.GetModelDoc2()
                        if ($dcDoc -and $dcDoc.GetType() -eq 2) {
                            $null = $dcDoc.ResolveAllLightWeightComponents($true)
                        }
                    } catch { }
                }
            }
        }
    } catch { }

    $rootComp = $null
    try {
        $cfg = $swModel.GetActiveConfiguration()
        if ($cfg) { $rootComp = $cfg.GetRootComponent3($true) }
    } catch { }

    if ($rootComp) {
        $seenModelPaths = @{}
        Collect-SolidWorksPartNumbersFromComponent -Component $rootComp -PartSet $partSet -SeenModelPaths $seenModelPaths
    } else {
        Write-Log "  Could not get root component for $AssemblyPath" "WARN"
    }

    # Include top-level assembly itself.
    $asmBase = Get-FileBaseNameUpper -PathOrName $AssemblyPath
    if (-not [string]::IsNullOrWhiteSpace($asmBase)) { $partSet[$asmBase] = $true }

    if (-not $alreadyOpen) {
        try {
            $title = $swModel.GetTitle()
            if (-not [string]::IsNullOrWhiteSpace($title) -and $swModelOwnerApp) { $swModelOwnerApp.CloseDoc($title) }
        } catch { }
    }

    return @($partSet.Keys | Sort-Object)
}

function Get-PartNumbersFromSolidWorksAssemblyFreshInstance {
    param([string]$AssemblyPath)

    $sw = $null
    try {
        $sw = New-Object -ComObject SldWorks.Application
        try { $sw.Visible = $true } catch { }
        try { $sw.UserControl = $true } catch { }
        Start-Sleep -Milliseconds 300
        $parts = Get-PartNumbersFromSolidWorksAssembly -SwApp $sw -AssemblyPath $AssemblyPath
        return @($parts)
    } catch {
        Write-Log "  Fresh SolidWorks instance failed: $($_.Exception.Message)" "WARN"
        return @()
    } finally {
        if ($sw) {
            try { $sw.ExitApp() } catch { }
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sw) } catch { }
        }
    }
}

function Get-SolidWorksAssemblyTopPartNumber {
    param(
        [object]$SwApp,
        [string]$AssemblyPath
    )

    if ($null -eq $SwApp -or [string]::IsNullOrWhiteSpace($AssemblyPath)) { return "" }

    $swModel = $null
    $swModelOwnerApp = $SwApp
    $alreadyOpen = $false
    $topPart = ""

    try {
        try {
            $existing = $SwApp.GetOpenDocumentByName($AssemblyPath)
            if ($existing) { $swModel = $existing; $alreadyOpen = $true }
        } catch { }

        if ($null -eq $swModel) {
            $open = Open-SolidWorksAssemblyModel -SwApp $SwApp -AssemblyPath $AssemblyPath
            $swModel = $open.Model
            if ($open.App) { $swModelOwnerApp = $open.App }
            if ($swModel) { $AssemblyPath = [string]$open.Path }
        }
        if ($null -eq $swModel) { return "" }

        try {
            $cfg = $swModel.GetActiveConfiguration()
            $rootComp = $null
            if ($cfg) { $rootComp = $cfg.GetRootComponent3($true) }
            if ($rootComp) { $topPart = Get-SwComponentPartNumber -Component $rootComp }
        } catch { }

        if ([string]::IsNullOrWhiteSpace($topPart)) {
            $topPart = Get-FileBaseNameUpper -PathOrName $AssemblyPath
        }
        return $topPart
    } finally {
        if (-not $alreadyOpen -and $swModel) {
            try {
                $title = $swModel.GetTitle()
                if (-not [string]::IsNullOrWhiteSpace($title) -and $swModelOwnerApp) { $swModelOwnerApp.CloseDoc($title) }
            } catch { }
        }
    }
}

function Get-AssemblyPathScore {
    param(
        [string]$AssemblyPath,
        [string]$AssemblyPart
    )
    $score = 0
    if ([string]::IsNullOrWhiteSpace($AssemblyPath)) { return $score }
    $pathUp = $AssemblyPath.ToUpperInvariant()
    $base = Get-FileBaseNameUpper -PathOrName $AssemblyPath

    if ($base -eq $AssemblyPart) { $score += 200 }
    elseif ($base.StartsWith($AssemblyPart)) { $score += 150 }
    elseif ($base.Contains($AssemblyPart)) { $score += 100 }

    if ($pathUp.Contains("\EPICOR\ORDERS\CAPITAL\")) { $score += 60 }
    if ($pathUp.Contains("\DRAWINGS\WORKING\")) { $score += 50 }
    if ($pathUp.Contains("\WORKING\")) { $score += 30 }
    if ($pathUp.Contains("\DRAWING")) { $score += 15 }
    if ($pathUp.Contains("\QUOTES\")) { $score -= 90 }
    if ($pathUp.Contains("\OBSOLETE\") -or $pathUp.Contains("\ARCHIVE\") -or $pathUp.Contains("\OLD\")) { $score -= 80 }
    return $score
}

function Test-DisallowedModelPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }

    $up = $Path.ToUpperInvariant()
    if ($up.Contains("\OBSOLETE\") -or
        $up.Contains("\ARCHIVE\") -or
        $up.Contains("\OLD\") -or
        $up.Contains("\DEPRECATED\") -or
        $up.Contains("\QUOTES\")) {
        return $true
    }

    $leaf = ""
    try { $leaf = [System.IO.Path]::GetFileNameWithoutExtension($Path).ToUpperInvariant() } catch { }
    if (-not [string]::IsNullOrWhiteSpace($leaf) -and ($leaf.Contains("-OBS") -or $leaf.EndsWith(" OBS"))) {
        return $true
    }
    return $false
}

function Convert-MappedPathToUnc {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    if ($Path -notmatch '^[A-Za-z]:\\') { return $Path }

    $drive = $Path.Substring(0, 2).ToUpper()
    try {
        $disk = Get-WmiObject -Class Win32_LogicalDisk -Filter ("DeviceID='" + $drive + "'") -ErrorAction SilentlyContinue
        if ($disk -and -not [string]::IsNullOrWhiteSpace($disk.ProviderName)) {
            return ($disk.ProviderName.TrimEnd('\') + $Path.Substring(2))
        }
    } catch { }
    return $Path
}

$script:swInteropInitDone = $false
$script:swOpenDocApisBroken = $false
$script:swTraversalDisabled = $false
$script:swTypedInteropReady = $false
$script:swTypedInteropInitTried = $false
$script:preferVbsMacroOnly = $preferVbsMacroOnly
$script:enableVbsHelperFallback = $enableVbsHelperFallback
$script:requireRunningSolidWorksForBomExpansion = $requireRunningSolidWorksForBomExpansion
$script:allowRecursiveModelSearchFallback = $allowRecursiveModelSearchFallback
$script:allowDirectComBomTraversalFallback = $allowDirectComBomTraversalFallback
$script:processedEntryIds = @{}

function Initialize-SolidWorksTypedInterop {
    if ($script:swTypedInteropInitTried) { return $script:swTypedInteropReady }
    $script:swTypedInteropInitTried = $true
    $script:swTypedInteropReady = $false

    $dlls = @(
        "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist\SolidWorks.Interop.sldworks.dll",
        "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\api\redist\SolidWorks.Interop.swconst.dll",
        "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SolidWorks.Interop.sldworks.dll",
        "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\SolidWorks.Interop.swconst.dll"
    )

    try {
        foreach ($dllPath in ($dlls | Select-Object -Unique)) {
            if (Test-Path $dllPath) {
                try { Add-Type -Path $dllPath -ErrorAction SilentlyContinue } catch { }
            }
        }
        $null = [SolidWorks.Interop.sldworks.ISldWorks]
        $null = [SolidWorks.Interop.swconst.swDocumentTypes_e]
        $script:swTypedInteropReady = $true
        Write-Log "  SolidWorks typed interop loaded (ISldWorks)." "INFO"
    } catch {
        Write-Log "  SolidWorks typed interop unavailable: $($_.Exception.Message)" "WARN"
    }
    return $script:swTypedInteropReady
}

function Get-SolidWorksAppForCalls {
    param([object]$SwApp)

    if ($null -eq $SwApp) { return $null }
    if (-not $script:swInteropInitDone) {
        Write-Log "  Using late-bound SolidWorks COM calls" "INFO"
        $script:swInteropInitDone = $true
    }

    if (Initialize-SolidWorksTypedInterop) {
        try {
            $typed = [SolidWorks.Interop.sldworks.ISldWorks]$SwApp
            if ($typed) { return $typed }
        } catch { }
    }
    return $SwApp
}

function Test-PathLooseEqual {
    param(
        [string]$A,
        [string]$B
    )
    if ([string]::IsNullOrWhiteSpace($A) -or [string]::IsNullOrWhiteSpace($B)) { return $false }

    $aNorm = ""
    $bNorm = ""
    try { $aNorm = [System.IO.Path]::GetFullPath($A).TrimEnd('\').ToLowerInvariant() } catch { $aNorm = $A.Trim().TrimEnd('\').ToLowerInvariant() }
    try { $bNorm = [System.IO.Path]::GetFullPath($B).TrimEnd('\').ToLowerInvariant() } catch { $bNorm = $B.Trim().TrimEnd('\').ToLowerInvariant() }

    if ($aNorm -eq $bNorm) { return $true }

    $aLeaf = ""
    $bLeaf = ""
    try { $aLeaf = [System.IO.Path]::GetFileName($aNorm) } catch { }
    try { $bLeaf = [System.IO.Path]::GetFileName($bNorm) } catch { }
    return (-not [string]::IsNullOrWhiteSpace($aLeaf) -and $aLeaf -eq $bLeaf)
}

function Find-OpenModelInSolidWorksApp {
    param(
        [object]$SwApp,
        [string]$TargetPath
    )

    if ($null -eq $SwApp -or [string]::IsNullOrWhiteSpace($TargetPath)) { return $null }

    try {
        $m = $SwApp.GetOpenDocumentByName($TargetPath)
        if ($m) { return $m }
    } catch { }

    $doc = $null
    try { $doc = $SwApp.ActiveDoc } catch { }
    if ($doc) {
        $p = ""
        try { $p = $doc.GetPathName() } catch { }
        if (Test-PathLooseEqual -A $p -B $TargetPath) { return $doc }
    }

    $current = $null
    try { $current = $SwApp.GetFirstDocument() } catch { }
    $guard = 0
    while ($current -and $guard -lt 300) {
        $guard++
        $p = ""
        try { $p = $current.GetPathName() } catch { }
        if (Test-PathLooseEqual -A $p -B $TargetPath) { return $current }
        try {
            $nextDoc = $null
            try { $nextDoc = $current.GetNext() } catch { }
            if (-not $nextDoc) {
                try { $nextDoc = $current.GetNext } catch { }
            }
            $current = $nextDoc
        } catch {
            break
        }
    }

    return $null
}

function Wait-ForSolidWorksModelFromShellOpen {
    param(
        [string]$TargetPath,
        [int]$TimeoutSeconds = 45
    )

    if ([string]::IsNullOrWhiteSpace($TargetPath)) {
        return @{ App = $null; Model = $null; Error = "TargetPath empty" }
    }

    try {
        Start-Process -FilePath $TargetPath | Out-Null
    } catch {
        return @{ App = $null; Model = $null; Error = "Shell open failed: $($_.Exception.Message)" }
    }

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        $swAny = $null
        try { $swAny = [Runtime.InteropServices.Marshal]::GetActiveObject("SldWorks.Application") } catch { }
        if ($null -eq $swAny) { continue }

        $m = Find-OpenModelInSolidWorksApp -SwApp $swAny -TargetPath $TargetPath
        if ($m) {
            return @{ App = $swAny; Model = $m; Error = "" }
        }
    }

    return @{ App = $null; Model = $null; Error = "Shell open timed out waiting for model in SolidWorks: $TargetPath" }
}

function Invoke-SolidWorksStaBomHelper {
    param(
        [string]$AssemblyPath,
        [int]$TimeoutSeconds = 240
    )

    if ([string]::IsNullOrWhiteSpace($AssemblyPath)) { return @() }
    $helperScript = Join-Path $scriptDir "SolidWorksBomHelper.vbs"
    $psStaHelperScript = Join-Path $scriptDir "SolidWorksBomStaHelper.ps1"

    # Macro-first mode: skip PS helper and use SolidWorks VBS helper directly.
    if ($script:preferVbsMacroOnly) {
        Write-Log "  PS STA helper disabled; using VBS macro helper only." "INFO"
    }
    elseif (Test-Path $psStaHelperScript) {
        $psTmpOut = Join-Path $env:TEMP ("sw_ps_sta_out_" + [guid]::NewGuid().ToString() + ".txt")
        $psTmpErr = Join-Path $env:TEMP ("sw_ps_sta_err_" + [guid]::NewGuid().ToString() + ".txt")
        $psKeepOutPath = ""
        try {
            Write-Log "  Attempting SolidWorks PS STA helper for $AssemblyPath" "WARN"
            # Build a single quoted argument string so special characters in paths
            # (notably '#') are preserved for PowerShell parameter binding.
            $qHelper = $psStaHelperScript.Replace('"', '""')
            $qAssembly = $AssemblyPath.Replace('"', '""')
            $qOut = $psTmpOut.Replace('"', '""')
            $psArgLine = "-NoProfile -STA -ExecutionPolicy Bypass -File ""$qHelper"" -AssemblyPath ""$qAssembly"" -OutputFile ""$qOut"""

            $psProc = Start-Process -FilePath "powershell.exe" `
                -ArgumentList $psArgLine `
                -PassThru `
                -WindowStyle Hidden `
                -RedirectStandardError $psTmpErr

            $psDeadline = (Get-Date).AddSeconds([Math]::Max(30, $TimeoutSeconds))
            $psNextHeartbeat = (Get-Date).AddSeconds(15)
            while ($true) {
                $hasExited = $false
                try { $psProc.Refresh(); $hasExited = $psProc.HasExited } catch { $hasExited = $true }
                if ($hasExited) { break }

                if ((Get-Date) -ge $psNextHeartbeat) {
                    Write-Log "  PS STA helper still running on $AssemblyPath" "INFO"
                    $psNextHeartbeat = (Get-Date).AddSeconds(15)
                }

                if ((Get-Date) -gt $psDeadline) {
                    try { $psProc.Kill() } catch { }
                    Write-Log "  PS STA helper timed out after $TimeoutSeconds sec for $AssemblyPath" "WARN"
                    break
                }
                Start-Sleep -Milliseconds 500
            }

            $psExitCode = $null
            try { $psProc.Refresh(); $psExitCode = [int]$psProc.ExitCode } catch { $psExitCode = $null }
            if ($null -eq $psExitCode) {
                Write-Log "  PS STA helper exit code unavailable for $AssemblyPath" "WARN"
            } elseif ($psExitCode -ne 0) {
                Write-Log "  PS STA helper exit code: $psExitCode for $AssemblyPath" "WARN"
            }

            $psParts = New-Object System.Collections.Generic.List[string]
            if (Test-Path $psTmpOut) {
                foreach ($lineRaw in @(Get-Content -Path $psTmpOut -ErrorAction SilentlyContinue)) {
                    $line = [string]$lineRaw
                    if ($line.StartsWith("PART|")) {
                        $pn = $line.Substring(5).Trim().ToUpperInvariant()
                        if (-not [string]::IsNullOrWhiteSpace($pn) -and -not $psParts.Contains($pn)) {
                            [void]$psParts.Add($pn)
                        }
                    } elseif ($line.StartsWith("LOG|")) {
                        Write-Log ("  PS STA helper: " + $line.Substring(4)) "INFO"
                    } elseif (-not [string]::IsNullOrWhiteSpace($line)) {
                        Write-Log ("  PS STA helper: " + $line) "INFO"
                    }
                }
            } else {
                Write-Log "  PS STA helper output file missing: $psTmpOut" "WARN"
            }
            if (Test-Path $psTmpErr) {
                foreach ($errLine in @(Get-Content -Path $psTmpErr -ErrorAction SilentlyContinue)) {
                    if (-not [string]::IsNullOrWhiteSpace($errLine)) {
                        Write-Log ("  PS STA helper stderr: " + $errLine) "WARN"
                    }
                }
            }

            if ($psParts.Count -gt 0) {
                Write-Log "  PS STA helper returned $($psParts.Count) unique part number(s)" "SUCCESS"
                return $psParts.ToArray()
            } else {
                if (Test-Path $psTmpOut) {
                    $assyBase = Get-FileBaseNameUpper -PathOrName $AssemblyPath
                    if ([string]::IsNullOrWhiteSpace($assyBase)) { $assyBase = "UNKNOWN_ASSY" }
                    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    $psKeepOutPath = Join-Path $env:TEMP ("sw_ps_sta_last_fail_" + $assyBase + "_" + $stamp + ".txt")
                    Copy-Item -Path $psTmpOut -Destination $psKeepOutPath -Force -ErrorAction SilentlyContinue
                    if (Test-Path $psKeepOutPath) {
                        Write-Log "  PS STA helper failure snapshot saved: $psKeepOutPath" "WARN"
                    }
                }
                Write-Log "  PS STA helper returned no parts for $AssemblyPath; falling back to VBS helper." "WARN"
            }
        } catch {
            Write-Log "  PS STA helper invocation failed for $AssemblyPath : $($_.Exception.Message)" "WARN"
        } finally {
            if ((Test-Path $psTmpOut) -and ($psTmpOut -ne $psKeepOutPath)) { Remove-Item $psTmpOut -Force -ErrorAction SilentlyContinue }
            if (Test-Path $psTmpErr) { Remove-Item $psTmpErr -Force -ErrorAction SilentlyContinue }
        }
    } else {
        Write-Log "  SolidWorks PS STA helper script not found: $psStaHelperScript" "WARN"
    }

    if (-not $script:enableVbsHelperFallback) {
        Write-Log "  VBS helper fallback disabled (emailMonitor.enableVbsHelperFallback=false)." "INFO"
        return @()
    }

    if (-not (Test-Path $helperScript)) {
        Write-Log "  SolidWorks VBS helper script not found: $helperScript" "WARN"
        return @()
    }

    $tmpOut = Join-Path $env:TEMP ("sw_sta_out_" + [guid]::NewGuid().ToString() + ".txt")
    $tmpErr = Join-Path $env:TEMP ("sw_sta_err_" + [guid]::NewGuid().ToString() + ".txt")
    $tmpArgs = Join-Path $env:TEMP ("sw_sta_args_" + [guid]::NewGuid().ToString() + ".txt")
    $vbsKeepOutPath = ""
    try {
        Write-Log "  Attempting SolidWorks VBS helper fallback for $AssemblyPath" "WARN"
        @(
            "ASSEMBLY=$AssemblyPath"
            "OUT=$tmpOut"
        ) | Set-Content -Path $tmpArgs -Encoding ASCII

        $argList = @(
            "//B"
            "//nologo"
            $helperScript
            "/argsfile"
            $tmpArgs
        )

        $proc = Start-Process -FilePath "wscript.exe" `
            -ArgumentList $argList `
            -PassThru `
            -WindowStyle Hidden `
            -RedirectStandardError $tmpErr

        $parts = New-Object System.Collections.Generic.List[string]
        $seenOutLineCount = 0
        $helperStart = Get-Date
        $deadline = (Get-Date).AddSeconds([Math]::Max(30, $TimeoutSeconds))
        $nextHeartbeat = (Get-Date).AddSeconds(15)
        while ($true) {
            if (Test-Path $tmpOut) {
                $allLines = @(Get-Content -Path $tmpOut -ErrorAction SilentlyContinue)
                if ($allLines.Count -gt $seenOutLineCount) {
                    for ($idx = $seenOutLineCount; $idx -lt $allLines.Count; $idx++) {
                        $line = [string]$allLines[$idx]
                        if ($line.StartsWith("PART|")) {
                            $pn = $line.Substring(5).Trim().ToUpperInvariant()
                            if (-not [string]::IsNullOrWhiteSpace($pn) -and -not $parts.Contains($pn)) {
                                [void]$parts.Add($pn)
                            }
                        } elseif ($line.StartsWith("LOG|")) {
                            Write-Log ("  VBS helper: " + $line.Substring(4)) "INFO"
                        } elseif (-not [string]::IsNullOrWhiteSpace($line)) {
                            Write-Log ("  VBS helper: " + $line) "INFO"
                        }
                    }
                    $seenOutLineCount = $allLines.Count
                }
            }

            $hasExited = $false
            try { $proc.Refresh(); $hasExited = $proc.HasExited } catch { $hasExited = $true }
            if ($hasExited) { break }

            if ((Get-Date) -ge $nextHeartbeat) {
                $elapsed = [int]([Math]::Round(((Get-Date) - $helperStart).TotalSeconds, 0))
                Write-Log "  VBS helper still running for $elapsed sec on $AssemblyPath" "INFO"
                $nextHeartbeat = (Get-Date).AddSeconds(15)
            }

            if ((Get-Date) -gt $deadline) {
                try { $proc.Kill() } catch { }
                Write-Log "  STA helper timed out after $TimeoutSeconds sec for $AssemblyPath" "WARN"
                return @()
            }
            Start-Sleep -Milliseconds 500
        }

        # Final drain in case helper flushed additional lines right before exit.
        if (Test-Path $tmpOut) {
            $allLines = @(Get-Content -Path $tmpOut -ErrorAction SilentlyContinue)
            if ($allLines.Count -gt $seenOutLineCount) {
                for ($idx = $seenOutLineCount; $idx -lt $allLines.Count; $idx++) {
                    $line = [string]$allLines[$idx]
                    if ($line.StartsWith("PART|")) {
                        $pn = $line.Substring(5).Trim().ToUpperInvariant()
                        if (-not [string]::IsNullOrWhiteSpace($pn) -and -not $parts.Contains($pn)) {
                            [void]$parts.Add($pn)
                        }
                    } elseif ($line.StartsWith("LOG|")) {
                        Write-Log ("  VBS helper: " + $line.Substring(4)) "INFO"
                    } elseif (-not [string]::IsNullOrWhiteSpace($line)) {
                        Write-Log ("  VBS helper: " + $line) "INFO"
                    }
                }
                $seenOutLineCount = $allLines.Count
            }
        }

        if (Test-Path $tmpErr) {
            foreach ($errLine in @(Get-Content -Path $tmpErr -ErrorAction SilentlyContinue)) {
                if (-not [string]::IsNullOrWhiteSpace($errLine)) {
                    Write-Log ("  VBS helper stderr: " + $errLine) "WARN"
                }
            }
        }

        $exitCode = $null
        try { $proc.Refresh(); $exitCode = [int]$proc.ExitCode } catch { $exitCode = $null }
        if (($null -eq $exitCode -or $exitCode -ne 0) -and $parts.Count -eq 0) {
            if ($null -eq $exitCode) {
                Write-Log "  VBS helper exited without a readable exit code for $AssemblyPath" "WARN"
            } else {
                Write-Log "  VBS helper exited with code $exitCode for $AssemblyPath" "WARN"
            }
        }
        if ($parts.Count -eq 0 -and (Test-Path $tmpOut)) {
            $assyBase = Get-FileBaseNameUpper -PathOrName $AssemblyPath
            if ([string]::IsNullOrWhiteSpace($assyBase)) { $assyBase = "UNKNOWN_ASSY" }
            $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $vbsKeepOutPath = Join-Path $env:TEMP ("sw_vbs_last_fail_" + $assyBase + "_" + $stamp + ".txt")
            Copy-Item -Path $tmpOut -Destination $vbsKeepOutPath -Force -ErrorAction SilentlyContinue
            if (Test-Path $vbsKeepOutPath) {
                Write-Log "  VBS helper failure snapshot saved: $vbsKeepOutPath" "WARN"
            }
        }
        return $parts.ToArray()
    } catch {
        Write-Log "  VBS helper invocation failed for $AssemblyPath : $($_.Exception.Message)" "WARN"
        return @()
    } finally {
        if ((Test-Path $tmpOut) -and ($tmpOut -ne $vbsKeepOutPath)) { Remove-Item $tmpOut -Force -ErrorAction SilentlyContinue }
        if (Test-Path $tmpErr) { Remove-Item $tmpErr -Force -ErrorAction SilentlyContinue }
        if (Test-Path $tmpArgs) { Remove-Item $tmpArgs -Force -ErrorAction SilentlyContinue }
    }
}

function Open-SolidWorksAssemblyModel {
    param(
        [object]$SwApp,
        [string]$AssemblyPath
    )

    if ($null -eq $SwApp -or [string]::IsNullOrWhiteSpace($AssemblyPath)) {
        return @{ Model = $null; Path = ""; Errors = @("Invalid input"); App = $null }
    }

    if ($script:swTraversalDisabled) {
        return @{ Model = $null; Path = ""; Errors = @("SolidWorks traversal disabled for this run (COM open methods unavailable)"); App = $null }
    }

    $app = Get-SolidWorksAppForCalls -SwApp $SwApp

    $pathsToTry = New-Object System.Collections.Generic.List[string]
    $full = $AssemblyPath
    try { $full = [System.IO.Path]::GetFullPath($AssemblyPath) } catch { }
    $pathsToTry.Add($full)

    $unc = Convert-MappedPathToUnc -Path $full
    if (-not [string]::Equals($unc, $full, [System.StringComparison]::OrdinalIgnoreCase)) {
        $pathsToTry.Add($unc)
    }

    $errsOut = New-Object System.Collections.Generic.List[string]
    $shellTimeoutSeconds = 18
    $typeLibFailureSeen = $false
    foreach ($path in ($pathsToTry | Select-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        if (Test-DisallowedModelPath -Path $path) {
            $errsOut.Add("Path filtered (disallowed location/name): $path")
            continue
        }
        if (-not (Test-Path $path)) {
            $errsOut.Add("Path not found: $path")
            continue
        }

        try { $app.SetCurrentWorkingDirectory((Split-Path -Path $path -Parent)) } catch { }

        # Primary path: OpenDoc6 variants. If this API is unavailable from this COM context,
        # continue to older APIs before shell-open fallback.
        if (-not $script:swOpenDocApisBroken) {
            $optsToTry = @(3, 1, 0, 2)
            if (Initialize-SolidWorksTypedInterop) {
                try {
                    $optsToTry = @(
                        [int]([SolidWorks.Interop.swconst.swOpenDocOptions_e]::swOpenDocOptions_Silent -bor [SolidWorks.Interop.swconst.swOpenDocOptions_e]::swOpenDocOptions_ReadOnly),
                        [int][SolidWorks.Interop.swconst.swOpenDocOptions_e]::swOpenDocOptions_Silent,
                        0
                    )
                } catch { }
            }
            foreach ($opt in $optsToTry) {
                $errs = 0
                $warns = 0
                $swModel = $null
                try {
                    $docType = 2
                    if (Initialize-SolidWorksTypedInterop) {
                        try { $docType = [int][SolidWorks.Interop.swconst.swDocumentTypes_e]::swDocASSEMBLY } catch { $docType = 2 }
                    }
                    $swModel = $app.OpenDoc6($path, $docType, $opt, "", [ref]$errs, [ref]$warns)
                } catch {
                    $msg = $_.Exception.Message
                    $errsOut.Add("OpenDoc6 exception (opt=$opt) for ${path}: $msg")
                    if ($msg -match 'TYPE_E_ELEMENTNOTFOUND|0x8002802B') {
                        $script:swOpenDocApisBroken = $true
                        $typeLibFailureSeen = $true
                        Write-Log "  OpenDoc6 unavailable from this PowerShell COM context; trying other open methods" "WARN"
                        break
                    }
                }
                if ($swModel) {
                    if ($opt -ne 3) {
                        Write-Log "  Opened assembly with OpenDoc6 option=${opt}: $path" "INFO"
                    }
                    return @{ Model = $swModel; Path = $path; Errors = @(); App = $app }
                }
                $errsOut.Add("OpenDoc6 failed (opt=$opt, err=$errs, warn=$warns) for $path")
            }
        } else {
            $errsOut.Add("OpenDoc6 skipped (marked unavailable in this run) for $path")
        }

        # Fallback 1: DocumentSpecification + OpenDoc7
        try {
            $spec = $app.GetOpenDocSpec($path)
            if ($spec) {
                try { $spec.DocumentType = 2 } catch { }
                try { $spec.Silent = $true } catch { }
                try { $spec.ReadOnly = $true } catch { }
                try { $spec.ConfigurationName = "" } catch { }
                $swModel = $app.OpenDoc7($spec)
                if ($swModel) {
                    Write-Log "  Opened assembly with OpenDoc7: $path" "INFO"
                    return @{ Model = $swModel; Path = $path; Errors = @(); App = $app }
                }
                $errsOut.Add("OpenDoc7 failed for $path")
            } else {
                $errsOut.Add("GetOpenDocSpec returned null for $path")
            }
        } catch {
            $msg = $_.Exception.Message
            $errsOut.Add("OpenDoc7 exception for ${path}: $msg")
            if ($msg -match 'TYPE_E_ELEMENTNOTFOUND|0x8002802B') { $typeLibFailureSeen = $true }
        }

        # Fallback 2: legacy OpenDoc
        try {
            $swModel = $app.OpenDoc($path, 2)
            if ($swModel) {
                Write-Log "  Opened assembly with legacy OpenDoc: $path" "INFO"
                return @{ Model = $swModel; Path = $path; Errors = @(); App = $app }
            }
            $errsOut.Add("OpenDoc failed for $path")
        } catch {
            $msg = $_.Exception.Message
            $errsOut.Add("OpenDoc exception for ${path}: $msg")
            if ($msg -match 'TYPE_E_ELEMENTNOTFOUND|0x8002802B') { $typeLibFailureSeen = $true }
        }

        # Fallback 3: LoadFile4
        try {
            $loadErr = 0
            $swModel = $app.LoadFile4($path, "", $null, [ref]$loadErr)
            if ($swModel) {
                Write-Log "  Opened assembly with LoadFile4: $path" "INFO"
                return @{ Model = $swModel; Path = $path; Errors = @(); App = $app }
            }
            $errsOut.Add("LoadFile4 failed (err=$loadErr) for $path")
        } catch {
            $msg = $_.Exception.Message
            $errsOut.Add("LoadFile4 exception for ${path}: $msg")
            if ($msg -match 'TYPE_E_ELEMENTNOTFOUND|0x8002802B') { $typeLibFailureSeen = $true }
        }

        # Fallback 4: shell-open and attach from whichever SolidWorks instance owns this document.
        $shellAttach = Wait-ForSolidWorksModelFromShellOpen -TargetPath $path -TimeoutSeconds $shellTimeoutSeconds
        if ($shellAttach.Model) {
            Write-Log "  Opened assembly via shell association: $path" "INFO"
            return @{ Model = $shellAttach.Model; Path = $path; Errors = @(); App = $shellAttach.App }
        }
        if (-not [string]::IsNullOrWhiteSpace($shellAttach.Error)) {
            $errsOut.Add($shellAttach.Error)
        }

        if ($typeLibFailureSeen) {
            $script:swTraversalDisabled = $true
            Write-Log "  SolidWorks COM open methods unavailable in this session; disabling further SW tree-walk attempts for this order." "WARN"
            break
        }
    }

    return @{ Model = $null; Path = ""; Errors = $errsOut; App = $null }
}

$script:modelIndexCacheByPath = @{}

function Get-ModelIndexRowsCached {
    param([string]$IndexPath)

    if ([string]::IsNullOrWhiteSpace($IndexPath) -or -not (Test-Path $IndexPath)) { return @() }

    if ($null -eq $script:modelIndexCacheByPath) { $script:modelIndexCacheByPath = @{} }
    $cacheKey = $IndexPath.Trim().ToLowerInvariant()
    if ($script:modelIndexCacheByPath.ContainsKey($cacheKey)) {
        return $script:modelIndexCacheByPath[$cacheKey]
    }

    try {
        Write-Log "  Loading model index ($IndexPath)..." "INFO"
        $rows = Import-Csv -Path $IndexPath
        $script:modelIndexCacheByPath[$cacheKey] = $rows
        Write-Log "  Model index loaded: $($rows.Count) rows" "INFO"
        return $rows
    } catch {
        Write-Log "  Failed to load model index $IndexPath : $($_.Exception.Message)" "WARN"
        return @()
    }
}

function Find-AssemblyFilesForPartNumber {
    param(
        [string]$AssemblyPart,
        [string[]]$CrawlRoots,
        [string]$OrderDocPath = ""
    )

    $assemblyPart = [string]$AssemblyPart
    if ([string]::IsNullOrWhiteSpace($assemblyPart)) { return @() }
    $assemblyPart = $assemblyPart.Trim().ToUpper()

    # When the BOM lists a hand-specific part (e.g. 17141-10-P58-L), the model file may
    # not carry the -L/-R suffix (one shared .sldasm for both hands). Compute the base
    # name now so we can search for it alongside the suffixed name.
    $assemblyPart_noLR = if ($assemblyPart -match '^(.+)-[LR]$') { $Matches[1] } else { $null }

    $candidateDirs = @{}
    $foundByPath = @{}

    # Fast path: resolve from prebuilt model index spreadsheet first.
    $indexHits = 0
    $hasAnyModelIndexFile = $false
    $modelIndexPaths = @()
    if (-not [string]::IsNullOrWhiteSpace($indexFolder)) {
        $modelIndexPaths += (Join-Path $indexFolder "model_index_all.csv")
        $modelIndexPaths += (Join-Path $indexFolder "model_index_clean.csv")
    }
    foreach ($idxPath in ($modelIndexPaths | Select-Object -Unique)) {
        if (Test-Path $idxPath) { $hasAnyModelIndexFile = $true }
        $rows = @(Get-ModelIndexRowsCached -IndexPath $idxPath)
        if ($rows.Count -eq 0) { continue }

        foreach ($row in $rows) {
            $fullPath = ""
            try { $fullPath = [string]$row.FullPath } catch { }
            if ([string]::IsNullOrWhiteSpace($fullPath)) { continue }
            if (Test-DisallowedModelPath -Path $fullPath) { continue }

            $ext = ""
            try { $ext = [string]$row.FileType } catch { }
            if ([string]::IsNullOrWhiteSpace($ext)) {
                try { $ext = [System.IO.Path]::GetExtension($fullPath).TrimStart('.').ToUpperInvariant() } catch { }
            }
            if ($ext -ne "SLDASM") { continue }

            $baseName = ""
            try { $baseName = [string]$row.BaseName } catch { }
            if ([string]::IsNullOrWhiteSpace($baseName)) {
                $baseName = Get-FileBaseNameUpper -PathOrName $fullPath
            } else {
                $baseName = $baseName.Trim().ToUpperInvariant()
            }
            $basePart = ""
            try { $basePart = [string]$row.BasePart } catch { }
            if (-not [string]::IsNullOrWhiteSpace($basePart)) {
                $basePart = $basePart.Trim().ToUpperInvariant()
            }

            $ok = $false
            $assyPattern = "^" + [regex]::Escape($assemblyPart) + "($|[-_ .].*)"
            if ($baseName -eq $assemblyPart -or $baseName -match $assyPattern) { $ok = $true }
            elseif ($basePart -eq $assemblyPart -or $basePart -match $assyPattern) { $ok = $true }
            # For hand-specific parts (P58-L), also accept the unsuffixed model (P58.sldasm)
            if (-not $ok -and $assemblyPart_noLR) {
                $basePattern = "^" + [regex]::Escape($assemblyPart_noLR) + "($|[-_ .].*)"
                if ($baseName -eq $assemblyPart_noLR -or $baseName -match $basePattern) { $ok = $true }
                elseif ($basePart -eq $assemblyPart_noLR -or $basePart -match $basePattern) { $ok = $true }
            }
            if (-not $ok) { continue }
            if (-not (Test-Path $fullPath)) { continue }

            $k = $fullPath.ToLowerInvariant()
            if (-not $foundByPath.ContainsKey($k)) {
                $foundByPath[$k] = [pscustomobject]@{ FullName = $fullPath }
                $indexHits++
            }
        }
    }
    if ($indexHits -gt 0) {
        Write-Log "  Model index match for $assemblyPart : $indexHits candidate(s)" "INFO"
    } elseif (-not $hasAnyModelIndexFile) {
        Write-Log "  Model index not found in $indexFolder (run LaunchModelCrawl.bat for faster model lookup)" "WARN"
    }

    function Add-CandidateDir {
        param([string]$Path)
        if ([string]::IsNullOrWhiteSpace($Path)) { return }
        if (-not (Test-Path $Path)) { return }
        $k = ([System.IO.Path]::GetFullPath($Path)).TrimEnd('\').ToLowerInvariant()
        if (-not $candidateDirs.ContainsKey($k)) { $candidateDirs[$k] = $Path }
    }

    if ($foundByPath.Count -eq 0) {
        if ($hasAnyModelIndexFile -and -not $script:allowRecursiveModelSearchFallback) {
            Write-Log "  No model-index hit for $assemblyPart; skipping recursive disk search fallback (allowRecursiveModelSearchFallback=false)." "WARN"
            return @()
        }

        # Optional targeted filesystem search fallback when index lookup misses.
        # This can recover ghost/non-cached PDM files but may be very slow on full roots.
        if (-not [string]::IsNullOrWhiteSpace($OrderDocPath)) {
            $docDir = Split-Path -Path $OrderDocPath -Parent
            if (Test-Path $docDir) {
                Add-CandidateDir -Path $docDir
                $docParent = Split-Path -Path $docDir -Parent
                if (Test-Path $docParent) { Add-CandidateDir -Path $docParent }
            }
        }

        foreach ($root in $CrawlRoots) {
            if ([string]::IsNullOrWhiteSpace($root) -or -not (Test-Path $root)) { continue }
            Add-CandidateDir -Path $root
        }

        foreach ($dir in $candidateDirs.Values) {
            try {
                # Search for the exact part name (with -L/-R if specified)
                $exact = @(Get-ChildItem -Path $dir -Recurse -Filter "$assemblyPart*.sldasm" -ErrorAction SilentlyContinue |
                           Where-Object { -not $_.PSIsContainer })
                foreach ($f in $exact) {
                    if (Test-DisallowedModelPath -Path $f.FullName) { continue }
                    $k = $f.FullName.ToLowerInvariant()
                    if (-not $foundByPath.ContainsKey($k)) { $foundByPath[$k] = $f }
                }
            } catch { }

            # For hand-specific parts (P58-L), also search the base name (P58.sldasm)
            # because the model may be a single parametric assembly for both hands.
            if ($assemblyPart_noLR) {
                try {
                    $baseExact = @(Get-ChildItem -Path $dir -Recurse -Filter "$assemblyPart_noLR*.sldasm" -ErrorAction SilentlyContinue |
                                   Where-Object { -not $_.PSIsContainer })
                    foreach ($f in $baseExact) {
                        if (Test-DisallowedModelPath -Path $f.FullName) { continue }
                        $k = $f.FullName.ToLowerInvariant()
                        if (-not $foundByPath.ContainsKey($k)) { $foundByPath[$k] = $f }
                    }
                } catch { }
            }
        }

        # Intentionally no broad prefix/contains fallback:
        # we only trust direct assembly-number matches from F80-driven names.
    }

    $ranked = @($foundByPath.Values | ForEach-Object {
        [pscustomobject]@{
            Path = $_.FullName
            Score = (Get-AssemblyPathScore -AssemblyPath $_.FullName -AssemblyPart $assemblyPart)
        }
    } | Sort-Object -Property @{Expression='Score';Descending=$true}, @{Expression='Path';Descending=$false})

    return @($ranked | Select-Object -ExpandProperty Path -First 40)
}

$script:pdfIndexCachePath = ""
$script:pdfIndexCacheRows = $null

function Get-PdfIndexRowsCached {
    param([string]$PdfIndexPath)

    if ([string]::IsNullOrWhiteSpace($PdfIndexPath) -or -not (Test-Path $PdfIndexPath)) {
        return @()
    }

    if ($script:pdfIndexCacheRows -and
        -not [string]::IsNullOrWhiteSpace($script:pdfIndexCachePath) -and
        [string]::Equals($script:pdfIndexCachePath, $PdfIndexPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $script:pdfIndexCacheRows
    }

    try {
        Write-Log "  Loading PDF index ($PdfIndexPath)..." "INFO"
        $rows = Import-Csv -Path $PdfIndexPath
        $script:pdfIndexCachePath = $PdfIndexPath
        $script:pdfIndexCacheRows = $rows
        Write-Log "  PDF index loaded: $($rows.Count) rows" "INFO"
        return $rows
    } catch {
        Write-Log "  Failed to load PDF index for fallback expansion: $($_.Exception.Message)" "WARN"
        return @()
    }
}

function Expand-PartsFromProjectIndexFallback {
    param(
        [System.Collections.Generic.List[string]]$Expanded,
        [string[]]$AssemblyParts,
        [string]$PdfIndexPath
    )

    # IMPORTANT: This fallback uses PDF index file paths to find sub-parts
    # that live in the SAME DIRECTORY as the F80 assembly files.
    # It does NOT do broad project-prefix matching (e.g. grabbing all 25347-*),
    # because the F80 specifies specific models (e.g. A02, A03) and we should
    # only collect parts that are inside those specific models — not the entire
    # project (which would include A01 and everything else).

    if ($null -eq $Expanded -or $AssemblyParts.Count -eq 0) { return 0 }

    Write-Log "=== BOM EXPANSION (PROJECT INDEX FALLBACK) ===" "WARN"
    Write-Log "  F80 assembly models: $($AssemblyParts -join ', ')" "INFO"

    $rows = @(Get-PdfIndexRowsCached -PdfIndexPath $PdfIndexPath)
    if ($rows.Count -eq 0) {
        Write-Log "  Project fallback skipped: PDF index unavailable or empty" "WARN"
        return 0
    }

    # Strategy: For each assembly part from the F80, find its directory in the
    # PDF index, then collect parts whose PDF files are in that same directory
    # (or subdirectories). This approximates "everything inside that model"
    # without needing SolidWorks.
    $before = $Expanded.Count
    foreach ($assyPn in $AssemblyParts) {
        Write-Log "  Looking for sub-parts of $assyPn in PDF index..." "INFO"

        # Find the directory where this assembly's PDF lives
        $assyRow = $rows | Where-Object {
            $bp = ""
            try { $bp = [string]$_.BasePart } catch { }
            $bp.Trim().ToUpper() -eq $assyPn
        } | Select-Object -First 1

        if ($null -eq $assyRow) {
            Write-Log "  Assembly $assyPn not found in PDF index - skipping" "WARN"
            continue
        }

        $assyDir = ""
        try { $assyDir = Split-Path -Path ([string]$assyRow.FullPath) -Parent } catch { }
        if ([string]::IsNullOrWhiteSpace($assyDir)) {
            Write-Log "  Could not determine directory for $assyPn - skipping" "WARN"
            continue
        }

        Write-Log "  Assembly $assyPn directory: $assyDir" "INFO"
        $assyDirUpper = $assyDir.ToUpperInvariant()

        # Collect parts that are in the same directory (these are likely sub-parts)
        $addedAssy = 0
        foreach ($row in $rows) {
            $bp = ""
            try { $bp = [string]$row.BasePart } catch { }
            if ([string]::IsNullOrWhiteSpace($bp)) { continue }
            $bp = $bp.Trim().ToUpper()

            $fp = ""
            try { $fp = [string]$row.FullPath } catch { }
            if ([string]::IsNullOrWhiteSpace($fp)) { continue }

            $fpDir = ""
            try { $fpDir = Split-Path -Path $fp -Parent } catch { }
            if ([string]::IsNullOrWhiteSpace($fpDir)) { continue }

            # Only include parts from the same directory as the assembly
            if ($fpDir.ToUpperInvariant() -ne $assyDirUpper) { continue }
            if ($bp -match '(?i)LOAD[\s_]?CERT|SCOPE|MANUAL') { continue }

            if (-not $Expanded.Contains($bp)) {
                [void]$Expanded.Add($bp)
                $addedAssy++
            }
        }
        Write-Log "  Assembly ${assyPn}: $addedAssy sub-parts found in same directory" "SUCCESS"
    }

    $added = $Expanded.Count - $before
    if ($added -gt 0) {
        Write-Log "  BOM expanded (directory-based fallback): $before -> $($Expanded.Count) parts" "SUCCESS"
    } else {
        Write-Log "  Directory-based fallback added no new parts" "WARN"
    }
    return $added
}

function Expand-PartsFromModelIndexFolderFallback {
    param(
        [System.Collections.Generic.List[string]]$Expanded,
        [string[]]$AssemblyPaths,
        [string[]]$AssemblyParts = @(),
        [string]$OrderDocPath = ""
    )

    if ($null -eq $Expanded -or $AssemblyPaths.Count -eq 0) { return 0 }

    $seedPrefixes = @()
    foreach ($assy in $AssemblyParts) {
        $pn = [string]$assy
        if ($pn -match '^(\d{4,6})-') { $seedPrefixes += $Matches[1] }
    }
    $seedPrefixes = @($seedPrefixes | Select-Object -Unique)

    $folders = New-Object System.Collections.Generic.List[string]
    foreach ($ap in $AssemblyPaths) {
        $path = [string]$ap
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        try {
            $dir = Split-Path -Path $path -Parent
            if (-not [string]::IsNullOrWhiteSpace($dir) -and (Test-Path $dir)) {
                $fullDir = [System.IO.Path]::GetFullPath($dir).TrimEnd('\')
                if (-not $folders.Contains($fullDir)) { [void]$folders.Add($fullDir) }
            }
        } catch { }
    }
    if ($folders.Count -eq 0) { return 0 }

    $rows = @()
    if (-not [string]::IsNullOrWhiteSpace($indexFolder)) {
        $rows += @(Get-ModelIndexRowsCached -IndexPath (Join-Path $indexFolder "model_index_clean.csv"))
        if ($rows.Count -eq 0) {
            $rows += @(Get-ModelIndexRowsCached -IndexPath (Join-Path $indexFolder "model_index_all.csv"))
        }
    }
    if ($rows.Count -eq 0) { return 0 }

    Write-Log ("=== BOM EXPANSION (MODEL INDEX FOLDER FALLBACK) ===") "WARN"
    Write-Log ("  Folder fallback roots: {0}" -f ($folders -join " | ")) "INFO"

    $before = $Expanded.Count
    foreach ($row in $rows) {
        $fullPath = ""
        try { $fullPath = [string]$row.FullPath } catch { }
        if ([string]::IsNullOrWhiteSpace($fullPath)) { continue }
        if (Test-DisallowedModelPath -Path $fullPath) { continue }

        $isUnderRoot = $false
        foreach ($root in $folders) {
            if ([string]::IsNullOrWhiteSpace($root)) { continue }
            if ($fullPath.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
                $isUnderRoot = $true
                break
            }
        }
        if (-not $isUnderRoot) { continue }

        $ext = ""
        try { $ext = [string]$row.FileType } catch { }
        if ([string]::IsNullOrWhiteSpace($ext)) {
            try { $ext = [System.IO.Path]::GetExtension($fullPath).TrimStart('.').ToUpperInvariant() } catch { }
        } else {
            $ext = $ext.Trim().ToUpperInvariant()
        }
        if ($ext -ne "SLDASM" -and $ext -ne "SLDPRT") { continue }

        $part = ""
        try { $part = [string]$row.BasePart } catch { }
        if ([string]::IsNullOrWhiteSpace($part)) {
            $part = Get-FileBaseNameUpper -PathOrName $fullPath
        } else {
            $part = $part.Trim().ToUpperInvariant()
        }
        if ([string]::IsNullOrWhiteSpace($part)) { continue }
        if ($part -match '\^') { continue } # Skip virtual/reference component names
        if ($part -match '(?i)^COPY OF ') { continue }
        if ($part -match '(?i)-OBS(?:$|[^A-Z0-9])') { continue }

        $looksRelevant = $false
        if (Test-DrawingLikePartNumber -PartNumber $part) {
            $looksRelevant = $true
        } elseif (-not (Test-HardwareLikePartNumber -PartNumber $part) -and $part -match '^[A-Z0-9][-A-Z0-9_.]{2,39}$') {
            $looksRelevant = $true
        }
        if (-not $looksRelevant) { continue }

        if ($seedPrefixes.Count -gt 0) {
            $hasPrefix = $false
            foreach ($pref in $seedPrefixes) {
                if ([string]::IsNullOrWhiteSpace($pref)) { continue }
                if ($part -like "$pref-*") { $hasPrefix = $true; break }
            }
            if (-not $hasPrefix -and -not $Expanded.Contains($part)) { continue }
        }

        if (-not $Expanded.Contains($part)) { [void]$Expanded.Add($part) }
    }

    $added = $Expanded.Count - $before
    if ($added -gt 0) {
        Write-Log "  Model-index folder fallback added $added part(s)" "SUCCESS"
    } else {
        Write-Log "  Model-index folder fallback added no parts" "WARN"
    }
    return $added
}

function Expand-AssemblyBOM {
    param(
        [string[]]$PartNumbers,
        [string]$PdfIndexPath,
        [string]$JobNumber = "",
        [string]$OrderDocPath = "",
        [string[]]$CrawlRoots = @()
    )
    $expanded = New-Object System.Collections.Generic.List[string]
    foreach ($pnRaw in $PartNumbers) {
        $pn = [string]$pnRaw
        if ([string]::IsNullOrWhiteSpace($pn)) { continue }
        $up = $pn.Trim().ToUpper()
        if (-not $expanded.Contains($up)) { [void]$expanded.Add($up) }
    }

    # Determine candidates that can map to .SLDASM roots from the F80 part list.
    # Require at least one alpha segment after a dash to avoid OCR-noise items such as
    # "35347-4062" triggering expensive model lookups.
    $assemblyParts = @(
        $expanded |
        Where-Object {
            $pn = [string]$_
            $u = $pn.Trim().ToUpperInvariant()
            (Test-DrawingLikePartNumber -PartNumber $u) -and
            -not (Test-HardwareLikePartNumber -PartNumber $u) -and
            ($u -match '^\d{4,6}(?:-[A-Z0-9]{1,12}){1,4}$') -and
            ($u -match '-[A-Z]')
        } |
        Select-Object -Unique
    )
    if ($assemblyParts.Count -eq 0) {
        Write-Log "  No assembly-style part numbers in F80 list - skipping BOM expansion" "INFO"
        return $expanded.ToArray()
    }
    Write-Log ("  Assembly candidates for BOM expansion: {0}" -f ($assemblyParts -join ", ")) "INFO"

    Write-Log "=== BOM EXPANSION (SOLIDWORKS TREE WALK) ===" "WARN"
    $script:swTraversalDisabled = $false

    Write-Log "  Resolving F80 assembly models from crawl roots (part-number based)..." "INFO"
    $matchedAssemblies = @{}
    $matchedAssemblyCandidates = @{}
    $assyByPath = @{}
    foreach ($assyPn in $assemblyParts) {
        $paths = @(Find-AssemblyFilesForPartNumber -AssemblyPart $assyPn -CrawlRoots $CrawlRoots -OrderDocPath $OrderDocPath)
        if ($paths.Count -gt 0) {
            $topPaths = @($paths | Select-Object -Unique -First 6)
            $bestPath = [string]$topPaths[0]
            $matchedAssemblies[$assyPn] = $bestPath
            $matchedAssemblyCandidates[$assyPn] = $topPaths
            Write-Log "  Matched assembly $assyPn -> $bestPath" "INFO"
            foreach ($p in $topPaths) {
                $k = $p.ToLowerInvariant()
                if (-not $assyByPath.ContainsKey($k)) { $assyByPath[$k] = $p }
            }
        } else {
            Write-Log "  Could not find .SLDASM for assembly $assyPn from crawl roots" "WARN"
        }
    }

    $assyFiles = @($assyByPath.Values | ForEach-Object { [pscustomobject]@{ FullName = $_ } })
    if ($assyFiles.Count -gt 0) {
        Write-Log "  Candidate assembly files discovered: $($assyFiles.Count)" "INFO"
    }

    $swApp = $null
    $createdSwApp = $false
    try {
        $swApp = [Runtime.InteropServices.Marshal]::GetActiveObject("SldWorks.Application")
        try { $swApp.Visible = $true } catch { }
        try { $swApp.UserControl = $true } catch { }
        Write-Log "  Using existing SolidWorks instance (forced visible)" "INFO"
    } catch {
        # Keep going without an attached SW COM session. The STA helper can still
        # launch/attach SolidWorks in its own process and return the full tree.
        if ($script:requireRunningSolidWorksForBomExpansion) {
            Write-Log "  SolidWorks not running - expansion requires active SW session (config). Using fallback." "WARN"
            if ($enableProjectIndexFallback) {
                [void](Expand-PartsFromProjectIndexFallback -Expanded $expanded -AssemblyParts $assemblyParts -PdfIndexPath $PdfIndexPath)
            } else {
                Write-Log "  Project-index fallback disabled (strict mode): not expanding by project/job prefix." "WARN"
            }
            return $expanded.ToArray()
        } else {
            Write-Log "  SolidWorks is not already running; continuing with STA helper-based BOM expansion." "WARN"
            $swApp = $null
        }
    }

    # Fallback mapping: read top-level part numbers from candidate SLDASM files.
    $unmatched = @($assemblyParts | Where-Object { -not $matchedAssemblies.ContainsKey($_) })
    if ($unmatched.Count -gt 0 -and $swApp) {
        Write-Log "  Filename match missed $($unmatched.Count) assembly(s); probing top-level part numbers via SolidWorks..." "WARN"
        $probeLimit = [Math]::Min(250, $assyFiles.Count)
        $probed = 0
        foreach ($file in ($assyFiles | Select-Object -First $probeLimit)) {
            $path = $file.FullName
            if ($matchedAssemblies.Values -contains $path) { continue }
            $probed++
            $topPn = Get-SolidWorksAssemblyTopPartNumber -SwApp $swApp -AssemblyPath $path
            if ([string]::IsNullOrWhiteSpace($topPn)) { continue }
            if ($unmatched -contains $topPn -and -not $matchedAssemblies.ContainsKey($topPn)) {
                $matchedAssemblies[$topPn] = $path
                Write-Log "  Property-matched assembly $topPn -> $path" "INFO"
                $unmatched = @($assemblyParts | Where-Object { -not $matchedAssemblies.ContainsKey($_) })
                if ($unmatched.Count -eq 0) { break }
            }
        }
        if ($unmatched.Count -gt 0) {
            Write-Log "  Still unmatched after probing $probed file(s): $($unmatched -join ', ')" "WARN"
        }
    } elseif ($unmatched.Count -gt 0) {
        Write-Log "  Filename match missed $($unmatched.Count) assembly(s); skipping COM probe because no active SolidWorks app is attached." "WARN"
    }

    if ($matchedAssemblies.Count -eq 0) {
        if ($createdSwApp -and $swApp) {
            try { $swApp.ExitApp() } catch { }
        }
        if ($swApp) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($swApp) } catch { }
        }
        Write-Log "  No F80 assembly numbers mapped to model files - keeping original parts" "WARN"
        if ($enableProjectIndexFallback) {
            [void](Expand-PartsFromProjectIndexFallback -Expanded $expanded -AssemblyParts $assemblyParts -PdfIndexPath $PdfIndexPath)
        } else {
            Write-Log "  Project-index fallback disabled (strict mode): not expanding by project/job prefix." "WARN"
        }
        return $expanded.ToArray()
    }

    $beforeCount = $expanded.Count
    $coveredAssemblyParts = @{}
    try {
        $orderedAssemblyKeys = @($matchedAssemblies.Keys | Sort-Object -Descending)
        foreach ($assyPn in $orderedAssemblyKeys) {
            if ($coveredAssemblyParts.ContainsKey($assyPn)) {
                Write-Log "    Skipping $assyPn (already covered by prior assembly traversal)." "INFO"
                continue
            }
            if ($script:swTraversalDisabled) {
                Write-Log "  Direct COM traversal disabled; attempting STA helper fallback for remaining assemblies." "WARN"
            }
            $assyPath = [string]$matchedAssemblies[$assyPn]
            Write-Log "  Traversing assembly tree: $assyPn" "INFO"
            $swParts = @()
            $candidatePaths = @()
            if ($matchedAssemblyCandidates.ContainsKey($assyPn)) {
                $candidatePaths = @($matchedAssemblyCandidates[$assyPn])
            }
            if ($candidatePaths.Count -eq 0) { $candidatePaths = @($assyPath) }
            $candidatePaths = @($candidatePaths |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not (Test-DisallowedModelPath -Path $_) } |
                Select-Object -Unique)
            if ($candidatePaths.Count -eq 0) {
                Write-Log "    No allowed candidate assembly paths for $assyPn (all filtered as obsolete/archive/quotes)." "WARN"
                continue
            }

            $attempt = 0
            $maxAttempts = [Math]::Min(3, $candidatePaths.Count)
            for ($i = 0; $i -lt $maxAttempts; $i++) {
                $candidatePath = [string]$candidatePaths[$i]
                if ([string]::IsNullOrWhiteSpace($candidatePath)) { continue }
                $attempt++
                if ($attempt -gt 1) {
                    Write-Log "    Retry path $attempt/$maxAttempts for $assyPn -> $candidatePath" "WARN"
                }
                # Preferred path: helper VBS opens the assembly in SolidWorks and runs the BOM macro in-process.
                $helperParts = @(Invoke-SolidWorksStaBomHelper -AssemblyPath $candidatePath)
                $helperSubParts = @($helperParts | Where-Object { $_ -ne $assyPn })
                if ($helperSubParts.Count -gt 0) {
                    Write-Log "    VBS macro helper returned $($helperParts.Count) unique part number(s)" "SUCCESS"
                    $swParts = @($helperParts | Sort-Object -Unique)
                    $assyPath = $candidatePath
                    break
                }
                if ($helperParts.Count -gt 0) {
                    Write-Log "    VBS macro helper returned only top-level assembly for $assyPn; trying direct COM fallback." "WARN"
                } else {
                    Write-Log "    VBS macro helper returned no parts for $assyPn; trying direct COM fallback." "WARN"
                }

                if ($swApp) {
                    if (-not $script:allowDirectComBomTraversalFallback) {
                        Write-Log "    Direct COM traversal fallback disabled (allowDirectComBomTraversalFallback=false) to avoid hangs; skipping this path." "WARN"
                    } else {
                        $directParts = @(Get-PartNumbersFromSolidWorksAssembly -SwApp $swApp -AssemblyPath $candidatePath -SkipHelperFallback)
                        $directSubParts = @($directParts | Where-Object { $_ -ne $assyPn })
                        if ($directSubParts.Count -gt 0) {
                            Write-Log "    Direct COM traversal returned $($directParts.Count) unique part number(s)" "SUCCESS"
                            $swParts = @($directParts | Sort-Object -Unique)
                            $assyPath = $candidatePath
                            break
                        }
                        if ($directParts.Count -gt 0) {
                            Write-Log "    Direct COM traversal returned only top-level assembly for $assyPn." "WARN"
                        } else {
                            Write-Log "    Direct COM traversal returned no parts for $assyPn on this path." "WARN"
                        }
                    }
                } else {
                    Write-Log "    Skipping direct COM fallback on this pass (no active SolidWorks app attached)." "WARN"
                }
            }

            if ($swParts.Count -eq 0 -and -not $createdSwApp) {
                Write-Log "    Retry with dedicated SolidWorks instance for $assyPn..." "WARN"
                $swParts = Get-PartNumbersFromSolidWorksAssemblyFreshInstance -AssemblyPath $assyPath
            }
            Write-Log "    SolidWorks returned $($swParts.Count) unique part number(s)" "INFO"

            if ($swParts.Count -gt 0) {
                foreach ($ap in $swParts) {
                    $apText = [string]$ap
                    if ([string]::IsNullOrWhiteSpace($apText)) { continue }
                    $apText = $apText.Trim().ToUpperInvariant()
                    if ($apText -match '^\d{4,6}-[A-Z][A-Z0-9]{1,6}$') {
                        $coveredAssemblyParts[$apText] = $true
                    }
                }
            }

            foreach ($spRaw in $swParts) {
                $sp = [string]$spRaw
                if ([string]::IsNullOrWhiteSpace($sp)) { continue }
                $sp = $sp.Trim().ToUpper()
                if (-not $expanded.Contains($sp)) { [void]$expanded.Add($sp) }
            }
        }
    } finally {
        if ($createdSwApp -and $swApp) {
            try { $swApp.ExitApp() } catch { }
        }
        if ($swApp) {
            try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($swApp) } catch { }
        }
    }

    $newCount = $expanded.Count - $beforeCount
    Write-Log "  BOM expanded via SolidWorks: $beforeCount -> $($expanded.Count) parts ($newCount added from assembly hierarchy)" "SUCCESS"
    if ($newCount -le 0) {
        # SolidWorks traversal returned no new parts.
        # Do NOT fall back to folder/prefix-based guessing — that grabs everything
        # with the same job number (including A01 and unrelated models).
        # The F80 specifies exact models; if SW can't expand them, keep only
        # what the F80 listed.  The user or index-manager can re-run once SW is
        # available with the correct documents open.
        Write-Log "  SolidWorks expansion returned 0 new parts." "WARN"
        Write-Log "  NOT falling back to folder/prefix guessing to avoid grabbing unrelated models (e.g. A01)." "WARN"
        Write-Log "  To expand: ensure SolidWorks has the target assembly open, then re-process." "WARN"
    }
    return $expanded.ToArray()
}

# ==============================================================================
#  Transmittal Email
# ==============================================================================

function Send-TransmittalEmail {
    param(
        [string]$Subject,
        [string]$JobNumber,
        [string]$ClientName,
        [string[]]$PartNumbers,
        [hashtable[]]$OrderLines = @(),   # full line info: Part, Description (with Rev/Desc/Qty), Rev, Qty
        [string]$OrderFolder,
        [string]$OrderDocPath = "",
        [int]$PdfsFound,
        [int]$DxfsFound,
        [string[]]$NotFound,
        [bool]$TestMode,
        [ValidateSet("Auto","Manual","Hold")]
        [string]$DispatchMode = "Auto",
        [object]$OutlookApp = $null
    )

    $f02Template = ""
    $templateCandidates = @(
        (Join-Path $scriptDir "F02 Document Transmittal v4.0.oft"),
        "U:\30-Common\Forms\NMT Blank Forms\F02 Document Transmittal v4.0.oft"
    )
    foreach ($cand in $templateCandidates) {
        if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path $cand)) {
            $f02Template = $cand
            break
        }
    }

    Write-Log "  Composing transmittal email..." "INFO"
    $outlook = $OutlookApp
    if ($null -eq $outlook) {
        try {
            $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        } catch {
            try {
                $outlook = New-Object -ComObject Outlook.Application
            } catch {
                $msg = $_.Exception.Message
                Write-Log "  Failed to get Outlook COM app for transmittal: $msg" "ERROR"
                if ($msg -match '80070520|0x80070520') {
                    Write-Log "  Outlook logon session missing. Keep Outlook desktop open in the same Windows user session as this monitor." "WARN"
                }
                throw
            }
        }
    }

    # Load the F02 Document Transmittal template
    $mail = $null
    if (-not [string]::IsNullOrWhiteSpace($f02Template) -and (Test-Path $f02Template)) {
        try {
            $mail = $outlook.CreateItemFromTemplate($f02Template)
            Write-Log "  Loaded F02 Document Transmittal template" "INFO"
        } catch {
            Write-Log "  WARNING: template load failed ($($_.Exception.Message)); using blank email." "WARN"
        }
    } else {
        Write-Log "  WARNING: F02 template not found at $f02Template  -  using blank email" "WARN"
    }
    if ($null -eq $mail) {
        $mail = $outlook.CreateItem(0)
    }

    $mailTo = if ($TestMode) { $transmittalToTest } else { $transmittalToProd }
    $mailCc = if ($TestMode) { $transmittalCcTest } else { $transmittalCcProd }
    $mail.To = $mailTo
    if ($mailCc) { $mail.CC = $mailCc }
    Write-Log "  Recipients: To='$mailTo' CC='$mailCc'" "INFO"

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"
    $effectiveJob = if ([string]::IsNullOrWhiteSpace($JobNumber)) { "UNKNOWN" } else { $JobNumber.Trim().ToUpperInvariant() }
    $transmittalNo = "T01"
    $nextTransmittalNum = 1
    try {
        if (-not $TestMode) {
            $historyFileForNo = Join-Path $indexFolder "transmittal_history.json"
            if (Test-Path $historyFileForNo) {
                # Parse history with PS 5.x-safe approach: handle single-object vs array
                $parsedHist = $null
                try { $parsedHist = (Get-Content -Path $historyFileForNo -Raw).Trim() | ConvertFrom-Json } catch {}
                $hist = [System.Collections.Generic.List[object]]::new()
                if ($null -ne $parsedHist) {
                    if ($parsedHist -is [System.Array]) {
                        foreach ($pi in $parsedHist) { $hist.Add($pi) }
                    } else {
                        $hist.Add($parsedHist)
                    }
                }
                $sameJob = @($hist | Where-Object { $_ -and ([string]$_.JobNumber).Trim().ToUpperInvariant() -eq $effectiveJob })
                $maxExistingNo = 0
                foreach ($hrow in $sameJob) {
                    $noText = ""
                    if ($hrow.PSObject.Properties['TransmittalNo']) {
                        $noText = [string]$hrow.TransmittalNo
                    }
                    if ([string]::IsNullOrWhiteSpace($noText)) { continue }
                    $mNo = [regex]::Match($noText.ToUpperInvariant(), '^T?(\d{1,3})$')
                    if ($mNo.Success) {
                        $n = [int]$mNo.Groups[1].Value
                        if ($n -gt $maxExistingNo) { $maxExistingNo = $n }
                    }
                }
                if ($maxExistingNo -gt 0) {
                    $nextTransmittalNum = $maxExistingNo + 1
                } elseif ($sameJob.Count -eq 1) {
                    $nextTransmittalNum = 2
                } else {
                    # Ambiguous history (multiple legacy entries without TransmittalNo): stay on T01.
                    $nextTransmittalNum = 1
                }
                if ($nextTransmittalNum -lt 1) { $nextTransmittalNum = 1 }
            }
        }
    } catch {
        Write-Log "  Could not derive next transmittal number from history: $($_.Exception.Message)" "WARN"
    }
    $transmittalNo = ("T{0:D2}" -f $nextTransmittalNum)
    $isCorrectionTransmittal = $nextTransmittalNum -gt 1
    Write-Log "  Transmittal number selected: $transmittalNo (correction=$isCorrectionTransmittal)" "INFO"
    $docFolder = ""
    if (-not [string]::IsNullOrWhiteSpace($OrderDocPath)) {
        try { $docFolder = Split-Path -Path $OrderDocPath -Parent } catch { $docFolder = "" }
    }
    if ([string]::IsNullOrWhiteSpace($docFolder)) { $docFolder = "C:\NMT_PDM" }
    $burnPath = Join-Path $docFolder "Burn Profiles"

    # --- Resolve proper project drawing path for burn profiles & CAD link ---
    $projectDrawingsPath = ""
    $jobDigits = ""
    if ($effectiveJob -match '\d{4,6}') { $jobDigits = $Matches[0] }
    $spareRoot = "C:\NMT_PDM\Projects\Spare Parts"
    $docLeaf = ""
    $docDerivedDrawingsPath = ""
    $candidateProjectDirs = [System.Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($OrderDocPath)) {
        try {
            $docLeaf = [System.IO.Path]::GetFileNameWithoutExtension($OrderDocPath)
            $docLeaf = ($docLeaf -replace '\s*-\s*F80A?$', '').Trim()
            if (-not [string]::IsNullOrWhiteSpace($docLeaf)) {
                $candidateProjectDirs.Add((Join-Path $spareRoot $docLeaf))
                $docDerivedDrawingsPath = Join-Path (Join-Path $spareRoot $docLeaf) "3 - Design\Drawings"
            }
        } catch { }
    }
    if (-not [string]::IsNullOrWhiteSpace($jobDigits) -and (Test-Path $spareRoot)) {
        Get-ChildItem -Path $spareRoot -Directory -Filter "$jobDigits*" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            ForEach-Object { $candidateProjectDirs.Add($_.FullName) }
    }
    if (-not [string]::IsNullOrWhiteSpace($docLeaf) -and (Test-Path $spareRoot)) {
        $tokens = @(
            $docLeaf.ToUpperInvariant().Split(@(' ', '-', '_', ',', '(', ')'), [System.StringSplitOptions]::RemoveEmptyEntries) |
            Where-Object { $_.Length -ge 3 -and $_ -notmatch '^(F80|DOCX|SPARE|PARTS|ORDER)$' }
        )
        if ($tokens.Count -gt 1) {
            $fuzzy = @(
                Get-ChildItem -Path $spareRoot -Directory -ErrorAction SilentlyContinue |
                ForEach-Object {
                    $nameU = $_.Name.ToUpperInvariant()
                    $score = 0
                    foreach ($tk in $tokens) {
                        if ($nameU.Contains($tk)) { $score++ }
                    }
                    if ($score -ge [Math]::Min(2, $tokens.Count)) {
                        [pscustomobject]@{ Path = $_.FullName; Score = $score; LastWriteTime = $_.LastWriteTime }
                    }
                } |
                Sort-Object Score, LastWriteTime -Descending
            )
            foreach ($m in $fuzzy) { $candidateProjectDirs.Add($m.Path) }
        }
    }
    $seenCandidates = @{}
    foreach ($cand in $candidateProjectDirs) {
        if ([string]::IsNullOrWhiteSpace($cand)) { continue }
        $key = $cand.ToUpperInvariant()
        if ($seenCandidates.ContainsKey($key)) { continue }
        $seenCandidates[$key] = $true
        if (-not (Test-Path $cand)) { continue }
        $drawings = Join-Path $cand "3 - Design\Drawings"
        if (Test-Path $drawings) { $projectDrawingsPath = $drawings; break }
        if ([string]::IsNullOrWhiteSpace($projectDrawingsPath)) { $projectDrawingsPath = $cand }
    }
    if (-not [string]::IsNullOrWhiteSpace($docDerivedDrawingsPath)) {
        # Prefer exact order-doc project path so transmittal always references this specific order.
        $projectDrawingsPath = $docDerivedDrawingsPath
        if (-not (Test-Path $projectDrawingsPath)) {
            Write-Log "  Project path (doc-derived) not present on disk yet: '$projectDrawingsPath'" "WARN"
        }
    }
    if ([string]::IsNullOrWhiteSpace($projectDrawingsPath) -and -not [string]::IsNullOrWhiteSpace($docFolder)) {
        $projectDrawingsPath = $docFolder
    }
    if ([string]::IsNullOrWhiteSpace($projectDrawingsPath)) { $projectDrawingsPath = "N/A" }
    $cadLinkPath = $projectDrawingsPath
    $burnProfilePath = if ($projectDrawingsPath -eq "N/A") { "N/A" } else { Join-Path $projectDrawingsPath "Burn Profiles" }
    $burnProfilesReady = $false
    if ($burnProfilePath -ne "N/A") {
        try {
            if (Test-Path $burnProfilePath) {
                $burnFiles = @(
                    Get-ChildItem -Path $burnProfilePath -File -Recurse -ErrorAction SilentlyContinue |
                    Where-Object { $_.Extension -match '^\.(dxf|dwg|nc|nc1|txt|csv|pdf)$' }
                )
                $burnProfilesReady = $burnFiles.Count -gt 0
            } else {
                $burnProfilesReady = $false
            }
        } catch { $burnProfilesReady = $false }
    }
    $cadLinkReady = $false
    if ($cadLinkPath -ne "N/A") {
        try { $cadLinkReady = Test-Path $cadLinkPath } catch { }
    }
    if (-not $cadLinkReady -and $PdfsFound -gt 0) { $cadLinkReady = $true }
    $burnProfileDisplayPath = if ($burnProfilesReady) { $burnProfilePath } else { "N/A" }
    # Always show the resolved project drawings path for CADLink when available; readiness still controls Yes/No checkbox.
    $cadLinkDisplayPath = if ($cadLinkPath -ne "N/A") { $cadLinkPath } else { "N/A" }
    Write-Log "  Project path: '$projectDrawingsPath' | BurnProfilesReady=$burnProfilesReady | CADLinkReady=$cadLinkReady" "INFO"

    # --- Build notes text from collected drawing files + structured order lines ---
    $normalizePartToken = {
        param([string]$s)
        if ([string]::IsNullOrWhiteSpace($s)) { return "" }
        return (($s.ToUpperInvariant()) -replace '[^A-Z0-9]', '')
    }
    $partCandidates = @(
        $PartNumbers |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object {
            $p = ([string]$_).Trim().ToUpperInvariant()
            $n = & $normalizePartToken $p
            [pscustomobject]@{ Part = $p; Norm = $n; Len = $n.Length }
        } |
        Where-Object { $_.Len -ge 6 } |
        Group-Object Norm |
        ForEach-Object { $_.Group | Select-Object -First 1 } |
        Sort-Object Len -Descending
    )
    $extractPartRevFromFile = {
        param([string]$LeafName)
        $name = [System.IO.Path]::GetFileNameWithoutExtension($LeafName).ToUpperInvariant()
        $base = $name
        $rev  = ""
        if ($name -match '^(?<base>.+?)_REV(?<rev>[A-Z0-9]+)$') {
            $base = $Matches['base']
            $rev  = $Matches['rev']
        }
        if ($base -match '^(?<base>.+?)-REV(?<rev>[A-Z0-9]+)$') {
            $base = $Matches['base']
            if ([string]::IsNullOrWhiteSpace($rev)) { $rev = $Matches['rev'] }
        }
        $baseNorm = & $normalizePartToken $base
        $matchedPart = ""
        foreach ($cand in $partCandidates) {
            if ([string]::IsNullOrWhiteSpace($cand.Norm)) { continue }
            if ($baseNorm.Contains($cand.Norm)) {
                $matchedPart = [string]$cand.Part
                break
            }
        }
        if ([string]::IsNullOrWhiteSpace($matchedPart)) {
            $fallback = ($base -replace '\s+', '' -replace '_', '-')
            if ($fallback -match '^\d{5}-\d{2}-[A-Z]\d{2,3}(?:-[LR])?$|^\d{4,6}-[A-Z0-9]{2,10}(?:-[A-Z0-9]{1,10})*$') {
                $matchedPart = $fallback
            }
        }
        return @{ Base = $matchedPart; Rev = $rev }
    }

    $drawMap = @{}
    $pdfFiles = @(Get-ChildItem -Path $OrderFolder -File -Filter "*.pdf" -ErrorAction SilentlyContinue)
    $dxfFiles = @(Get-ChildItem -Path (Join-Path $OrderFolder "DXFs") -File -Filter "*.dxf" -ErrorAction SilentlyContinue)
    foreach ($f in $pdfFiles) {
        $pr = & $extractPartRevFromFile $f.Name
        if ([string]::IsNullOrWhiteSpace($pr.Base)) { continue }
        if (-not $drawMap.ContainsKey($pr.Base)) { $drawMap[$pr.Base] = @{ Part = $pr.Base; HasPdf = $false; HasDxf = $false; PdfRev = ""; DxfRev = "" } }
        $drawMap[$pr.Base].HasPdf = $true
        if ($pr.Rev) { $drawMap[$pr.Base].PdfRev = $pr.Rev }
    }
    foreach ($f in $dxfFiles) {
        $pr = & $extractPartRevFromFile $f.Name
        if ([string]::IsNullOrWhiteSpace($pr.Base)) { continue }
        if (-not $drawMap.ContainsKey($pr.Base)) { $drawMap[$pr.Base] = @{ Part = $pr.Base; HasPdf = $false; HasDxf = $false; PdfRev = ""; DxfRev = "" } }
        $drawMap[$pr.Base].HasDxf = $true
        if ($pr.Rev) { $drawMap[$pr.Base].DxfRev = $pr.Rev }
    }

    $releasedDrawings = @(
        $drawMap.Values |
        Where-Object { $_.HasPdf -or $_.HasDxf } |
        ForEach-Object {
            $m = [regex]::Match($_.Part, '(\d{2,4})(?:-[A-Z0-9]+)?$')
            [pscustomobject]@{ Part = $_; SortNum = if ($m.Success) { [int]$m.Groups[1].Value } else { 0 } }
        } |
        Sort-Object -Property @{Expression='SortNum';Descending=$true}, @{Expression={ $_.Part };Descending=$false} |
        ForEach-Object { $_.Part }
    )

    $noteParts = @($releasedDrawings | ForEach-Object { $_.Part })
    if ($noteParts.Count -eq 0) {
        $noteParts = @($PartNumbers | Where-Object { Test-DrawingLikePartNumber -PartNumber $_ } | Sort-Object -Unique | Select-Object -First 20)
    }
    if ($noteParts.Count -eq 0) {
        $noteParts = @($PartNumbers | Where-Object { -not (Test-HardwareLikePartNumber -PartNumber $_) } | Sort-Object -Unique | Select-Object -First 12)
    }

    $normalizeQty = {
        param([string]$q)
        if ([string]::IsNullOrWhiteSpace($q)) { return "" }
        $n = ($q -replace ',','.') -replace '\s*EA\s*$',''
        $n = $n -replace '\.0+$',''
        $n = $n -replace '(\d+)\.\d+$','$1'
        return $n.Trim()
    }
    $normalizeRev = {
        param([string]$r)
        if ([string]::IsNullOrWhiteSpace($r)) { return "" }
        $rv = $r.ToUpperInvariant() -replace '[^A-Z0-9]', ''
        if ([string]::IsNullOrWhiteSpace($rv)) { return "" }
        if ($rv.Length -eq 1 -and $rv -eq 'O') { return "0" }
        return $rv
    }
    $normalizeDesc = {
        param([string]$d)
        if ([string]::IsNullOrWhiteSpace($d)) { return "" }
        $x = [string]$d
        # OCR sometimes prefixes descriptions with a stray opening bracket/paren.
        $x = [regex]::Replace($x, '^\s*[\(\[\{]+\s*', '')
        $x = ($x -replace ',', ' ')
        # OCR sometimes uses a period as an intra-phrase separator: "BUSHING. LINK ARM"
        $x = [regex]::Replace($x, '\.(?=\s*[A-Za-z])', ' ')
        $x = ($x -replace '\s*@\s*', ' & ')
        # Common OCR merge in this template family.
        $x = [regex]::Replace($x, '\bLINKARM\b', 'Link Arm', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        # OCR often injects trailing "(D)" on bucket liners where the human transmittal omits it.
        $x = [regex]::Replace($x, '\s*\(D\)\s*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $x = [regex]::Replace($x, '\s+', ' ').Trim()
        return $x
    }
    $normalizeOcrDrawingQty = {
        param([string]$q, [string]$pn)
        if ([string]::IsNullOrWhiteSpace($q)) { return "" }
        if (-not (Test-DrawingLikePartNumber -PartNumber $pn)) { return $q }
        $raw = ($q -replace '[^0-9]', '')
        if ([string]::IsNullOrWhiteSpace($raw)) { return $q }
        $qInt = 0
        if (-not [int]::TryParse($raw, [ref]$qInt)) { return $q }
        # OCR can drop decimal points in F80 image rows: "3.00 EA" -> "300 EA".
        # For drawing-like parts, map implausible 100-based values back to intended order qty.
        if ($qInt -ge 100 -and ($qInt % 100 -eq 0)) {
            $reduced = [int]($qInt / 100)
            if ($reduced -ge 1 -and $reduced -le 50) {
                Write-Log "  OCR Qty normalize: $pn Qty $qInt -> $reduced (decimal dropped in OCR)" "WARN"
                return [string]$reduced
            }
        }
        return $q
    }
    $extractOcrDescHint = {
        param([string]$ocrText)
        if ([string]::IsNullOrWhiteSpace($ocrText)) { return "" }
        $u = $ocrText.ToUpperInvariant()
        $dropTokens = @("MFGJOBTYPE","PART","DESCRIPTION","ORDER","QUANTITY","NO","JOB","MANUFACTURED","ADDITIONAL","CHARGES")
        $keyWords = @("ROPE","GUIDE","BRACKET","HOUSING","LINER","BUCKET","SHAFT","BUSHING","PLATE","WHEEL","BLOCK","SPRING","ROLLER")
        $best = ""
        $seqMatches = [regex]::Matches($u, '([A-Z]{3,}(?:\s+[A-Z]{3,}){1,12})')
        foreach ($sm in $seqMatches) {
            $raw = [string]$sm.Groups[1].Value
            if ([string]::IsNullOrWhiteSpace($raw)) { continue }
            if ($raw -match '(?i)JOB\s+CHARGES|ADDITIONAL\s+CHARGES|NMT\s+NMT') { continue }
            $words = @($raw -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
            if ($words.Count -lt 2) { continue }
            $idx = @()
            for ($wi = 0; $wi -lt $words.Count; $wi++) {
                if ($keyWords -contains $words[$wi]) { $idx += $wi }
            }
            if ($idx.Count -eq 0) { continue }
            $start = [int]$idx[0]
            $end = [int]$idx[$idx.Count - 1]
            $slice = @($words[$start..$end] | Where-Object { $dropTokens -notcontains $_ })
            if ($slice.Count -lt 2) { continue }
            $cand = ($slice -join ' ')
            $cand = [regex]::Replace($cand, '\s+', ' ').Trim()
            if ($cand.Length -lt 8 -or $cand.Length -gt 70) { continue }
            if ($cand.Length -gt $best.Length) { $best = $cand }
        }
        if ([string]::IsNullOrWhiteSpace($best)) { return "" }
        try {
            $best = (Get-Culture).TextInfo.ToTitleCase($best.ToLowerInvariant())
        } catch { }
        return (& $normalizeDesc $best)
    }
    $getPartSort = {
        param([string]$part)
        if ([string]::IsNullOrWhiteSpace($part)) { return 0 }
        $u = $part.ToUpperInvariant()
        $m = [regex]::Match($u, '-P(\d{2,3})(?:-(?:L&R|[LR]))?$')
        if ($m.Success) { return [int]$m.Groups[1].Value }
        $m2 = [regex]::Match($u, '(\d{2,4})(?:-[A-Z0-9]+)?$')
        if ($m2.Success) { return [int]$m2.Groups[1].Value }
        return 0
    }
    $parseContextLine = {
        param([string]$contextText, [string]$pn)
        $parsed = @{ Rev = ""; Desc = ""; Qty = "" }
        if ([string]::IsNullOrWhiteSpace($contextText)) { return $parsed }
        $tmp = [string]$contextText
        $tmp = [regex]::Replace($tmp, "^\s*$([regex]::Escape($pn))\b\s*", "", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $revMatch = [regex]::Match($tmp, '(?i)\bRev\.?\s*([A-Z0-9]+)\b')
        if ($revMatch.Success) {
            $parsed.Rev = $revMatch.Groups[1].Value
            $tmp = [regex]::Replace($tmp, '(?i)\bRev\.?\s*[A-Z0-9]+\b', ' ')
        }
        $qtyMatch = [regex]::Match($tmp, '(?i)\bQty\s*:\s*([0-9]+(?:[,.][0-9]+)?)\b')
        if ($qtyMatch.Success) {
            $parsed.Qty = $qtyMatch.Groups[1].Value
            $tmp = [regex]::Replace($tmp, '(?i)\bQty\s*:\s*[0-9]+(?:[,.][0-9]+)?\b', ' ')
        }
        $tmp = ($tmp -replace '\|', ' ')
        $parsed.Desc = & $normalizeDesc $tmp
        return $parsed
    }
    $singleQtyHint = ""
    $singleDescHint = ""
    if (-not [string]::IsNullOrWhiteSpace($docLeaf)) {
        $mQtyHint = [regex]::Match($docLeaf, '(?i)\bQTY\s*[-_ ]*(\d{1,3})\b')
        if ($mQtyHint.Success) { $singleQtyHint = & $normalizeQty $mQtyHint.Groups[1].Value }

        $docDesc = [string]$docLeaf
        $docDesc = [regex]::Replace($docDesc, '^\s*\d{4,6}\s*-\s*', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if (-not [string]::IsNullOrWhiteSpace($ClientName)) {
            $docDesc = [regex]::Replace($docDesc, '^\s*' + [regex]::Escape($ClientName) + '\s*-\s*', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        }
        $docDesc = [regex]::Replace($docDesc, '(?i)\s*-\s*F80A?\s*$', '')
        $docDesc = [regex]::Replace($docDesc, '(?i)\bQTY\s*[-_ ]*\d+\b', '')
        $docDesc = [regex]::Replace($docDesc, '\s*-\s*$', '')
        $docDesc = & $normalizeDesc $docDesc
        if (-not [string]::IsNullOrWhiteSpace($docDesc)) { $singleDescHint = $docDesc }
    }
    $ocrDescHint = & $extractOcrDescHint $script:ocrInferenceText
    if (-not [string]::IsNullOrWhiteSpace($ocrDescHint)) { $singleDescHint = $ocrDescHint }

    $olMap = @{}
    $orderedNoteParts = [System.Collections.Generic.List[string]]::new()
    $seenOrderedParts = @{}
    foreach ($ol in $OrderLines) {
        if ($null -eq $ol -or -not $ol.Part) { continue }
        $pn = ([string]$ol.Part).Trim().ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($pn)) { continue }
        # Resolve internal part cross-reference (e.g., "Our Part: 4823-P2-03" for customer part 40393469)
        $ip = ""
        if ($ol -is [hashtable] -and $ol.ContainsKey('InternalPart')) { $ip = ([string]$ol.InternalPart).Trim().ToUpperInvariant() }
        elseif ($ol.PSObject -and $ol.PSObject.Properties['InternalPart']) { $ip = ([string]$ol.InternalPart).Trim().ToUpperInvariant() }
        $effectivePn = if (-not [string]::IsNullOrWhiteSpace($ip)) { $ip } else { $pn }
        if ((Test-DrawingLikePartNumber -PartNumber $effectivePn) -and $PartNumbers -contains $effectivePn -and -not $seenOrderedParts.ContainsKey($effectivePn)) {
            $seenOrderedParts[$effectivePn] = $true
            [void]$orderedNoteParts.Add($effectivePn)
        }
        if (-not $olMap.ContainsKey($pn)) {
            $olMap[$pn] = @{
                Description  = [string]$ol.Description
                Rev          = [string]$ol.Rev
                Qty          = [string]$ol.Qty
                CustomerPart = ""
            }
        }
        # Also index under internal part so drawing-matched lookups find order data
        if (-not [string]::IsNullOrWhiteSpace($ip) -and -not $olMap.ContainsKey($ip)) {
            $olMap[$ip] = @{
                Description  = [string]$ol.Description
                Rev          = [string]$ol.Rev
                Qty          = [string]$ol.Qty
                CustomerPart = $pn
            }
        }
    }

    $goodDesc = { param($s,$pn) $s -ne '' -and $s -ne $pn -and $s -notmatch '^(Table Row|OCR Image|Shape|Full Text|Project Overview Text|Nested Table Row)' }
    $noteLineObjects = @()
    foreach ($pn in $noteParts) {
        $entry = if ($olMap.ContainsKey($pn)) { $olMap[$pn] } else { $null }
        $draw = if ($drawMap.ContainsKey($pn)) { $drawMap[$pn] } else { $null }
        $ctxParsed = @{ Rev = ""; Desc = ""; Qty = "" }
        if ($script:ocrPartContext -and $script:ocrPartContext.ContainsKey($pn)) {
            $ctxParsed = & $parseContextLine ([string]$script:ocrPartContext[$pn]) $pn
        }
        $fileRev = ""
        if ($draw) {
            $pdfRevNorm = & $normalizeRev ([string]$draw.PdfRev)
            $dxfRevNorm = & $normalizeRev ([string]$draw.DxfRev)
            if (-not [string]::IsNullOrWhiteSpace($pdfRevNorm) -and $pdfRevNorm -ne "0") {
                $fileRev = $pdfRevNorm
            } elseif (-not [string]::IsNullOrWhiteSpace($dxfRevNorm) -and $dxfRevNorm -ne "0") {
                $fileRev = $dxfRevNorm
            } elseif (-not [string]::IsNullOrWhiteSpace($pdfRevNorm)) {
                $fileRev = $pdfRevNorm
            } elseif (-not [string]::IsNullOrWhiteSpace($dxfRevNorm)) {
                $fileRev = $dxfRevNorm
            }
        }

        $desc = ""
        if ($entry -and (& $goodDesc $entry.Description $pn)) {
            $desc = [string]$entry.Description
            $desc = [regex]::Replace($desc, "^\s*$([regex]::Escape($pn))\b\s*", "")
            $desc = [regex]::Replace($desc, '(?i)(^|\s+)Rev\.?\s*[A-Z0-9]+\b', ' ')
            $desc = [regex]::Replace($desc, '(?i)(^|\s+)Qty:\s*\S+\b', ' ')
            $desc = & $normalizeDesc $desc
        }
        if ([string]::IsNullOrWhiteSpace($desc) -and (& $goodDesc $ctxParsed.Desc $pn)) {
            $desc = & $normalizeDesc ([string]$ctxParsed.Desc)
        }

        $rev = ""
        if ($entry) { $rev = & $normalizeRev $entry.Rev }
        if ([string]::IsNullOrWhiteSpace($rev) -and -not [string]::IsNullOrWhiteSpace($ctxParsed.Rev)) { $rev = & $normalizeRev $ctxParsed.Rev }
        if ([string]::IsNullOrWhiteSpace($rev) -and -not [string]::IsNullOrWhiteSpace($fileRev)) { $rev = & $normalizeRev $fileRev }
        $qty = ""
        $qtySource = ""
        $qtyIsDefaultFromDraw = $false
        if ($entry) {
            $qty = & $normalizeQty $entry.Qty
            if (-not [string]::IsNullOrWhiteSpace($qty)) { $qtySource = "entry" }
        }
        if ([string]::IsNullOrWhiteSpace($qty) -and -not [string]::IsNullOrWhiteSpace($ctxParsed.Qty)) {
            $qty = & $normalizeQty $ctxParsed.Qty
            if (-not [string]::IsNullOrWhiteSpace($qty)) { $qtySource = "context" }
        }
        if ([string]::IsNullOrWhiteSpace($desc) -and $noteParts.Count -eq 1 -and -not [string]::IsNullOrWhiteSpace($singleDescHint)) {
            $desc = & $normalizeDesc $singleDescHint
        }
        if ([string]::IsNullOrWhiteSpace($qty) -and $noteParts.Count -eq 1 -and -not [string]::IsNullOrWhiteSpace($singleQtyHint)) {
            $qty = & $normalizeQty $singleQtyHint
            if (-not [string]::IsNullOrWhiteSpace($qty)) { $qtySource = "hint" }
        }
        if (-not [string]::IsNullOrWhiteSpace($qty) -and $qtySource -eq "context") {
            $qty = & $normalizeOcrDrawingQty $qty $pn
        }
        $displayPrefix = $pn
        if ($entry -and -not [string]::IsNullOrWhiteSpace([string]$entry.CustomerPart)) {
            $displayPrefix = ("{0} our part # {1}" -f ([string]$entry.CustomerPart).Trim(), $pn)
        }
        $line = $displayPrefix
        if (-not [string]::IsNullOrWhiteSpace($rev) -and $rev -ne "0") { $line += " Rev.$rev" }
        if (-not [string]::IsNullOrWhiteSpace($desc)) { $line += " $desc" }
        if (-not [string]::IsNullOrWhiteSpace($qty)) { $line += " Qty: $qty" }
        $hasDetail = $false
        if (-not [string]::IsNullOrWhiteSpace($rev) -and $rev -ne "0") { $hasDetail = $true }
        if (-not [string]::IsNullOrWhiteSpace($desc)) { $hasDetail = $true }
        if (-not [string]::IsNullOrWhiteSpace($qty) -and -not $qtyIsDefaultFromDraw) { $hasDetail = $true }
        $orderIx = $orderedNoteParts.IndexOf($pn)
        $noteLineObjects += [pscustomobject]@{ Part = $pn; Line = $line.Trim(); HasDetail = $hasDetail; OrderIndex = $orderIx }
    }
    $detailedOnly = @($noteLineObjects | Where-Object { $_.HasDetail })
    if ($detailedOnly.Count -gt 0) {
        $noteLineObjects = $detailedOnly
    }
    # Consolidate mirrored left/right parts into one "L&R" line when data matches.
    $byPart = @{}
    foreach ($obj in $noteLineObjects) {
        $lineText = [string]$obj.Line
        if ([string]::IsNullOrWhiteSpace($lineText)) { continue }
        $mLine = [regex]::Match($lineText, '^(?<part>\S+)(?:\s+Rev\.(?<rev>[A-Z0-9]+))?(?:\s+(?<desc>.*?))?(?:\s+Qty:\s*(?<qty>[0-9]+(?:[.,][0-9]+)?))?$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if (-not $mLine.Success) {
            $fallbackPartKey = if ($obj.PSObject.Properties.Name -contains 'Part' -and -not [string]::IsNullOrWhiteSpace([string]$obj.Part)) {
                [string]$obj.Part
            } else {
                $lineText
            }
            $byPart[$fallbackPartKey] = [pscustomobject]@{ Part = $fallbackPartKey; Rev = ""; Desc = ""; Qty = ""; Raw = $lineText; OrderIndex = [int]$obj.OrderIndex }
            continue
        }
        $partKey = if ($obj.PSObject.Properties.Name -contains 'Part' -and -not [string]::IsNullOrWhiteSpace([string]$obj.Part)) {
            [string]$obj.Part
        } else {
            [string]$mLine.Groups['part'].Value
        }
        $byPart[$partKey] = [pscustomobject]@{
            Part       = $partKey
            Rev        = [string]$mLine.Groups['rev'].Value
            Desc       = [string]$mLine.Groups['desc'].Value
            Qty        = [string]$mLine.Groups['qty'].Value
            Raw        = $lineText
            OrderIndex = [int]$obj.OrderIndex
        }
    }
    $consumed = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $mergedObjects = [System.Collections.Generic.List[object]]::new()
    foreach ($p in @($byPart.Keys)) {
        if ($consumed.Contains($p)) { continue }
        $mLR = [regex]::Match($p, '^(?<base>.+)-(?<side>[LR])$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if (-not $mLR.Success) {
            [void]$mergedObjects.Add($byPart[$p]); [void]$consumed.Add($p); continue
        }
        $base = [string]$mLR.Groups['base'].Value
        $other = if ($mLR.Groups['side'].Value.ToUpperInvariant() -eq 'L') { "$base-R" } else { "$base-L" }
        if (-not $byPart.ContainsKey($other) -or $consumed.Contains($other)) {
            [void]$mergedObjects.Add($byPart[$p]); [void]$consumed.Add($p); continue
        }
        $a = $byPart[$p]; $b = $byPart[$other]
        $revA = (& $normalizeRev $a.Rev); $revB = (& $normalizeRev $b.Rev)
        $descA = (& $normalizeDesc $a.Desc); $descB = (& $normalizeDesc $b.Desc)
        $qtyA = (& $normalizeQty $a.Qty); $qtyB = (& $normalizeQty $b.Qty)
        if ([string]::IsNullOrWhiteSpace($descA)) { $descA = $descB }
        if ([string]::IsNullOrWhiteSpace($descB)) { $descB = $descA }
        if ([string]::IsNullOrWhiteSpace($qtyA)) { $qtyA = $qtyB }
        if ([string]::IsNullOrWhiteSpace($qtyB)) { $qtyB = $qtyA }
        if ($revA -eq $revB -and $descA -eq $descB -and $qtyA -eq $qtyB) {
            $line = "$base L&R"
            if (-not [string]::IsNullOrWhiteSpace($revA) -and $revA -ne "0") { $line += " Rev.$revA" }
            if (-not [string]::IsNullOrWhiteSpace($descA)) { $line += " $descA" }
            if (-not [string]::IsNullOrWhiteSpace($qtyA)) { $line += " Qty: $qtyA Each L&R" }
            $baseOrder = if ($a.OrderIndex -ge 0) { [int]$a.OrderIndex } elseif ($b.OrderIndex -ge 0) { [int]$b.OrderIndex } else { -1 }
            [void]$mergedObjects.Add([pscustomobject]@{ Part = "$base-L&R"; Rev = $revA; Desc = $descA; Qty = $qtyA; Raw = $line.Trim(); OrderIndex = $baseOrder })
            [void]$consumed.Add($p); [void]$consumed.Add($other)
        } else {
            [void]$mergedObjects.Add($a); [void]$mergedObjects.Add($b)
            [void]$consumed.Add($p); [void]$consumed.Add($other)
        }
    }
    $useOriginalOrder = ($orderedNoteParts.Count -gt 0)
    $sortedMerged = if ($useOriginalOrder) {
        @($mergedObjects | Sort-Object -Property @{Expression={ if ($_.PSObject.Properties.Name -contains 'OrderIndex' -and $_.OrderIndex -ge 0) { [int]$_.OrderIndex } else { 9999 } };Descending=$false}, @{Expression='Part';Descending=$false})
    } else {
        @($mergedObjects | Sort-Object -Property @{Expression={ & $getPartSort $_.Part };Descending=$true}, @{Expression='Part';Descending=$false})
    }
    $noteLineObjects = @($sortedMerged | ForEach-Object {
        $txt = [string]$_.Raw
        if ([string]::IsNullOrWhiteSpace($txt)) {
            $txt = [string]$_.Part
            if (-not [string]::IsNullOrWhiteSpace($_.Rev) -and $_.Rev -ne "0") { $txt += " Rev.$($_.Rev)" }
            if (-not [string]::IsNullOrWhiteSpace($_.Desc)) { $txt += " $($_.Desc)" }
            if (-not [string]::IsNullOrWhiteSpace($_.Qty)) { $txt += " Qty: $($_.Qty)" }
        }
        [pscustomobject]@{ Line = $txt.Trim(); HasDetail = $true; OrderIndex = if ($_.PSObject.Properties.Name -contains 'OrderIndex') { $_.OrderIndex } else { -1 } }
    })
    $noteLines = @($noteLineObjects | ForEach-Object { [string]$_.Line } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($noteLines.Count -eq 0 -and $noteParts.Count -gt 0) {
        $noteLines = @($noteParts | Select-Object -First 1)
    }
    Write-Log ("  Transmittal drawing lines ({0}): {1}" -f $noteLines.Count, ($noteLines -join " || ")) "INFO"

    $pluralWord = if ($noteLines.Count -eq 1) { "drawing has" } else { "drawings have" }
    $notesLines = @("The following $pluralWord been released for Spare Parts job #$effectiveJob") + $noteLines

    # --- Populate HTMLBody ---
    $htmlTemplate = ""
    try { $htmlTemplate = [string]$mail.HTMLBody } catch { $htmlTemplate = "" }
    if (-not [string]::IsNullOrWhiteSpace($htmlTemplate)) {
        $h = $htmlTemplate

        # Core placeholders
        $h = $h.Replace("265??", $effectiveJob)
        $h = $h.Replace("T0?", $transmittalNo)

        # Path placeholders (appears twice: first = burn profiles, second = CAD link)
        $pathPlaceholder = "J:\Epicor\Orders\Capital\??"
        $burnPathHtml = [System.Security.SecurityElement]::Escape($burnProfileDisplayPath)
        $cadPathHtml = [System.Security.SecurityElement]::Escape($cadLinkDisplayPath)
        Write-Log "  HTMLBody length: $($h.Length) | burnPath: '$burnProfileDisplayPath' | cadPath: '$cadLinkDisplayPath'" "INFO"
        $idx1 = $h.IndexOf($pathPlaceholder)
        Write-Log "  Path placeholder 1 at index: $idx1" "INFO"
        if ($idx1 -ge 0) {
            $h = $h.Substring(0, $idx1) + $burnPathHtml + $h.Substring($idx1 + $pathPlaceholder.Length)
        } else {
            Write-Log "  Path placeholder 1 NOT found - burn profile path will be missing (check _debug_transmittal_htmlbody.html)" "WARN"
        }
        $idx2 = $h.IndexOf($pathPlaceholder)
        Write-Log "  Path placeholder 2 at index: $idx2" "INFO"
        if ($idx2 -ge 0) {
            $h = $h.Substring(0, $idx2) + $cadPathHtml + $h.Substring($idx2 + $pathPlaceholder.Length)
        } else {
            Write-Log "  Path placeholder 2 NOT found - CADLink path will be missing" "WARN"
        }

        # Resolve a section span using a start title and optional next section title.
        $getSectionSlice = {
            param([string]$Html, [string]$SectionTitle, [string]$NextSectionTitle = "")
            $sIdx = $Html.IndexOf($SectionTitle, [System.StringComparison]::OrdinalIgnoreCase)
            if ($sIdx -lt 0) { return $null }

            $endIdx = -1
            if (-not [string]::IsNullOrWhiteSpace($NextSectionTitle)) {
                $nextIdx = $Html.IndexOf($NextSectionTitle, $sIdx, [System.StringComparison]::OrdinalIgnoreCase)
                if ($nextIdx -gt $sIdx) {
                    $endIdx = $nextIdx
                }
            }
            if ($endIdx -lt 0) {
                $tblEnd = $Html.IndexOf('</table>', $sIdx, [System.StringComparison]::OrdinalIgnoreCase)
                if ($tblEnd -lt 0) { return $null }
                $endIdx = $tblEnd + 8
            }

            [pscustomobject]@{
                Start  = $sIdx
                End    = $endIdx
                Before = $Html.Substring(0, $sIdx)
                Sect   = $Html.Substring($sIdx, $endIdx - $sIdx)
                After  = $Html.Substring($endIdx)
            }
        }

        # Checkboxes - map labels to nearest checkbox token in the same section.
        $setSectionCheckboxes = {
            param([ref]$HtmlRef, [string]$SectionTitle, [hashtable]$LabelChecks, [string]$NextSectionTitle = "")
            if ($null -eq $LabelChecks -or $LabelChecks.Count -eq 0) { return $false }
            $html = [string]$HtmlRef.Value

            $slice = & $getSectionSlice $html $SectionTitle $NextSectionTitle
            if ($null -eq $slice) { return $false }
            $before = [string]$slice.Before
            $sect   = [string]$slice.Sect
            $after  = [string]$slice.After

            $candidates = [System.Collections.Generic.List[object]]::new()
            $attrMatches = [regex]::Matches($sect, 'CheckBoxIsChecked="([tf])"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($attrMatches.Count -gt 0) {
                foreach ($am in $attrMatches) {
                    $candidates.Add([pscustomobject]@{
                        Index  = $am.Index
                        Length = $am.Length
                        Kind   = "attr"
                    })
                }
            } else {
                $glyphPattern = (([string][char]0x2610) + "|" + ([string][char]0x25A1) + "|" + ([string][char]0x2611) + "|" + ([string][char]0x2612) + "|" + ([string][char]0x2713) + "|" + ([string][char]0x2714))
                $glyphMatches = [regex]::Matches($sect, $glyphPattern)
                if ($glyphMatches.Count -gt 0) {
                    foreach ($gm in $glyphMatches) {
                        $candidates.Add([pscustomobject]@{
                            Index  = $gm.Index
                            Length = $gm.Length
                            Kind   = "glyph"
                        })
                    }
                } else {
                    $entityPattern = '&#(?:x2610|x25A1|x2611|x2612|x2713|x2714|9744|9633|9745|9746|10003|10004);'
                    $entityMatches = [regex]::Matches($sect, $entityPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    foreach ($em in $entityMatches) {
                        $candidates.Add([pscustomobject]@{
                            Index  = $em.Index
                            Length = $em.Length
                            Kind   = "entity"
                        })
                    }
                }
            }
            if ($candidates.Count -eq 0) { return $false }

            $replacements = [System.Collections.Generic.List[object]]::new()
            $used = @{}

            foreach ($label in $LabelChecks.Keys) {
                $lblIdx = $sect.IndexOf([string]$label, [System.StringComparison]::OrdinalIgnoreCase)
                if ($lblIdx -lt 0) { continue }

                $bestIdx = -1
                $bestDist = [int]::MaxValue
                for ($i = 0; $i -lt $candidates.Count; $i++) {
                    if ($used.ContainsKey($i)) { continue }
                    $ci = [int]$candidates[$i].Index
                    $dist = [Math]::Abs($ci - $lblIdx)
                    if ($dist -lt $bestDist) {
                        $bestDist = $dist
                        $bestIdx = $i
                    }
                }
                if ($bestIdx -lt 0) { continue }
                $used[$bestIdx] = $true
                $checked = [bool]$LabelChecks[$label]
                $c = $candidates[$bestIdx]
                $replacementText = ""
                if ($c.Kind -eq "attr") {
                    $desiredAttr = if ($checked) { "t" } else { "f" }
                    $replacementText = ('CheckBoxIsChecked="' + $desiredAttr + '"')
                } elseif ($c.Kind -eq "glyph") {
                    $replacementText = if ($checked) { [string][char]0x2611 } else { [string][char]0x2610 }
                } else {
                    $replacementText = if ($checked) { '&#9745;' } else { '&#9744;' }
                }
                $replacements.Add([pscustomobject]@{
                    Index  = [int]$c.Index
                    Length = [int]$c.Length
                    Text   = $replacementText
                })
            }
            if ($replacements.Count -eq 0) { return $false }

            foreach ($r in ($replacements | Sort-Object Index -Descending)) {
                $sect = $sect.Substring(0, $r.Index) + $r.Text + $sect.Substring($r.Index + $r.Length)
            }
            $HtmlRef.Value = $before + $sect + $after
            return $true
        }
        $setSectionRowCheckbox = {
            param([ref]$HtmlRef, [string]$SectionTitle, [string]$Label, [bool]$Checked, [string]$NextSectionTitle = "")
            $html = [string]$HtmlRef.Value
            $slice = & $getSectionSlice $html $SectionTitle $NextSectionTitle
            if ($null -eq $slice) { return $false }
            $before = [string]$slice.Before
            $sect   = [string]$slice.Sect
            $after  = [string]$slice.After
            $desiredAttr = if ($Checked) { "t" } else { "f" }

            $rowMatches = [regex]::Matches($sect, '(?is)<tr\b.*?</tr>')
            $didRowUpdate = $false
            $rowFound = $false
            $rowHadKnownToken = $false
            foreach ($rm in $rowMatches) {
                $row = [string]$rm.Value
                if ($row -notmatch [regex]::Escape($Label)) { continue }
                $rowFound = $true

                $row2 = [string]$row
                if ([regex]::IsMatch($row, 'CheckBoxIsChecked="([tf])"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
                    $rowHadKnownToken = $true
                    $lblIdx = $row2.IndexOf([string]$Label, [System.StringComparison]::OrdinalIgnoreCase)
                    if ($lblIdx -lt 0) { $lblIdx = 0 }
                    $attrMatchesRow = [regex]::Matches($row2, 'CheckBoxIsChecked="([tf])"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    if ($attrMatchesRow.Count -gt 0) {
                        $bestA = $attrMatchesRow[0]
                        $bestDistA = [Math]::Abs($bestA.Index - $lblIdx)
                        foreach ($amr in $attrMatchesRow) {
                            $dA = [Math]::Abs($amr.Index - $lblIdx)
                            if ($dA -lt $bestDistA) { $bestDistA = $dA; $bestA = $amr }
                        }
                        $row2 = $row2.Substring(0, $bestA.Index) + ('CheckBoxIsChecked="' + $desiredAttr + '"') + $row2.Substring($bestA.Index + $bestA.Length)
                    }
                }
                # Also force the visible glyph in this row; some templates render glyphs from static text.
                $glyphPattern = (([string][char]0x2610) + "|" + ([string][char]0x25A1) + "|" + ([string][char]0x2611) + "|" + ([string][char]0x2612) + "|" + ([string][char]0x2713) + "|" + ([string][char]0x2714))
                if ([regex]::IsMatch($row2, $glyphPattern)) {
                    $rowHadKnownToken = $true
                    $replacementGlyph = if ($Checked) { [string][char]0x2611 } else { [string][char]0x2610 }
                    $lblIdx = $row2.IndexOf([string]$Label, [System.StringComparison]::OrdinalIgnoreCase)
                    if ($lblIdx -lt 0) { $lblIdx = 0 }
                    $glyphMatchesRow = [regex]::Matches($row2, $glyphPattern)
                    if ($glyphMatchesRow.Count -gt 0) {
                        $bestG = $glyphMatchesRow[0]
                        $bestDistG = [Math]::Abs($bestG.Index - $lblIdx)
                        foreach ($gmr in $glyphMatchesRow) {
                            $dG = [Math]::Abs($gmr.Index - $lblIdx)
                            if ($dG -lt $bestDistG) { $bestDistG = $dG; $bestG = $gmr }
                        }
                        $row2 = $row2.Substring(0, $bestG.Index) + $replacementGlyph + $row2.Substring($bestG.Index + $bestG.Length)
                    }
                } else {
                    # Some Outlook/Word templates store checkbox symbols as HTML entities.
                    $entityPattern = '&#(?:x2610|x25A1|x2611|x2612|x2713|x2714|9744|9633|9745|9746|10003|10004);'
                    if ([regex]::IsMatch($row2, $entityPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
                        $rowHadKnownToken = $true
                        $replacementEntity = if ($Checked) { '&#9745;' } else { '&#9744;' }
                        $lblIdx = $row2.IndexOf([string]$Label, [System.StringComparison]::OrdinalIgnoreCase)
                        if ($lblIdx -lt 0) { $lblIdx = 0 }
                        $entityMatchesRow = [regex]::Matches($row2, $entityPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                        if ($entityMatchesRow.Count -gt 0) {
                            $bestE = $entityMatchesRow[0]
                            $bestDistE = [Math]::Abs($bestE.Index - $lblIdx)
                            foreach ($emr in $entityMatchesRow) {
                                $dE = [Math]::Abs($emr.Index - $lblIdx)
                                if ($dE -lt $bestDistE) { $bestDistE = $dE; $bestE = $emr }
                            }
                            $row2 = $row2.Substring(0, $bestE.Index) + $replacementEntity + $row2.Substring($bestE.Index + $bestE.Length)
                        }
                    }
                }
                if ($row2 -ne $row) {
                    $sect = $sect.Substring(0, $rm.Index) + $row2 + $sect.Substring($rm.Index + $rm.Length)
                    $didRowUpdate = $true
                }
                break
            }

            if ($rowFound) {
                if ($didRowUpdate) {
                    $HtmlRef.Value = $before + $sect + $after
                    return $true
                }
                # Row existed but no known checkbox token was found/changed.
                # Return false so caller can try section-level fallback.
                return [bool]$rowHadKnownToken
            }

            # Fallback to distance-based label mapping only when row-label lookup fails.
            & $setSectionCheckboxes $HtmlRef $SectionTitle ([ordered]@{ $Label = $Checked }) $NextSectionTitle
            return $false
        }
        $setSectionCheckboxesByOrder = {
            param([ref]$HtmlRef, [string]$SectionTitle, [bool[]]$Checks, [string]$NextSectionTitle = "")
            if ($null -eq $Checks -or $Checks.Count -eq 0) { return }
            $html = [string]$HtmlRef.Value
            $slice = & $getSectionSlice $html $SectionTitle $NextSectionTitle
            if ($null -eq $slice) { return }
            $before = [string]$slice.Before
            $sect   = [string]$slice.Sect
            $after  = [string]$slice.After

            $attrMatches = [regex]::Matches($sect, 'CheckBoxIsChecked="([tf])"', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($attrMatches.Count -gt 0) {
                $limit = [Math]::Min($Checks.Count, $attrMatches.Count)
                for ($i = $limit - 1; $i -ge 0; $i--) {
                    $desiredAttr = if ([bool]$Checks[$i]) { "t" } else { "f" }
                    $am = $attrMatches[$i]
                    $sect = $sect.Substring(0, $am.Index) + ('CheckBoxIsChecked="' + $desiredAttr + '"') + $sect.Substring($am.Index + $am.Length)
                }
                $HtmlRef.Value = $before + $sect + $after
                return
            }

            # Glyph-only template fallback.
            $glyphPattern = (([string][char]0x2610) + "|" + ([string][char]0x25A1) + "|" + ([string][char]0x2611) + "|" + ([string][char]0x2612) + "|" + ([string][char]0x2713) + "|" + ([string][char]0x2714))
            $glyphMatches = [regex]::Matches($sect, $glyphPattern)
            if ($glyphMatches.Count -gt 0) {
                $limitGlyph = [Math]::Min($Checks.Count, $glyphMatches.Count)
                for ($i = $limitGlyph - 1; $i -ge 0; $i--) {
                    $replacementGlyph = if ([bool]$Checks[$i]) { [string][char]0x2611 } else { [string][char]0x2610 }
                    $gm = $glyphMatches[$i]
                    $sect = $sect.Substring(0, $gm.Index) + $replacementGlyph + $sect.Substring($gm.Index + $gm.Length)
                }
                $HtmlRef.Value = $before + $sect + $after
                return
            }

            # HTML entity checkbox fallback.
            $entityPattern = '&#(?:x2610|x25A1|x2611|x2612|x2713|x2714|9744|9633|9745|9746|10003|10004);'
            $entityMatches = [regex]::Matches($sect, $entityPattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($entityMatches.Count -eq 0) { return }
            $limitEntity = [Math]::Min($Checks.Count, $entityMatches.Count)
            for ($i = $limitEntity - 1; $i -ge 0; $i--) {
                $replacementEntity = if ([bool]$Checks[$i]) { '&#9745;' } else { '&#9744;' }
                $em = $entityMatches[$i]
                $sect = $sect.Substring(0, $em.Index) + $replacementEntity + $sect.Substring($em.Index + $em.Length)
            }
            $HtmlRef.Value = $before + $sect + $after
        }

        # Documents Issued For
        $checkInformation = $false
        $checkConstruction = [bool]$script:ocrHasManufactured
        $checkProcurement = $true
        $checkRevision = $false
        $docsEndTitle = "Burn Profiles updated in job folder"
        $okInfo = & $setSectionRowCheckbox ([ref]$h) "Documents Issued For" "Information" $checkInformation $docsEndTitle
        $okConstr = & $setSectionRowCheckbox ([ref]$h) "Documents Issued For" "Construction" $checkConstruction $docsEndTitle
        $okProc = & $setSectionRowCheckbox ([ref]$h) "Documents Issued For" "Procurement" $checkProcurement $docsEndTitle
        $okRev = & $setSectionRowCheckbox ([ref]$h) "Documents Issued For" "Revision" $checkRevision $docsEndTitle
        # Fallback to label-distance mapping when row-label targeting is unavailable.
        # If labels are not discoverable in HTML, use an order-based last resort.
        if (-not ($okInfo -and $okConstr -and $okProc -and $okRev)) {
            $docsAppliedByLabel = & $setSectionCheckboxes ([ref]$h) "Documents Issued For" ([ordered]@{
                "Information" = $checkInformation
                "Construction" = $checkConstruction
                "Procurement" = $checkProcurement
                "Revision" = $checkRevision
            }) $docsEndTitle
            if (-not $docsAppliedByLabel) {
                # Last-resort fallback for templates where labels are unavailable in HTML.
                & $setSectionCheckboxesByOrder ([ref]$h) "Documents Issued For" @([bool]$checkInformation, [bool]$checkConstruction, [bool]$checkProcurement, [bool]$checkRevision) $docsEndTitle
            }
        }
        Write-Log ("  Checkbox targets (Docs): Information={0} Construction={1} Procurement={2} Revision={3}" -f $checkInformation, $checkConstruction, $checkProcurement, $checkRevision) "INFO"

        # Yes/No/N/A sections
        $burnIsNA = ($burnProfilePath -eq "N/A")
        $burnCheckYes = [bool]($burnProfilesReady -and -not $burnIsNA)
        $burnCheckNo  = [bool](-not $burnProfilesReady -and -not $burnIsNA)
        $burnCheckNA  = [bool]$burnIsNA
        $okBurnYes = & $setSectionRowCheckbox ([ref]$h) "Burn Profiles updated in job folder" "Yes" $burnCheckYes "CADLink completed in job folder"
        $okBurnNo  = & $setSectionRowCheckbox ([ref]$h) "Burn Profiles updated in job folder" "No"  $burnCheckNo  "CADLink completed in job folder"
        $okBurnNA  = & $setSectionRowCheckbox ([ref]$h) "Burn Profiles updated in job folder" "N/A" $burnCheckNA  "CADLink completed in job folder"
        if (-not ($okBurnYes -and $okBurnNo -and $okBurnNA)) {
            & $setSectionCheckboxesByOrder ([ref]$h) "Burn Profiles updated in job folder" @($burnCheckYes, $burnCheckNo, $burnCheckNA) "CADLink completed in job folder"
        }
        $okCadYes = & $setSectionRowCheckbox ([ref]$h) "CADLink completed in job folder" "Yes" ([bool]$cadLinkReady) "Notes / Reason for Change"
        $okCadNo = & $setSectionRowCheckbox ([ref]$h) "CADLink completed in job folder" "No" ([bool](-not $cadLinkReady)) "Notes / Reason for Change"
        if (-not ($okCadYes -and $okCadNo)) {
            & $setSectionCheckboxesByOrder ([ref]$h) "CADLink completed in job folder" @([bool]$cadLinkReady, [bool](-not $cadLinkReady)) "Notes / Reason for Change"
        }
        Write-Log ("  Checkbox targets (Burn/CAD): BurnYes={0} BurnNo={1} BurnNA={2} CADYes={3} CADNo={4}" -f $burnCheckYes, $burnCheckNo, $burnCheckNA, [bool]$cadLinkReady, [bool](-not $cadLinkReady)) "INFO"

        # Notes - insert into the WHITE content row BELOW the orange "Notes / Reason for Change" header.
        # Structure: Row 1 (orange header td) then Row 2 (white body td where notes belong).
        $notesHdrIdx = $h.IndexOf('Notes / Reason for Change', [System.StringComparison]::OrdinalIgnoreCase)
        if ($notesHdrIdx -ge 0) {
            $closeTdIdx = $h.IndexOf('</td>', $notesHdrIdx, [System.StringComparison]::OrdinalIgnoreCase)
            if ($closeTdIdx -ge 0) {
                $closeRowIdx = $h.IndexOf('</tr>', $closeTdIdx, [System.StringComparison]::OrdinalIgnoreCase)
                if ($closeRowIdx -ge 0) {
                    $bodyTdMatch = [regex]::Match($h.Substring($closeRowIdx + 5), '(?i)<td\b')
                    if ($bodyTdMatch.Success) {
                        $bodyTdStart = $closeRowIdx + 5 + $bodyTdMatch.Index
                        $bodyTdClose = $h.IndexOf('>', $bodyTdStart)
                        if ($bodyTdClose -ge 0) {
                            $insertAt = $bodyTdClose + 1
                            $notesHtml = @(
                                '<div style="margin:0; padding:2px 0 4px 0; font-family:Calibri,Arial,sans-serif; color:black; font-size:12pt; line-height:1.2; font-weight:700;">'
                                ($notesLines | ForEach-Object {
                                    $escaped = [System.Security.SecurityElement]::Escape($_)
                                    '<div style="margin:0 0 2px 0; padding:0; color:black; font-size:12pt; line-height:1.2; font-weight:700;">' + $escaped + '</div>'
                                }) -join ""
                                '</div>'
                            ) -join ""
                            $h = $h.Substring(0, $insertAt) + $notesHtml + $h.Substring($insertAt)
                        }
                    }
                }
            }
        }

        if ($h -ne $htmlTemplate) {
            $mail.HTMLBody = $h
            Write-Log "  HTMLBody populated (job/transmittal/checkboxes/notes/paths)." "INFO"
        }
    }

    # Subject format aligned to transmittal convention.
    $mail.Subject = "Document Transmittal ($effectiveJob-$transmittalNo)"

    $savedToDrafts = $false
    $draftEntryId = ""
    if ($DispatchMode -in @("Manual","Hold")) {
        try {
            $mail.Save()
            $savedToDrafts = $true
            try { $draftEntryId = [string]$mail.EntryID } catch { $draftEntryId = "" }
            Write-Log "Transmittal draft saved to Outlook Drafts (Mode=$DispatchMode)" "SUCCESS"
        } catch {
            Write-Log "Transmittal draft save failed (Mode=$DispatchMode): $($_.Exception.Message)" "ERROR"
            throw
        }
        if ($DispatchMode -eq "Manual") {
            try {
                $mail.Display()
                Write-Log "Manual mode: transmittal draft displayed for review" "SUCCESS"
            } catch {
                Write-Log "Manual mode: draft display failed: $($_.Exception.Message)" "WARN"
            }
        }
    } else {
        $mail.Send()
        Write-Log "Transmittal email sent to $mailTo" "SUCCESS"
    }
    return [pscustomobject]@{
        TransmittalNo = $transmittalNo
        IsCorrection  = [bool]$isCorrectionTransmittal
        DraftEntryId  = $draftEntryId
        DispatchMode  = $DispatchMode
    }
}

# ==============================================================================
#  Outlook Logic
# ==============================================================================

function Get-OutlookFolder {
    param([object]$Namespace, [string]$FolderPath)
    $parts = $FolderPath -split '\\'; $targetName = $parts[-1]
    $script:foundFolder = $null
    function Search-Folders([object]$Parent, [string]$Name) {
        if ($script:foundFolder) { return }
        foreach ($f in $Parent.Folders) {
            if ($f.Name -eq $Name) { $script:foundFolder = $f; return }
            Search-Folders -Parent $f -Name $Name
        }
    }
    try {
        $folder = $Namespace.DefaultStore.GetRootFolder()
        foreach ($p in $parts) { $folder = $folder.Folders.Item($p) }
        if ($folder) { return $folder }
    } catch { }
    Search-Folders -Parent $Namespace.DefaultStore.GetRootFolder() -Name $targetName
    return $script:foundFolder
}

function Ensure-OutlookFolder {
    param([object]$Namespace, [string]$FolderPath)
    $f = Get-OutlookFolder -Namespace $Namespace -FolderPath $FolderPath
    if ($f) { return $f }
    try {
        $root = $Namespace.DefaultStore.GetRootFolder()
        return $root.Folders.Add($FolderPath)
    } catch { return $null }
}

function Get-PdfPathFromEmail {
    param([string]$EmailBody, [string]$EmailSubject, [object]$Attachments)
    # 1. Check attachments for Sales Order PDF
    foreach ($att in $Attachments) {
        if ($att.FileName -match '(?i)Sales\s*Order\s*Acknowledgment.*\.pdf') {
            # Save to temp and return path
            $tempPath = Join-Path $env:TEMP $att.FileName
            $att.SaveAsFile($tempPath)
            return $tempPath
        }
    }
    # 2. Check body for J:\ or other paths
    $patterns = @(
        '(?mi)([A-Z]:\\[^\r\n]+Sales\s*Order\s*Acknowledgment[^\r\n]+\.pdf)',
        '(?mi)([A-Z]:\\[^\r\n]+\d{5}[^\r\n]+\.pdf)'
    )
    foreach ($p in $patterns) {
        if ($EmailBody -match $p) { return $matches[1].Trim() }
        if ($EmailSubject -match $p) { return $matches[1].Trim() }
    }
    return $null
}

function Extract-PartsFromPdf {
    param([string]$PdfPath)
    if (-not (Test-Path $PdfPath)) { return @() }

    # Find pdftotext - check common locations
    $pdftotext = @(
        "C:\Program Files\poppler\bin\pdftotext.exe",
        "C:\Program Files\Git\mingw64\bin\pdftotext.exe",
        "C:\Program Files (x86)\poppler\bin\pdftotext.exe"
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    $ghostscript = "C:\Program Files\gs\gs9.56.1\bin\gswin64c.exe"
    $tesseract = "C:\Program Files\Tesseract-OCR\tesseract.exe"
    $pdfText = ""

    # Strategy 1: pdftotext (works for digitally-generated PDFs like Epicor orders)
    if ($pdftotext -and (Test-Path $pdftotext)) {
        try {
            $pdfText = & $pdftotext -layout $PdfPath - 2>$null
            if ($pdfText -is [array]) { $pdfText = $pdfText -join "`n" }
            Write-Log "  pdftotext extracted $($pdfText.Length) chars" "INFO"
        } catch { $pdfText = "" }
    }

    # Strategy 2: Ghostscript render to image + Tesseract OCR (for scanned PDFs)
    if ([string]::IsNullOrWhiteSpace($pdfText) -and (Test-Path $ghostscript) -and (Test-Path $tesseract)) {
        try {
            $tmpImg = Join-Path $env:TEMP "so_pdf_render_$(Get-Random).png"
            & $ghostscript -q -dNOPAUSE -dBATCH -sDEVICE=png16m -r300 -dFirstPage=1 -dLastPage=5 "-sOutputFile=$tmpImg" $PdfPath 2>$null
            if (Test-Path $tmpImg) {
                $ocrOut = Join-Path $env:TEMP "so_ocr_$(Get-Random)"
                & $tesseract $tmpImg $ocrOut --psm 6 2>$null
                $ocrFile = "$ocrOut.txt"
                if (Test-Path $ocrFile) {
                    $pdfText = Get-Content $ocrFile -Raw
                    Remove-Item $ocrFile -Force -ErrorAction SilentlyContinue
                }
                Remove-Item $tmpImg -Force -ErrorAction SilentlyContinue
                Write-Log "  Ghostscript+Tesseract extracted $($pdfText.Length) chars" "INFO"
            }
        } catch { }
    }

    if ([string]::IsNullOrWhiteSpace($pdfText)) {
        Write-Log "  Could not extract text from PDF: $PdfPath" "ERROR"
        return @()
    }

    # Parse the extracted text for Sales Order parts
    $lines = $pdfText -split "`n"
    $parts = [System.Collections.Generic.List[object]]::new()
    $orderNumber = "UNKNOWN"
    $inTable = $false
    $currentPart = $null
    $skipDescPrefixes = @('REL ', 'NEED BY', 'SHIP BY', 'QUANTITY', 'UNIT PRICE', 'EXT. PRICE', 'PAGE ', 'SALES ORDER ACKNOWLEDG', 'ORDERACK:', 'LINE TOTAL', 'TOTAL TAX', 'ORDER TOTAL', 'CANADIAN DOLLARS', 'EXT. ', 'LINE MISCELLANEOUS', 'ORDER MISCELLANEOUS')
    $normalizeQtyText = {
        param([string]$q)
        if ([string]::IsNullOrWhiteSpace($q)) { return "" }
        $x = ([string]$q).Trim()
        $x = $x -replace ',', '.'
        $x = $x -replace '\.0+$', ''
        return $x
    }
    $parsePdfRow = {
        param([string]$rowText)
        if ([string]::IsNullOrWhiteSpace($rowText)) { return $null }
        $t = ($rowText -replace '\s+$', '').Trim()
        $m = [regex]::Match(
            $t,
            '^(?<line>\d{1,3})\s+(?<part>[A-Z0-9][A-Z0-9._-]{3,})\s+(?<rev>[A-Z0-9]+)\b',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
        if (-not $m.Success) { return $null }

        $lineNo = [string]$m.Groups['line'].Value.Trim()
        $partNum = [string]$m.Groups['part'].Value.Trim().ToUpperInvariant()
        $rev = [string]$m.Groups['rev'].Value.Trim().ToUpperInvariant()

        if ($partNum -notmatch '^[A-Z0-9][A-Z0-9._-]{3,}$') { return $null }
        if ($partNum -match '(?i)^(JOB|TOTAL|SUB|TAX|NET|FREIGHT|CHARGES)') { return $null }
        # Reject date-like tokens (e.g., "22-Jun-2026") that pdftotext can align to look like part numbers
        if ($partNum -match '^\d{1,2}-[A-Z]{3}-\d{2,4}$') { return $null }
        if ($rev -notmatch '^[A-Z0-9]+$') { return $null }

        # Check if captured "rev" is actually the integer part of a decimal qty (e.g., "3" from "3.00EA")
        $revEndPos = $m.Groups['rev'].Index + $m.Groups['rev'].Length
        if ($revEndPos -lt $t.Length) {
            $afterRev = $t.Substring($revEndPos)
            if ($afterRev -match '^[.,]\d') {
                # This "rev" is the integer part of a decimal number (qty/price), not a revision
                $rev = ""
            }
        }

        $qty = ""
        $qtyMatch = [regex]::Match($t, '(?<qty>\d+(?:[.,]\d+)?)EA\b', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($qtyMatch.Success) {
            $qty = & $normalizeQtyText $qtyMatch.Groups['qty'].Value
        }

        return [ordered]@{
            Line = $lineNo
            Part = $partNum
            Rev  = $rev
            Qty  = $qty
        }
    }
    $isLikelyDescriptionLine = {
        param([string]$text)
        if ([string]::IsNullOrWhiteSpace($text)) { return $false }
        $u = $text.Trim().ToUpperInvariant()
        if ($u -match '^\d{1,3}\s+') { return $false }
        foreach ($prefix in $skipDescPrefixes) {
            if ($u.StartsWith($prefix)) { return $false }
        }
        # Reject page headers/footers that can appear mid-page in pdftotext output
        if ($u -match '(?i)SALES\s+ORDER\s+ACKNOWLEDG') { return $false }
        if ($u -match '\d+\s+OF\s+\d+') { return $false }
        if ($u -match '^\d+\s+\d{1,2}-[A-Z]{3}-\d{4}\b') { return $false }
        if ($u -match '^\d[\d\s,\.]+$') { return $false }
        # If line starts with alpha text, strip trailing price columns before final check
        if ($u -match '^[A-Z]') {
            return ($u -match '[A-Z]{3,}')
        }
        if ($u -match '\b\d+(?:[.,]\d+)?EA\b') { return $false }
        if ($u -match '\b\d{1,3}(?:,\d{3})+\.\d{2}\b') { return $false }
        return ($u -match '[A-Z]{3,}')
    }
    $isPriceNoiseLine = {
        param([string]$text)
        if ([string]::IsNullOrWhiteSpace($text)) { return $false }
        $u = $text.Trim().ToUpperInvariant()
        # If the line starts with 3+ alpha chars, it likely has a description with trailing price columns — not pure noise
        if ($u -match '^[A-Z]{3,}') { return $false }
        if ($u -match '\b\d+(?:[.,]\d+)?EA\b') { return $true }
        if ($u -match '\b\d{1,3}(?:,\d{3})*\.\d{2,3}\b') { return $true }
        if ($u -match '^\d[\d\s,\.EA]+$') { return $true }
        return $false
    }
    # Strip trailing price/amount noise from a line, keeping only the leading text (description)
    $stripTrailingPriceNoise = {
        param([string]$text)
        if ([string]::IsNullOrWhiteSpace($text)) { return "" }
        # Remove trailing price-like segments: "5,125.65", "1,657.02", "0.00", "276.170", "3.00EA"
        $cleaned = [regex]::Replace($text, '\s+\d+(?:[.,]\d+)?EA\b.*$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $cleaned = [regex]::Replace($cleaned, '\s+\d{1,3}(?:,\d{3})*\.\d{2,3}\s*$', '')
        return $cleaned.Trim()
    }
    $tryParseQuantityLine = {
        param([string]$text)
        if ([string]::IsNullOrWhiteSpace($text)) { return "" }
        # Match: rel# date date qty  (with optional trailing price noise from pdftotext -layout)
        $m = [regex]::Match($text.Trim(), '^\d+\s+\d{1,2}-[A-Z]{3}-\d{4}\s+\d{1,2}-[A-Z]{3}-\d{4}\s+(?<qty>\d+(?:[.,]\d+)?)', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($m.Success) { return (& $normalizeQtyText $m.Groups['qty'].Value) }
        return ""
    }
    $flushCurrentPart = {
        if ($null -eq $currentPart) { return }
        $desc = [string]$currentPart.Description
        $desc = [regex]::Replace($desc, '\s+', ' ').Trim()
        $currentPart.Description = $desc
        [void]$parts.Add([PSCustomObject]@{
            Part         = [string]$currentPart.Part
            Order        = [string]$currentPart.Order
            Line         = [string]$currentPart.Line
            Rev          = [string]$currentPart.Rev
            Qty          = [string]$currentPart.Qty
            Description  = [string]$currentPart.Description
            InternalPart = [string]$currentPart.InternalPart
        })
    }

    foreach ($line in $lines) {
        $line = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        if ($line -match '(?i)Sales Order:\s*(\d+)') {
            $orderNumber = $matches[1]
        }

        # Detect table start
        if ($line -match '(?i)Line\s+Part\s+Number') {
            $inTable = $true
            continue
        }

        if ($inTable) {
            $parsedRow = & $parsePdfRow $line
            if ($null -ne $parsedRow) {
                & $flushCurrentPart
                $currentPart = [ordered]@{
                    Part         = [string]$parsedRow.Part
                    Order        = $orderNumber
                    Line         = [string]$parsedRow.Line
                    Rev          = [string]$parsedRow.Rev
                    Qty          = [string]$parsedRow.Qty
                    Description  = ""
                    InternalPart = ""
                    AwaitQty     = $false
                }
                continue
            }

            if ($null -ne $currentPart) {
                if ($line -match '^(?i)Rel\s+Need\s+By\s+Ship\s+By\s+Quantity\b') {
                    $currentPart.AwaitQty = $true
                    continue
                }
                if ([bool]$currentPart.AwaitQty) {
                    $qtyFromLine = & $tryParseQuantityLine $line
                    if (-not [string]::IsNullOrWhiteSpace($qtyFromLine)) {
                        $currentPart.Qty = $qtyFromLine
                        $currentPart.AwaitQty = $false
                        continue
                    }
                }
                if ($line -match '^\d{1,3}\s+\S+') {
                    & $flushCurrentPart
                    $currentPart = $null
                    continue
                }
                # Detect "Our Part:" cross-reference to internal NMT part number
                if ($line -match '(?i)(?:Our\s+Part|NMT\s+Part|Internal\s+Part)[:#\s]+\s*([A-Z0-9][A-Z0-9._-]{3,})') {
                    $currentPart.InternalPart = $matches[1].Trim().ToUpperInvariant()
                    continue
                }
                if (& $isPriceNoiseLine $line) {
                    continue
                }
                if (& $isLikelyDescriptionLine $line) {
                    $descText = & $stripTrailingPriceNoise $line
                    if (-not [string]::IsNullOrWhiteSpace($descText)) {
                        if ([string]::IsNullOrWhiteSpace($currentPart.Description)) {
                            $currentPart.Description = $descText
                        } else {
                            $currentPart.Description += " " + $descText
                        }
                    }
                }
            }
        }
    }

    & $flushCurrentPart

    Write-Log "  Extracted $($parts.Count) parts from Sales Order $orderNumber" "INFO"
    return @($parts)
}

function Get-HistoryClientName {
    param(
        [string]$CurrentClientName,
        [string]$JobFolderName = "",
        [string]$DocxPath = ""
    )

    if (-not [string]::IsNullOrWhiteSpace($CurrentClientName) -and $CurrentClientName -ne "Epicor Order") {
        return $CurrentClientName.Trim()
    }

    foreach ($candidate in @($JobFolderName, [System.IO.Path]::GetFileNameWithoutExtension($DocxPath))) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        $m = [regex]::Match($candidate, '^\s*\d{4,6}\s*-\s*(.+?)\s*-\s*')
        if ($m.Success) {
            $value = $m.Groups[1].Value.Trim()
            if (-not [string]::IsNullOrWhiteSpace($value)) { return $value }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($CurrentClientName)) { return $CurrentClientName.Trim() }
    return ""
}

function Get-HistoryOrderName {
    param(
        [string]$Subject,
        [string]$JobNumber,
        [string]$ClientName,
        [string]$DocxPath = "",
        [string]$SourcePdfPath = ""
    )

    if (-not [string]::IsNullOrWhiteSpace($DocxPath)) {
        $docBase = [System.IO.Path]::GetFileNameWithoutExtension($DocxPath)
        $docBase = [regex]::Replace($docBase, '\s*-\s*F80$', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase).Trim()
        if (-not [string]::IsNullOrWhiteSpace($docBase)) { return $docBase }
    }

    if (-not [string]::IsNullOrWhiteSpace($SourcePdfPath)) {
        $pdfBase = [System.IO.Path]::GetFileNameWithoutExtension($SourcePdfPath).Trim()
        if (-not [string]::IsNullOrWhiteSpace($pdfBase)) { return $pdfBase }
    }

    if (-not [string]::IsNullOrWhiteSpace($JobNumber) -and -not [string]::IsNullOrWhiteSpace($ClientName)) {
        return "$JobNumber - $ClientName"
    }

    if (-not [string]::IsNullOrWhiteSpace($JobNumber)) { return "Job $JobNumber" }
    if (-not [string]::IsNullOrWhiteSpace($Subject)) { return $Subject.Trim() }
    return "Order"
}

function Get-CleanNotFoundParts {
    param([string[]]$Parts = @())

    $clean = [System.Collections.Generic.List[string]]::new()
    foreach ($raw in @($Parts)) {
        $part = [string]$raw
        if ([string]::IsNullOrWhiteSpace($part)) { continue }
        $part = $part.Trim().ToUpperInvariant()
        if ($part.Length -lt 4 -or $part.Length -gt 40) { continue }
        if ($part -match '^\d{1,2}[-/][A-Z]{3}[-/]\d{2,4}$') { continue }
        if ($part -match '^[A-Z]{3,10}\d{1,2}[-/][A-Z]{3}[-/]\d{2,4}$') { continue }
        if ($part -match '^\d{1,3}(?:,\d{3})*(?:\.\d+)?$') { continue }
        if ($part -match '^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$') { continue }
        if (Test-DrawingLikePartNumber -PartNumber $part -or Test-HardwareLikePartNumber -PartNumber $part) {
            if (-not $clean.Contains($part)) { [void]$clean.Add($part) }
        }
    }
    return $clean.ToArray()
}

function Convert-OrderLinesToHistoryParts {
    param([object[]]$OrderLines = @())

    $parts = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    foreach ($line in @($OrderLines)) {
        if ($null -eq $line) { continue }
        $part = [string]$line.Part
        if ([string]::IsNullOrWhiteSpace($part)) { continue }
        $part = $part.Trim().ToUpperInvariant()
        if ($seen.ContainsKey($part)) { continue }
        $seen[$part] = $true
        if (-not (Test-DrawingLikePartNumber -PartNumber $part) -and -not (Test-HardwareLikePartNumber -PartNumber $part)) {
            continue
        }
        $parts.Add([ordered]@{
            Part        = $part
            Rev         = [string]$line.Rev
            Qty         = [string]$line.Qty
            Description = [string]$line.Description
        })
    }
    return @($parts)
}

function Convert-OrderLinesToHashtableArray {
    param([object[]]$OrderLines = @())

    $result = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($line in @($OrderLines)) {
        if ($null -eq $line) { continue }
        $item = @{
            Part         = [string]$line.Part
            Description  = [string]$line.Description
            Rev          = [string]$line.Rev
            Qty          = [string]$line.Qty
            Order        = [string]$line.Order
            Line         = [string]$line.Line
            InternalPart = [string]$line.InternalPart
        }
        $result.Add($item)
    }
    return @($result)
}

function Get-DispatchStatusText {
    param(
        [string]$Mode,
        [bool]$Succeeded
    )
    if (-not $Succeeded) {
        switch ($Mode) {
            "Manual" { "Draft Failed" }
            "Hold"   { "Draft Failed" }
            default  { "Transmittal Failed" }
        }
        return
    }
    switch ($Mode) {
        "Manual" { "Draft Ready" }
        "Hold"   { "Draft Saved" }
        default  { "Transmittal Sent" }
    }
}

function Get-HistoryValidatedParts {
    param(
        [object[]]$OrderLines = @(),
        [string[]]$PartNumbers = @(),
        [string]$OrderFolder = ""
    )

    $normalizePartToken = {
        param([string]$s)
        if ([string]::IsNullOrWhiteSpace($s)) { return "" }
        return (($s.ToUpperInvariant()) -replace '[^A-Z0-9]', '')
    }

    $candidateParts = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    foreach ($line in @($OrderLines)) {
        if ($null -eq $line -or -not $line.Part) { continue }
        $pn = ([string]$line.Part).Trim().ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($pn) -or $seen.ContainsKey($pn)) { continue }
        if (-not (Test-DrawingLikePartNumber -PartNumber $pn) -and -not (Test-HardwareLikePartNumber -PartNumber $pn)) { continue }
        $seen[$pn] = $true
        $candidateParts.Add([pscustomobject]@{
            Part = $pn
            Rev  = [string]$line.Rev
            Qty  = [string]$line.Qty
            Description = [string]$line.Description
        })
    }
    foreach ($pnRaw in @($PartNumbers)) {
        $pn = ([string]$pnRaw).Trim().ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($pn) -or $seen.ContainsKey($pn)) { continue }
        if (-not (Test-DrawingLikePartNumber -PartNumber $pn) -and -not (Test-HardwareLikePartNumber -PartNumber $pn)) { continue }
        $seen[$pn] = $true
        $candidateParts.Add([pscustomobject]@{
            Part = $pn
            Rev  = ""
            Qty  = ""
            Description = ""
        })
    }

    $drawMap = @{}
    if (-not [string]::IsNullOrWhiteSpace($OrderFolder) -and (Test-Path $OrderFolder)) {
        $partCandidates = @($candidateParts | ForEach-Object {
            $n = & $normalizePartToken $_.Part
            [pscustomobject]@{ Part = $_.Part; Norm = $n; Len = $n.Length }
        } | Where-Object { $_.Len -ge 4 } | Sort-Object Len -Descending)

        $extractPartRevFromFile = {
            param([string]$LeafName)
            $name = [System.IO.Path]::GetFileNameWithoutExtension($LeafName).ToUpperInvariant()
            $base = $name
            $rev  = ""
            if ($name -match '^(?<base>.+?)_REV(?<rev>[A-Z0-9]+)$') {
                $base = $Matches['base']
                $rev  = $Matches['rev']
            }
            if ($base -match '^(?<base>.+?)-REV(?<rev>[A-Z0-9]+)$') {
                $base = $Matches['base']
                if ([string]::IsNullOrWhiteSpace($rev)) { $rev = $Matches['rev'] }
            }
            $baseNorm = & $normalizePartToken $base
            $matchedPart = ""
            foreach ($cand in $partCandidates) {
                if ([string]::IsNullOrWhiteSpace($cand.Norm)) { continue }
                if ($baseNorm.Contains($cand.Norm)) {
                    $matchedPart = [string]$cand.Part
                    break
                }
            }
            return @{ Base = $matchedPart; Rev = $rev }
        }

        $pdfFiles = @(Get-ChildItem -Path $OrderFolder -File -Filter "*.pdf" -ErrorAction SilentlyContinue)
        $dxfFiles = @(Get-ChildItem -Path (Join-Path $OrderFolder "DXFs") -File -Filter "*.dxf" -ErrorAction SilentlyContinue)
        foreach ($f in $pdfFiles) {
            $pr = & $extractPartRevFromFile $f.Name
            if ([string]::IsNullOrWhiteSpace($pr.Base)) { continue }
            if (-not $drawMap.ContainsKey($pr.Base)) { $drawMap[$pr.Base] = @{ Part = $pr.Base; HasPdf = $false; HasDxf = $false; PdfRev = ""; DxfRev = "" } }
            $drawMap[$pr.Base].HasPdf = $true
            if ($pr.Rev) { $drawMap[$pr.Base].PdfRev = $pr.Rev }
        }
        foreach ($f in $dxfFiles) {
            $pr = & $extractPartRevFromFile $f.Name
            if ([string]::IsNullOrWhiteSpace($pr.Base)) { continue }
            if (-not $drawMap.ContainsKey($pr.Base)) { $drawMap[$pr.Base] = @{ Part = $pr.Base; HasPdf = $false; HasDxf = $false; PdfRev = ""; DxfRev = "" } }
            $drawMap[$pr.Base].HasDxf = $true
            if ($pr.Rev) { $drawMap[$pr.Base].DxfRev = $pr.Rev }
        }
    }

    $validated = [System.Collections.Generic.List[object]]::new()
    foreach ($item in $candidateParts) {
        $part = $item.Part
        $hits = @(Find-PartInHistoryIndex -PartNumber $part)
        $dxfHits = @(Find-PartInHistoryIndex -PartNumber $part -Dxf)
        $hasDrawing = ($hits.Count -gt 0)
        $hasDxf = ($dxfHits.Count -gt 0)
        $indexRev = "N/A"
        $pdfPath = ""
        $dxfPath = ""
        if ($hasDrawing) {
            $bestHit = $hits | Sort-Object {
                $rv = Normalize-Rev $_.Rev
                try { [int]$rv } catch { try { [double]$rv } catch { 0 } }
            } -Descending | Select-Object -First 1
            $indexRev = Normalize-Rev $bestHit.Rev
            $pdfPath = [string]$bestHit.FullPath
        }
        if ($hasDxf) {
            $bestDxf = $dxfHits | Sort-Object {
                $rv = Normalize-Rev $_.Rev
                try { [int]$rv } catch { try { [double]$rv } catch { 0 } }
            } -Descending | Select-Object -First 1
            $dxfPath = [string]$bestDxf.FullPath
        }

        $draw = if ($drawMap.ContainsKey($part)) { $drawMap[$part] } else { $null }
        $fileRev = ""
        if ($draw) {
            $pdfRevNorm = Normalize-Rev ([string]$draw.PdfRev)
            $dxfRevNorm = Normalize-Rev ([string]$draw.DxfRev)
            if ($pdfRevNorm -ne "NA") { $fileRev = $pdfRevNorm }
            elseif ($dxfRevNorm -ne "NA") { $fileRev = $dxfRevNorm }
        }
        if ([string]::IsNullOrWhiteSpace($fileRev) -and $indexRev -ne "N/A") { $fileRev = $indexRev }

        $epicorInfo = Get-EpicorPartInfo -PartNumber $part
        $inEpicor = if ($null -ne $epicorInfo) { [bool]$epicorInfo.Exists } else { $null }
        $epicorRev = if ($null -ne $epicorInfo -and $epicorInfo.Exists) { [string]$epicorInfo.LatestRev } else { "" }
        $effectiveIndexRev = if (-not [string]::IsNullOrWhiteSpace($fileRev)) { $fileRev } else { $indexRev }
        $revCompare = Get-RevisionComparison -OrderRev ([string]$item.Rev) -EpicorRev $epicorRev -IndexRev $effectiveIndexRev

        $validated.Add([ordered]@{
            Part        = $part
            Rev         = [string]$revCompare.OrderRev
            Qty         = [string]$item.Qty
            Description = [string]$item.Description
            EpicorExists = $inEpicor
            EpicorRev   = [string]$revCompare.EpicorRev
            HasDrawing  = $hasDrawing -or ($draw -and $draw.HasPdf)
            HasDxf      = $hasDxf -or ($draw -and $draw.HasDxf)
            IndexRev    = [string]$revCompare.IndexRev
            RevMatch    = [string]$revCompare.Status
            RevNote     = [string]$revCompare.Note
            EpicorMatch = [bool]$revCompare.EpicorMatches
            IndexMatch  = [bool]$revCompare.IndexMatches
            PdfPath     = $pdfPath
            DxfPath     = $dxfPath
        })
    }

    return @($validated)
}

function Process-Order {
    param(
        [object]$MailItem,
        [object]$OutlookApp = $null,
        [string]$DocxPath = ""
    )
    $subject = $MailItem.Subject; $sender = $MailItem.SenderName
    Write-Log ("=" * 60) "INFO"
    Write-Log "Processing Sales Order: '$subject' from $sender" "INFO"

    # Show email body preview for debugging
    $bodyPreview = ($MailItem.Body -replace '[\r\n]+', ' ').Trim()
    $bodyPreview = $bodyPreview.Substring(0, [Math]::Min(300, $bodyPreview.Length))
    Write-Log "  Email body: $bodyPreview" "INFO"

    $script:lastProcessOrderFailReason = ""
    $pdfPath  = Get-PdfPathFromEmail  -EmailBody $MailItem.Body -EmailSubject $subject -Attachments $MailItem.Attachments

    $partNumbers = @(); $orderLines = @(); $jobNumber = ""; $clientName = ""; $resolvedDocPath = ""; $soPdfPath = $null; $jobFolderName = ""

    if ($pdfPath) {
        Write-Log "  Sales Order PDF detected: $pdfPath" "SUCCESS"
        $pdfResults = Extract-PartsFromPdf -PdfPath $pdfPath
        if ($pdfResults) {
            $partNumbers = @($pdfResults | ForEach-Object {
                $_.Part
                if (-not [string]::IsNullOrWhiteSpace($_.InternalPart)) { $_.InternalPart }
            })
            $orderLines = $pdfResults
            $jobNumber = $pdfResults[0].Order
            $clientName = "Epicor Order"
            Write-Log "  Extracted $($partNumbers.Count) parts from Sales Order $jobNumber" "SUCCESS"
        } else {
            Write-Log "  Failed to extract parts from PDF. Possible OCR error." "ERROR"
        }
    } else {
        Write-Log "  No PDF attachment in email - trying PDM docx path flow..." "INFO"

        # Extract docx path from HTMLBody if not already provided
        if ([string]::IsNullOrWhiteSpace($DocxPath)) {
            try {
                $htmlBody = [string]$MailItem.HTMLBody
                $stripped = $htmlBody -replace '<[^>]+>', ' ' -replace '&nbsp;', ' ' -replace '&amp;', '&' -replace '&#92;', '\' -replace '\s+', ' '
                $DocxPath = Get-DocxPathFromEmail -EmailBody $stripped -EmailSubject $subject
            } catch { }
        }

        if ([string]::IsNullOrWhiteSpace($DocxPath)) {
            Write-Log "  No Sales Order PDF or F80 docx path found. Cannot process." "WARN"
            return $false
        }

        Write-Log "  F80 docx path from email: $DocxPath" "INFO"
        # Resolve the docx path relative to PDM vault
        $resolvedDocPath = if ($DocxPath -match '^[A-Z]:\\') { $DocxPath } else { Join-Path $pdmVaultPath $DocxPath.TrimStart('\') }
        Write-Log "  Resolved docx path: $resolvedDocPath" "INFO"

        # Extract job number from the docx filename (e.g. "27129 - PTFI - ... - F80.docx" -> "27129")
        $docxFileName = Split-Path $DocxPath -Leaf
        $jobMatch = [regex]::Match($docxFileName, '(\d{4,6})')
        if (-not $jobMatch.Success) {
            Write-Log "  Could not extract job number from docx filename: $docxFileName" "ERROR"
            return $false
        }
        $jobNumber = $jobMatch.Groups[1].Value
        Write-Log "  Extracted job number: $jobNumber" "SUCCESS"

        # Search for the Sales Order Acknowledgment PDF on the Epicor file server
        foreach ($root in $crawlRoots) {
            $searchPaths = @(
                (Join-Path $root "Orders\Spare Parts"),
                (Join-Path $root "Epicor\Orders\Spare Parts")
            )
            foreach ($sp in $searchPaths) {
                if (-not (Test-Path $sp)) { continue }
                # Find the job folder (e.g. "27000-27199\27129 - PTFI - ...")
                $jobFolders = @(Get-ChildItem -Path $sp -Directory -Recurse -Depth 1 -ErrorAction SilentlyContinue |
                    Where-Object { $_.Name -match "^$jobNumber\b" })
                foreach ($jf in $jobFolders) {
                    $salesDir = Join-Path $jf.FullName "50 - Sales"
                    if (Test-Path $salesDir) {
                        $soPdfs = @(Get-ChildItem -Path $salesDir -Filter "Sales Order Acknowledgment*.pdf" -ErrorAction SilentlyContinue)
                        if ($soPdfs.Count -eq 0) {
                            $soPdfs = @(Get-ChildItem -Path $salesDir -Filter "Sales Order*.pdf" -ErrorAction SilentlyContinue)
                        }
                        if ($soPdfs.Count -gt 0) {
                            $soPdfPath = ($soPdfs | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
                            $jobFolderName = $jf.Name
                            Write-Log "  Found Sales Order PDF: $soPdfPath" "SUCCESS"
                            break
                        }
                    }
                }
                if ($soPdfPath) { break }
            }
            if ($soPdfPath) { break }
        }

        if (-not $soPdfPath) {
            Write-Log "  Sales Order PDF not found on file server for job $jobNumber" "WARN"
            $script:lastProcessOrderFailReason = "FileNotFound"
            return $false
        }

        # Extract parts from the Sales Order PDF
        $pdfResults = Extract-PartsFromPdf -PdfPath $soPdfPath
        if ($pdfResults) {
            $partNumbers = @($pdfResults | ForEach-Object {
                $_.Part
                if (-not [string]::IsNullOrWhiteSpace($_.InternalPart)) { $_.InternalPart }
            })
            $orderLines = $pdfResults
            if (-not $jobNumber) { $jobNumber = $pdfResults[0].Order }
            $clientName = Get-HistoryClientName -CurrentClientName $clientName -JobFolderName $jobFolderName -DocxPath $resolvedDocPath
            if ([string]::IsNullOrWhiteSpace($clientName)) { $clientName = "Epicor Order" }
            Write-Log "  Extracted $($partNumbers.Count) parts from Sales Order $jobNumber" "SUCCESS"
        } else {
            Write-Log "  Failed to extract parts from Sales Order PDF: $soPdfPath" "ERROR"
            return $false
        }
    }

    if ($partNumbers.Count -eq 0) { 
        Write-Log "No parts found in order source. Skipping." "WARN"
        return $false 
    }

    # Expand assembly trees (subassemblies/components) before collection so
    # PDF/DXF capture includes full model content, not only top-level F80 rows.
    if ($enableAssemblyBomExpansion) {
        $PdfIndexPath = Join-Path $indexFolder "pdf_index_clean.csv"
        $ocrPartCount = $partNumbers.Count
        $partNumbers = Expand-AssemblyBOM -PartNumbers $partNumbers -PdfIndexPath $PdfIndexPath -JobNumber $jobNumber `
            -OrderDocPath $resolvedDocPath -CrawlRoots $crawlRoots
        Write-Log "Parts: $ocrPartCount from F80 -> $($partNumbers.Count) after BOM expansion" "INFO"
    } else {
        Write-Log "BOM expansion disabled by config (emailMonitor.enableAssemblyBomExpansion=false)." "WARN"
    }

    $folderName = "$(Get-Date -Format 'yyyyMMdd_HHmmss')_$($subject -replace '[^\w]','_')"
    $orderFolder = Join-Path $outputRoot "Orders\$folderName"
    New-Item -ItemType Directory -Path $orderFolder -Force | Out-Null

    $bomFile = Join-Path $orderFolder "order_bom.txt"
    $partNumbers | Set-Content -Path $bomFile -Encoding UTF8

    Write-Log "Running Collector for $($partNumbers.Count) parts..." "INFO"
    $collectorArgs = @(
        "-ExecutionPolicy", "Bypass",
        "-File", (Join-Path $scriptDir "SimpleCollector.ps1"),
        $bomFile,
        $orderFolder,
        $collectMode,
        $configPath
    )
    $collectorTimeoutMinutes = 12
    $proc = Start-Process -FilePath "powershell.exe" -ArgumentList $collectorArgs -PassThru -NoNewWindow
    if (-not $proc.WaitForExit($collectorTimeoutMinutes * 60 * 1000)) {
        Write-Log "Collector timeout after $collectorTimeoutMinutes minute(s); terminating collector process (PID=$($proc.Id))." "ERROR"
        try { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue } catch { }
        throw "Collector timed out after $collectorTimeoutMinutes minute(s)."
    }
    Write-Log "Collector finished with exit code: $($proc.ExitCode)" "INFO"

    # --- Read collector results for transmittal ---
    $summaryPath = Join-Path $env:TEMP "collector_summary.json"
    $pdfsFound = 0; $dxfsFound = 0; $notFoundParts = @()
    if (Test-Path $summaryPath) {
        try {
            $summary = Get-Content $summaryPath -Raw | ConvertFrom-Json
            $pdfsFound = $summary.pdfsFound
            $dxfsFound = $summary.dxfsFound
            $notFoundParts = @($summary.notFound)
        } catch { Write-Log "Could not read collector summary" "WARN" }
    }

    # --- Send Transmittal Email ---
    $transmittalOk = $false
    $transmittalNoUsed = ""
    $draftEntryId = ""
    $statusMsg = Get-DispatchStatusText -Mode $DispatchMode -Succeeded $true
    $effectiveOrderLines = if (@($orderLines).Count -gt 0) { @(Convert-OrderLinesToHashtableArray -OrderLines $orderLines) } else { @() }
    try {
        $transmittalMeta = Send-TransmittalEmail -Subject $subject -JobNumber $jobNumber -ClientName $clientName `
            -PartNumbers $partNumbers -OrderLines $effectiveOrderLines `
            -OrderFolder $orderFolder -OrderDocPath $resolvedDocPath `
            -PdfsFound $pdfsFound -DxfsFound $dxfsFound -NotFound $notFoundParts `
            -TestMode $TestMode -DispatchMode $DispatchMode -OutlookApp $OutlookApp
        if ($transmittalMeta -and $transmittalMeta.PSObject.Properties['TransmittalNo']) {
            $transmittalNoUsed = [string]$transmittalMeta.TransmittalNo
        }
        if ($transmittalMeta -and $transmittalMeta.PSObject.Properties['DraftEntryId']) {
            $draftEntryId = [string]$transmittalMeta.DraftEntryId
        }
        $transmittalOk = $true
    } catch {
        Write-Log "Failed to create transmittal email: $($_.Exception.Message)" "ERROR"
        $statusMsg = Get-DispatchStatusText -Mode $DispatchMode -Succeeded $false
    }

    # --- Write dashboard summary so the Hub UI can show what happened ---
    try {
        $historyClient = Get-HistoryClientName -CurrentClientName $clientName -JobFolderName $jobFolderName -DocxPath $resolvedDocPath
        $historyOrderName = Get-HistoryOrderName -Subject $subject -JobNumber $jobNumber -ClientName $historyClient -DocxPath $resolvedDocPath -SourcePdfPath $(if ($soPdfPath) { $soPdfPath } else { $pdfPath })
        [string[]]$nfArr = @(Get-CleanNotFoundParts -Parts @($notFoundParts) | Select-Object -First 10)
        if ($null -eq $nfArr) { $nfArr = @() }
        $historyParts = @(Get-HistoryValidatedParts -OrderLines $effectiveOrderLines -PartNumbers $partNumbers -OrderFolder $orderFolder | Select-Object -First 50)
        $flaggedParts = @($historyParts | Where-Object { $_.RevMatch -eq "Mismatch" } | ForEach-Object {
            $systemRev = if ($_.EpicorRev) { $_.EpicorRev } elseif ($_.IndexRev) { $_.IndexRev } else { "N/A" }
            "$($_.Part): Order Rev $($_.Rev) vs System Rev $systemRev"
        })
        $dispatchState = if ($transmittalOk) {
            switch ($DispatchMode) {
                "Manual" { "Draft Ready" }
                "Hold"   { "Draft Saved" }
                default  { "Sent" }
            }
        } else {
            switch ($DispatchMode) {
                "Manual" { "Draft Failed" }
                "Hold"   { "Draft Failed" }
                default  { "Send Failed" }
            }
        }

        $dashSummary = [ordered]@{
            Timestamp       = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Status          = $statusMsg
            OrderName       = $historyOrderName
            EmailSubject    = $subject
            Sender          = $sender
            JobNumber       = $jobNumber
            Client          = $historyClient
            PartsExtracted  = $partNumbers.Count
            OrderLinesCount = @($effectiveOrderLines).Count
            Parts           = $historyParts
            FlaggedParts    = $flaggedParts
            PdfsCollected   = $pdfsFound
            DxfsCollected   = $dxfsFound
            TransmittalSent = ($DispatchMode -eq "Auto" -and $transmittalOk)
            DraftCreated    = ($DispatchMode -in @("Manual","Hold") -and $transmittalOk)
            TestMode        = [bool]$TestMode
            DispatchMode    = $DispatchMode
            DispatchState   = $dispatchState
            TransmittalNo   = $transmittalNoUsed
            DraftEntryId    = $draftEntryId
            NotFoundParts   = $nfArr
            SourceDocPath   = $resolvedDocPath
            SourcePdfPath   = $(if ($soPdfPath) { $soPdfPath } else { $pdfPath })
            OutputFolder    = $orderFolder
        }

        # Write last_email_summary.json — this single object is safe for ConvertTo-Json -Depth 5
        # because it's a flat hashtable (arrays of strings inside hashtables are fine).
        $summaryPath2 = Join-Path $indexFolder "last_email_summary.json"
        Write-Log "Writing dashboard summary to: $summaryPath2" "INFO"
        $dashSummary | ConvertTo-Json -Depth 5 | Set-Content -Path $summaryPath2 -Encoding UTF8 -Force
        Write-Log "Dashboard summary written OK" "INFO"

        # ============================================================
        # Maintain persistent history (last 50 orders)
        # BULLETPROOF PS 5.x approach:
        #   1. Read existing history JSON file as raw text
        #   2. Parse with ConvertFrom-Json, normalize to a proper list
        #   3. Serialize EACH entry individually with ConvertTo-Json -Compress
        #   4. Join them into a hand-built JSON array string
        #   5. Write the raw string directly — never pass an array through ConvertTo-Json
        # ============================================================
        $historyFile = Join-Path $indexFolder "transmittal_history.json"
        $historyList = [System.Collections.Generic.List[object]]::new()
        if (Test-Path $historyFile) {
            try {
                $rawHist = (Get-Content $historyFile -Raw).Trim()
                if (-not [string]::IsNullOrWhiteSpace($rawHist) -and $rawHist -ne "[]") {
                    $parsed = $rawHist | ConvertFrom-Json
                    if ($null -ne $parsed) {
                        # ConvertFrom-Json returns a single PSCustomObject for 1-item arrays
                        if ($parsed -is [System.Array]) {
                            foreach ($item in $parsed) { $historyList.Add($item) }
                        } else {
                            $historyList.Add($parsed)
                        }
                    }
                }
            } catch {
                Write-Log "Could not parse existing history file, starting fresh: $($_.Exception.Message)" "WARN"
            }
        }

        # Serialize each entry individually — this is the ONLY safe approach in PS 5.x
        $jsonLines = [System.Collections.Generic.List[string]]::new()

        # New entry first (prepend)
        $jsonLines.Add(($dashSummary | ConvertTo-Json -Depth 5 -Compress))

        # Existing entries
        foreach ($h in $historyList) {
            # Skip any corrupt entries that lack a Timestamp (e.g. {"value":[],"Count":N} ghosts)
            if ($null -ne $h -and $null -ne $h.Timestamp) {
                $jsonLines.Add(($h | ConvertTo-Json -Depth 5 -Compress))
            }
        }

        # Cap at 50 entries
        if ($jsonLines.Count -gt 50) {
            $jsonLines = [System.Collections.Generic.List[string]]::new($jsonLines.GetRange(0, 50))
        }

        # Build the JSON array string manually
        $histContent = if ($jsonLines.Count -gt 0) { "[`n" + ($jsonLines -join ",`n") + "`n]" } else { "[]" }

        # Write using .NET directly to avoid any PowerShell pipeline quirks
        [System.IO.File]::WriteAllText($historyFile, $histContent, [System.Text.Encoding]::UTF8)
        Write-Log "Transmittal history updated ($($historyList.Count + 1) entries)" "INFO"
    } catch {
        Write-Log "ERROR writing dashboard summary/history: $($_.Exception.Message)" "ERROR"
    }

    # --- Push tray notification so the user sees the result immediately ---
    try {
        $notifyTitle = if ($TestMode) { "Draft Ready: $jobNumber" } else { "Order Complete: $jobNumber" }
        if ($transmittalOk) {
            $notifyMsg = "$statusMsg - $pdfsFound PDF(s), $dxfsFound DXF(s) collected"
        } elseif (($pdfsFound + $dxfsFound) -gt 0) {
            $notifyMsg = "$pdfsFound PDF(s), $dxfsFound DXF(s) collected - $statusMsg"
        } else {
            $notifyMsg = "No drawings found for $($partNumbers.Count) part(s) - $statusMsg"
        }
        Push-HubNotification -Title $notifyTitle -Message $notifyMsg -FolderPath $orderFolder
    } catch { }

    return $transmittalOk
}

function Run-OrderCheck {
    Write-Log "Checking for orders..." "INFO"
    Write-EmailProgress -Step "scanning" -Detail "Connecting to Outlook..."
    $outlook = $null
    try {
        $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    } catch {
        try {
            $outlook = New-Object -ComObject Outlook.Application
        } catch {
            $msg = $_.Exception.Message
            Write-Log "Outlook COM unavailable: $msg" "ERROR"
            if ($msg -match '80070520|0x80070520') {
                Write-Log "Outlook logon session missing. Open Outlook desktop in the same Windows user session and run this monitor from a non-admin terminal." "WARN"
            }
            return
        }
    }
    $namespace = $outlook.GetNamespace("MAPI")

    if ($forceSendReceiveBeforeScan) {
        try {
            Write-Log "Forcing Outlook Send/Receive before unread scan..." "INFO"
            $namespace.SendAndReceive($false)
            if ($sendReceiveWaitSeconds -gt 0) {
                Start-Sleep -Seconds $sendReceiveWaitSeconds
            }
        } catch {
            Write-Log "Outlook Send/Receive refresh failed: $($_.Exception.Message)" "WARN"
        }
    }

    $inbox = Ensure-OutlookFolder -Namespace $namespace -FolderPath $monitorFolder
    if (-not $inbox) { Write-Log "Folder $monitorFolder not found." "ERROR"; return }
    
    $doneFolder = Ensure-OutlookFolder -Namespace $namespace -FolderPath $processedFolder
    
    $items = $inbox.Items
    $totalItems = $items.Count
    $pendingOrders = [System.Collections.Generic.List[object]]::new()
    # First pass: count unread orders
    # NOTE: foreach on Outlook COM collections silently yields nothing in Start-Job (MTA context).
    # Use indexed .Item($i) access instead - it marshals correctly across apartment boundaries.
    for ($i = 1; $i -le $totalItems; $i++) {
        try {
            $item = $items.Item($i)
            if ($item.UnRead -and $item.Subject -match "order") {
                $entryId = ""
                $storeId = ""
                try { $entryId = [string]$item.EntryID } catch { }
                try { $storeId = [string]$item.Parent.StoreID } catch { }
                $pendingOrders.Add([pscustomobject]@{
                    EntryID = $entryId
                    StoreID = $storeId
                })
            }
        } catch { }
    }
    $unreadOrders = $pendingOrders.Count
    Write-Log "Folder has $totalItems items, $unreadOrders unread order(s)" "INFO"
    if ($unreadOrders -gt 0) {
        Write-EmailProgress -Step "found" -Detail "Found $unreadOrders unread order(s) to process"
    }

    foreach ($pending in $pendingOrders) {
        $item = $null
        $entryId = ""
        $storeId = ""
        try { $entryId = [string]$pending.EntryID } catch { }
        try { $storeId = [string]$pending.StoreID } catch { }

        if (-not [string]::IsNullOrWhiteSpace($entryId)) {
            try {
                if (-not [string]::IsNullOrWhiteSpace($storeId)) {
                    $item = $namespace.GetItemFromID($entryId, $storeId)
                } else {
                    $item = $namespace.GetItemFromID($entryId)
                }
            } catch { }
        }

        if ($null -eq $item) {
            Write-Log "Skipping queued unread order email that no longer exists (EntryID=$entryId)." "WARN"
            continue
        }

            $success = $false
            if ([string]::IsNullOrWhiteSpace($entryId)) { try { $entryId = [string]$item.EntryID } catch { } }
            if ([string]::IsNullOrWhiteSpace($storeId)) { try { $storeId = [string]$item.Parent.StoreID } catch { } }

            if (-not [string]::IsNullOrWhiteSpace($entryId)) {
                if ($script:processedEntryIds.ContainsKey($entryId)) {
                    Write-Log "Skipping duplicate in-session email (already processed): EntryID=$entryId" "WARN"
                    continue
                }
                $script:processedEntryIds[$entryId] = $true
            }

            $candidateDocx = $null
            $bodyText = ""
            $subjectText = ""
            try { $bodyText = [string]$item.Body } catch { $bodyText = "" }
            try { $subjectText = [string]$item.Subject } catch { $subjectText = "" }
            # PDM notifications are HTML emails - plain text Body may lose the .docx path
            # Fall back to HTMLBody (strip tags) if Body has no .docx reference
            if ($bodyText -notmatch '(?i)\.docx\b') {
                try {
                    $htmlBody = [string]$item.HTMLBody
                    if ($htmlBody -match '(?i)\.docx\b') {
                        $bodyText = $htmlBody -replace '<[^>]+>', ' ' -replace '&nbsp;', ' ' -replace '&amp;', '&' -replace '&#92;', '\' -replace '\s+', ' '
                        Write-Log "  Using HTMLBody (plain text Body lacked .docx reference)" "INFO"
                    }
                } catch { }
            }
            try {
                $candidateDocx = Get-DocxPathFromEmail -EmailBody $bodyText -EmailSubject $subjectText
            } catch { $candidateDocx = $null }
            $mentionsDocx = ($bodyText -match '(?i)\.docx\b') -or ($subjectText -match '(?i)\.docx\b') -or ($bodyText -match '(?i)\bF80A?\.docx\b') -or ($subjectText -match '(?i)\bF80A?\.docx\b')
            if ($mentionsDocx) {
                Write-Log "  Email appears to reference a docx/F80 file." "INFO"
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$candidateDocx)) {
                Write-Log "  Extracted docx path candidate: $candidateDocx" "INFO"
            } else {
                Write-Log "  No docx path extracted from email body/HTML." "INFO"
            }
            if ([string]::IsNullOrWhiteSpace([string]$candidateDocx) -and -not $mentionsDocx) {
                Write-Log "Skipping unread 'Order' email without F80 .docx reference (likely non-order state-change notification)." "INFO"
                Write-EmailProgress -Step "complete" -Order $item.Subject -Detail "Skipped non-F80 notification"
                continue
            }

            $markedEarly = $false
            for ($attempt = 1; $attempt -le 2 -and -not $markedEarly; $attempt++) {
                try {
                    $item.UnRead = $false
                    try { $item.Save() } catch { }
                    $markedEarly = $true
                    Write-Log "Marked order email as read before processing (attempt $attempt)." "INFO"
                } catch {
                    Start-Sleep -Milliseconds 300
                }
            }

            Push-HubNotification -Title "Order Found" -Message "Processing: $($item.Subject)"
            Write-EmailProgress -Step "processing" -Order $item.Subject -Detail "Processing email..."
            try {
                $success = Process-Order -MailItem $item -OutlookApp $outlook -DocxPath $candidateDocx
            } catch {
                Write-Log "Process-Order exception: $($_.Exception.Message)" "ERROR"
                Write-EmailProgress -Step "failed" -Order $item.Subject -Detail "Error: $($_.Exception.Message)"
                $success = $false
            } finally {
                # Always attempt to mark as read, even if processing hit transient RPC issues.
                $markReadOk = $false
                for ($attempt = 1; $attempt -le 3 -and -not $markReadOk; $attempt++) {
                    try {
                        $targetItem = $item
                        if (($null -eq $targetItem -or -not $targetItem) -and -not [string]::IsNullOrWhiteSpace($entryId)) {
                            if (-not [string]::IsNullOrWhiteSpace($storeId)) {
                                $targetItem = $namespace.GetItemFromID($entryId, $storeId)
                            } else {
                                $targetItem = $namespace.GetItemFromID($entryId)
                            }
                        }
                        if ($targetItem) {
                            $targetItem.UnRead = $false
                            try { $targetItem.Save() } catch { }
                            $markReadOk = $true
                        }
                    } catch {
                        Start-Sleep -Milliseconds 500
                    }
                }
                if (-not $markReadOk) {
                    Write-Log "WARNING: could not mark order email as read after processing (EntryID=$entryId)" "WARN"
                }
            }

            # Write result to email progress
            if ($success) {
                Write-EmailProgress -Step "complete" -Order $item.Subject -Detail "Order processed successfully"
            } elseif ($script:lastProcessOrderFailReason -eq "FileNotFound") {
                Write-EmailProgress -Step "pending_retry" -Order $item.Subject -Detail "F80 not in vault yet - will retry next check"
            } else {
                Write-EmailProgress -Step "failed" -Order $item.Subject -Detail "Processing failed - check logs"
            }

            # If F80 wasn't in the vault yet, restore unread so the next auto-check retries
            if (-not $success -and $script:lastProcessOrderFailReason -eq "FileNotFound") {
                try {
                    $retryItem = $item
                    if (($null -eq $retryItem -or -not $retryItem) -and -not [string]::IsNullOrWhiteSpace($entryId)) {
                        $retryItem = if ($storeId) { $namespace.GetItemFromID($entryId, $storeId) } else { $namespace.GetItemFromID($entryId) }
                    }
                    if ($retryItem) {
                        $retryItem.UnRead = $true
                        try { $retryItem.Save() } catch { }
                        # Remove from dedup so the retry is not blocked in this session
                        if (-not [string]::IsNullOrWhiteSpace($entryId)) {
                            $script:processedEntryIds.Remove($entryId) | Out-Null
                        }
                        Write-Log "F80 not found in vault - email restored to UNREAD, will retry next check" "WARN"
                        Push-HubNotification -Title "Order Pending Retry" -Message "F80 not in vault yet: $($item.Subject)"
                    }
                } catch { }
            }

            if (-not $TestMode -and $success -and $doneFolder) {
                try {
                    $item.Move($doneFolder) | Out-Null
                } catch {
                    Write-Log "WARNING: could not move processed email: $($_.Exception.Message)" "WARN"
                }
            }
    }

    # --- Write idle progress when no orders found ---
    if ($unreadOrders -eq 0) {
        Write-EmailProgress -Step "idle" -Detail "Inbox checked - no new orders"
    }

    # --- Write "no orders" summary so dashboard shows last check was clean ---
    if ($unreadOrders -eq 0) {
        try {
            @{
                Timestamp       = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                Status          = "No orders found"
                OrderName       = ""
                JobNumber       = ""
                Client          = ""
                PartsExtracted  = 0
                PdfsCollected   = 0
                DxfsCollected   = 0
                TransmittalSent = $false
                DraftCreated    = $false
                TestMode        = [bool]$TestMode
                DispatchMode    = $DispatchMode
                NotFoundParts   = @()
            } | ConvertTo-Json | Set-Content -Path (Join-Path $indexFolder "last_email_summary.json") -Encoding UTF8 -Force
        } catch { }
    }
}

# --- Main ---
Write-Log "Email Order Monitor v2.0 started" "INFO"
$mutexSuffix = if ($TestMode) { "TEST" } else { "PROD" }
try {
    if (-not [string]::IsNullOrWhiteSpace($Config)) {
        $cfgLeaf = [System.IO.Path]::GetFileName($Config)
        if (-not [string]::IsNullOrWhiteSpace($cfgLeaf)) {
            $mutexSuffix = ($cfgLeaf -replace '[^A-Za-z0-9]', '_').ToUpperInvariant()
        }
    }
} catch { }
$mutexName = "NMT_EmailOrderMonitor_Mutex_${mutexSuffix}"
Write-Log "Instance mutex: $mutexName" "INFO"
$mutex = New-Object System.Threading.Mutex($false, $mutexName)
$mutexAcquired = $false
try {
    $mutexAcquired = $mutex.WaitOne(0)
} catch [System.Threading.AbandonedMutexException] {
    # Previous instance was killed/crashed without releasing the mutex.
    # WaitOne() still acquires it on this path - safe to continue.
    Write-Log "Previous instance exited abnormally (mutex abandoned) - proceeding with fresh start" "WARN"
    $mutexAcquired = $true
}
if (-not $mutexAcquired) { Write-Log "Another instance is already running. Exiting." "WARN"; exit }

if ($Watch) {
    while ($true) {
        try { Run-OrderCheck } catch { Write-Log "Error: $($_.Exception.Message)" "ERROR" }
        Start-Sleep -Seconds $PollInterval
    }
} else {
    try {
        Run-OrderCheck
    } catch {
        $msg = $_.Exception.Message
        Write-Log "Fatal monitor error: $msg" "ERROR"
        try {
            Write-EmailProgress -Step "failed" -Detail ("Monitor failed: " + $msg)
        } catch { }
        exit 1
    }
}

