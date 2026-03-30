# ==============================================================================
#  EpicorOrderMonitor.ps1 - Nordic Minesteel Technologies
#  Monitors Outlook for Epicor Sales Order Acknowledgement PDFs
#  Pipeline: Email -> PDF -> Text Extraction -> Part Validation ->
#            Epicor Check -> Rev Validation -> Collect Drawings -> Transmittal
# ==============================================================================

param(
    [switch]$TestMode,                # Safe mode: draft only, no email move
    [string]$Config   = "config.json", # Path to config file
    [string]$PdfPath  = "",            # Direct PDF path (bypass email, for testing)
    [switch]$RunTest                   # Run against built-in test orders (no email/PDF needed)
)

# ==============================================================================
# Logging & Progress
# ==============================================================================

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $prefix = switch ($Level) { "ERROR" { "[!]" } "WARN" { "[~]" } "SUCCESS" { "[+]" } default { "[ ]" } }
    $entry  = "$ts $prefix $Message"
    if ($script:logFile) { Add-Content -Path $script:logFile -Value $entry -ErrorAction SilentlyContinue }
    $color  = switch ($Level) { "ERROR" { "Red" } "WARN" { "Yellow" } "SUCCESS" { "Green" } default { "Gray" } }
    Write-Host $entry -ForegroundColor $color
}

function Write-EmailProgress {
    param([string]$Step, [string]$Order = "", [string]$Detail = "")
    try {
        @{
            Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            Step      = $Step
            Order     = $Order
            Detail    = $Detail
        } | ConvertTo-Json | Set-Content -Path (Join-Path $indexFolder "email_progress.json") -Encoding UTF8 -Force
    } catch { }
}

function Push-HubNotification {
    param([string]$Title, [string]$Message, [string]$FolderPath = "")
    try {
        $stamp      = (Get-Date).ToString("yyyyMMdd_HHmmss_fff")
        $notifyFile = Join-Path $indexFolder "notify_${stamp}.json"
        @{ Title = $Title; Message = $Message; FolderPath = $FolderPath } |
            ConvertTo-Json | Set-Content $notifyFile -Encoding UTF8 -Force
    } catch { }
}

# ==============================================================================
# Configuration
# ==============================================================================

$scriptDir  = Split-Path $PSCommandPath -Parent
$configPath = if ([System.IO.Path]::IsPathRooted($Config)) { $Config } else { Join-Path $scriptDir $Config }
$cfg        = @{}
if (Test-Path $configPath) {
    try { $cfg = Get-Content $configPath -Raw | ConvertFrom-Json } catch {
        Write-Host "WARN: Could not load config from $configPath" -ForegroundColor Yellow
    }
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

$indexFolder    = if ($cfg.indexFolder)       { $cfg.indexFolder }       else { "C:\Users\dlebel\Documents\PDFIndex" }
$outputRoot     = if ($cfg.outputFolder)      { $cfg.outputFolder }      else { "C:\Users\dlebel\Documents\AssemblyPDFs" }
$logFolder      = if ($cfg.logFolder)         { $cfg.logFolder }         else { (Join-Path $outputRoot "MacroLogs") }
$crawlRoots     = if ($cfg.crawlRoots)       { @($cfg.crawlRoots) }     else { @("C:\NMT_PDM", "J:\Epicor") }
$tesseractExe   = if ($cfg.tesseractPath)     { $cfg.tesseractPath }     else { "C:\Program Files\Tesseract-OCR\tesseract.exe" }
$ghostscriptExe = if ($cfg.ghostscriptPath)   { $cfg.ghostscriptPath }   else { "C:\Program Files\gs\gs9.56.1\bin\gswin64c.exe" }
$pdftotextExe   = if ($cfg.pdftotextPath)     { $cfg.pdftotextPath }     else { "C:\Program Files\poppler\bin\pdftotext.exe" }
# Fallback: pdftotext bundled with Git for Windows
if (-not (Test-Path $pdftotextExe)) {
    $gitPdftotext = "C:\Program Files\Git\mingw64\bin\pdftotext.exe"
    if (Test-Path $gitPdftotext) { $pdftotextExe = $gitPdftotext }
}

# Epicor API config
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
$script:epicorHeaders = $null   # built lazily on first use

$soCfg              = if ($cfg.salesOrderMonitor) { $cfg.salesOrderMonitor } elseif ($cfg.emailMonitor) { $cfg.emailMonitor } else { @{} }
$emailInboxFolder   = if ($soCfg.inboxFolder)               { $soCfg.inboxFolder }               else { "Inbox" }
$processedFolderName = if ($soCfg.processedFolder)          { $soCfg.processedFolder }           else { "Inbox\SalesOrders_Processed" }
$subjectKeywords    = if ($soCfg.subjectKeywords)           { @($soCfg.subjectKeywords) }        else { @("Sales Order", "Order Acknowledgement", "SO Acknowledgement") }
$collectMode        = if ($soCfg.collectMode)               { $soCfg.collectMode }               else { "BOTH" }
$transmittalTo      = if ($TestMode) {
    if ($soCfg.transmittalToTest) { $soCfg.transmittalToTest } else { "dlebel@nmtech.com" }
} else {
    if ($soCfg.transmittalToProd) { $soCfg.transmittalToProd } else { "dlebel@nmtech.com" }
}

# Logging setup
$monitorLogDir = Join-Path $logFolder "SalesOrderMonitor"
if (-not (Test-Path $monitorLogDir)) { New-Item -ItemType Directory -Path $monitorLogDir -Force | Out-Null }
$script:logFile = Join-Path $monitorLogDir "so_monitor_$(Get-Date -Format 'yyyy-MM-dd').log"

# Processed email cache (prevents double-processing)
$processedCachePath = Join-Path $indexFolder "so_processed_emails.json"

# SimpleCollector path (parent directory)
$simpleCollectorPath = Join-Path (Split-Path $scriptDir -Parent) "SimpleCollector.ps1"
if (-not (Test-Path $simpleCollectorPath)) { $simpleCollectorPath = Join-Path $scriptDir "SimpleCollector.ps1" }

# ==============================================================================
# PDF Text Extraction - five-strategy pipeline
# ==============================================================================

# Helper: decode a PDF content stream and extract visible text (Tj / TJ operators)
function Get-TextFromPdfStream {
    param([byte[]]$StreamBytes, [bool]$IsFlate)
    try {
        $raw = $StreamBytes
        if ($IsFlate -and $raw.Length -gt 2) {
            # zlib stream: skip 2-byte CMF/FLG header before DeflateStream
            $ms      = [System.IO.MemoryStream]::new($raw, 2, $raw.Length - 2)
            $deflate = [System.IO.Compression.DeflateStream]::new($ms, [System.IO.Compression.CompressionMode]::Decompress)
            $out     = [System.IO.MemoryStream]::new()
            $deflate.CopyTo($out)
            $deflate.Dispose(); $ms.Dispose()
            $raw = $out.ToArray()
        }
        $cs  = [System.Text.Encoding]::GetEncoding('iso-8859-1').GetString($raw)
        $sb  = [System.Text.StringBuilder]::new()
        $lastCharWasNL = $false

        # Match text-showing operators AND newline-producing positioning operators
        # PDF Td uses TWO numbers: tx ty Td  (not one)
        $ops = [regex]::Matches($cs,
            '\(([^\)\\]*(?:\\.[^\)\\]*)*)\)\s*(?:Tj|TJ)' +
            '|\[([^\]]*)\]\s*TJ' +
            '|[-\d.]+\s+[-\d.]+\s+Td\b' +
            '|\bTD\b|\bT\*\b|\bET\b')

        foreach ($op in $ops) {
            # Positioning/newline operator (no capture groups matched)
            if (-not $op.Groups[1].Success -and -not $op.Groups[2].Success) {
                if ($sb.Length -gt 0) { $sb.Append("`n") | Out-Null }
                $lastCharWasNL = $true
                continue
            }

            $chunks = @()
            if ($op.Groups[1].Success -and $op.Groups[1].Value) {
                $chunks = @($op.Groups[1].Value)
            } elseif ($op.Groups[2].Success -and $op.Groups[2].Value) {
                $chunks = [regex]::Matches($op.Groups[2].Value, '\(([^\)\\]*(?:\\.[^\)\\]*)*)\)') |
                          ForEach-Object { $_.Groups[1].Value }
            }
            foreach ($chunk in $chunks) {
                # Decode PDF string escapes
                $t = $chunk -replace '\\n',' ' -replace '\\r',' ' -replace '\\t',' '
                $t = [regex]::Replace($t, '\\([0-7]{3})', { [char][Convert]::ToInt32($args[0].Groups[1].Value, 8) })
                $t = $t -replace '\\(.)', '$1'
                if ($t.Trim()) {
                    # Add a space between consecutive text objects so cells don't merge
                    if ($sb.Length -gt 0 -and -not $lastCharWasNL) {
                        $last = $sb.ToString()[-1]
                        if ($last -ne ' ' -and $last -ne "`n" -and $t[0] -ne ' ') {
                            $sb.Append(' ') | Out-Null
                        }
                    }
                    $sb.Append($t) | Out-Null
                    $lastCharWasNL = $false
                }
            }
        }
        return $sb.ToString()
    } catch { return "" }
}

# Post-processor: Crystal Reports concatenates all text; insert newlines before
# known Epicor Sales Order Acknowledgement keywords to restore line structure.
function Format-EpicorPdfText {
    param([string]$Text)
    # Normalize line endings first
    $t = $Text -replace "`r`n", "`n" -replace "`r", "`n"

    # Insert newline BEFORE these keywords when they don't already start a line
    $breaks = @(
        'Sales Order\s*[:#]',
        'Sold To\s*:',
        'Ship To\s*:',
        'Bill To\s*:',
        'Our Part\s*:',
        'Line\s*Part\s*Number',
        'Line Total\s*:',
        'Order Total\s*:',
        'Total Tax',
        'Subtotal\s*:',
        'JOB CHARGES',
        'Rel Need By',
        'Rel Date',
        'Ship By',
        'Payment Terms',
        'Sales Representative',
        'Acknowledgement\b'
    )
    foreach ($kw in $breaks) {
        # Only add \n if not already at start of line
        $t = [regex]::Replace($t, "(?<!\n)($kw)", "`n`$1")
    }
    # Fix specific Crystal Reports merge: "LinePart" -> "Line Part"
    $t = $t -replace '\bLinePart\b', 'Line Part'
    # Collapse multiple blank lines
    $t = [regex]::Replace($t, '\n{3,}', "`n`n")
    return $t.Trim()
}

# Strategy A: native .NET PDF stream parser - works for digital Epicor PDFs, no tools needed
function Get-NativePdfText {
    param([string]$PdfPath)
    try {
        $enc   = [System.Text.Encoding]::GetEncoding('iso-8859-1')
        $bytes = [System.IO.File]::ReadAllBytes($PdfPath)
        $pdf   = $enc.GetString($bytes)
        $sb    = [System.Text.StringBuilder]::new()
        $pos   = 0

        while ($true) {
            $si = $pdf.IndexOf('stream', $pos)
            if ($si -lt 0) { break }
            $ei = $pdf.IndexOf('endstream', $si + 6)
            if ($ei -lt 0) { break }

            # Data starts after "stream" + optional CRLF
            $di = $si + 6
            while ($di -lt $pdf.Length -and ($pdf[$di] -eq "`r" -or $pdf[$di] -eq "`n")) { $di++ }

            # Dictionary immediately before this stream
            $ds = $pdf.LastIndexOf('<<', $si)
            $dict = if ($ds -ge 0) { $pdf.Substring($ds, [Math]::Min($si - $ds, 600)) } else { "" }

            $isImage = $dict -match '/Subtype\s*/Image'
            $isFlate  = $dict -match '/FlateDecode'

            if (-not $isImage) {
                $rawBytes = $enc.GetBytes($pdf.Substring($di, $ei - $di))
                $text = Get-TextFromPdfStream -StreamBytes $rawBytes -IsFlate $isFlate
                if ($text.Trim()) { $sb.Append($text).Append("`n") | Out-Null }
            }
            $pos = $ei + 9
        }
        $raw = $sb.ToString().Trim()
        return Format-EpicorPdfText -Text $raw
    } catch { return "" }
}

function Get-PdfText {
    param([string]$PdfPath)
    Write-Log "Extracting text from: $([System.IO.Path]::GetFileName($PdfPath))"

    # Strategy 1: Native .NET PDF stream parser - instant, no tools, works for all digital Epicor PDFs
    Write-Log "  [1/4] Native PDF parser (no tools needed)..." "INFO"
    $nativeText = Get-NativePdfText -PdfPath $PdfPath
    if (-not [string]::IsNullOrWhiteSpace($nativeText) -and $nativeText.Length -gt 50) {
        Write-Log "  Native PDF parser: $($nativeText.Length) chars extracted" "SUCCESS"
        return $nativeText
    }
    Write-Log "  Native parser: $($nativeText.Length) chars - insufficient, trying next" "WARN"

    # Strategy 2: pdftotext.exe (poppler) - excellent layout-preserving extraction
    if (Test-Path $pdftotextExe) {
        Write-Log "  [2/4] pdftotext.exe..." "INFO"
        try {
            $outFile = [System.IO.Path]::GetTempFileName() + ".txt"
            & $pdftotextExe -layout $PdfPath $outFile 2>$null
            if (Test-Path $outFile) {
                $text = Get-Content $outFile -Raw
                Remove-Item $outFile -Force -ErrorAction SilentlyContinue
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    Write-Log "  pdftotext: $($text.Length) chars" "SUCCESS"
                    return $text
                }
            }
        } catch { Write-Log "  pdftotext error: $($_.Exception.Message)" "WARN" }
    } else {
        Write-Log "  [2/4] pdftotext.exe not found at $pdftotextExe - skipping" "INFO"
    }

    # Strategy 3: Ghostscript + Tesseract (for scanned/image-based PDFs)
    if ((Test-Path $ghostscriptExe) -and (Test-Path $tesseractExe)) {
        Write-Log "  [3/4] Ghostscript + Tesseract OCR..." "INFO"
        $tmpDir = $null
        try {
            $tmpDir = Join-Path $env:TEMP "so_ocr_$(Get-Random)"
            New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
            Write-Log "    Rendering PDF pages to PNG..." "INFO"
            & $ghostscriptExe -dBATCH -dNOPAUSE -sDEVICE=png16m -r300 `
                "-sOutputFile=$tmpDir\page_%03d.png" $PdfPath 2>$null
            $pages   = @(Get-ChildItem $tmpDir -Filter "*.png" | Sort-Object Name)
            Write-Log "    OCR-ing $($pages.Count) page(s)..." "INFO"
            $allText = ""
            foreach ($page in $pages) {
                $base = $page.FullName -replace '\.png$', ''
                & $tesseractExe $page.FullName $base --psm 6 2>$null
                $txtFile = "$base.txt"
                if (Test-Path $txtFile) {
                    $allText += (Get-Content $txtFile -Raw) + "`n"
                    Remove-Item $txtFile -Force -ErrorAction SilentlyContinue
                }
            }
            if (-not [string]::IsNullOrWhiteSpace($allText)) {
                Write-Log "  GS+Tesseract: $($allText.Length) chars" "SUCCESS"
                return $allText
            }
        } catch { Write-Log "  GS+Tesseract error: $($_.Exception.Message)" "WARN" }
        finally {
            if ($tmpDir -and (Test-Path $tmpDir)) { Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
        }
    } else {
        Write-Log "  [3/4] GS/Tesseract not found - skipping OCR strategies" "INFO"
    }

    # Strategy 4: Tesseract directly on PDF
    if (Test-Path $tesseractExe) {
        Write-Log "  [4/4] Tesseract direct on PDF..." "INFO"
        try {
            $outBase = Join-Path $env:TEMP "so_tess_$(Get-Random)"
            & $tesseractExe $PdfPath $outBase --psm 6 2>$null
            $txtFile = "$outBase.txt"
            if (Test-Path $txtFile) {
                $text = Get-Content $txtFile -Raw
                Remove-Item $txtFile -Force -ErrorAction SilentlyContinue
                if (-not [string]::IsNullOrWhiteSpace($text)) {
                    Write-Log "  Tesseract: $($text.Length) chars" "SUCCESS"
                    return $text
                }
            }
        } catch { Write-Log "  Tesseract error: $($_.Exception.Message)" "WARN" }
    }

    Write-Log "  All extraction methods failed. PDF may be encrypted or image-only." "WARN"
    return ""
}

# ==============================================================================
# Part Extraction - Epicor Sales Order Acknowledgement format
# ==============================================================================

$script:PartPatterns = @(
    '\b\d{4,6}[-_][A-Z0-9]{1,8}(?:[-_][A-Z0-9]{1,8}){0,4}\b',  # NMT std: 4823-P4-34, 1202-WM-440
    '\b[A-Z]{2,5}-\d+[.,]\d{1,3}(?:[-x][A-Z0-9.]+){1,5}\b',     # Hardware: FHCS-0.625-11x2.000-ZP
    '\b\d{5}-\d{2}-[A-Z]\d{2,3}(?:-[A-Z])?\b',                   # Job part: 17141-10-P67
    '\b\d{8,10}\b',                                               # Long customer part: 40617527
    '\b[A-Z]{1,4}\d{1,5}[A-Z]{0,2}\b',                           # Short alpha-num: BRG03, W010A
    '\b[A-Z]{1,6}\d{1,4}(?:-[A-Z0-9]{1,6}){1,8}\b',             # Alpha-dash: BF12C11-P
    '\b\d{5}[A-Z]\d{3,4}\b',                                      # McMaster: 90107A030
    '\b[A-Z]{2,4}\d{2,6}(?:-[A-Z]{1,3})?\b',                     # Prefix: PB23056-FS
    '\b\d{2}-\d{2}-\d{2,3}(?:-[A-Z0-9]{1,4}){0,3}\b'            # Short assy: 02-05-00-C30
)


function Get-SalesOrderParts {
    param([string]$Text)

    # Epicor Crystal Reports PDFs produce vertical text: each cell on its own line.
    # We use a state-machine index walk (not foreach) so we can lookahead/skip lines.
    $lines       = @(($Text -split "`n") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
    $parts       = [System.Collections.Generic.List[object]]::new()
    $orderNumber = "UNKNOWN"
    $clientName  = ""
    $inTable     = $false
    $lastLineNum = 0
    $cur         = $null    # current line-item hashtable

    # Column-header labels to skip when inside a line item
    $colHeaders = @('Rev','Order Qty','Unit Price','Ext. Price','Part Number/Description',
                    'Part Number','Rel','Need By','Quantity','Ship By','Rel Need By','Rel Date',
                    'EA','Description','Line')

    # Noise patterns:
    # - discard-only rows that should never become line items
    # - note-only rows that belong to the current item but should not discard it
    $discardKw = '(?i)^(JOB\s+CHARGES|JOB\s*/\s*ORDER\s+ADDITIONAL|RUBBER|TAPERED\s+ROLLER|HINGE\s+ASSEMBLY)'
    $skipOnlyKw = '(?i)^(Hardware\s+for|Supplied\s+w)'

    # Date pattern
    $datePattern = '^\d{1,2}[-/][A-Za-z]{3}[-/]\d{2,4}$|^\d{1,2}/\d{1,2}/\d{4}$'

    # Helper: find best part match in a single token
    $findPart = {
        param($tok)
        $best = $null
        foreach ($pat in $script:PartPatterns) {
            if ($tok -match $pat) {
                $m = $matches[0]
                if (-not $best -or $m.Length -gt $best.Length) { $best = $m }
            }
        }
        return $best
    }

    # Inline row variant from pdftotext/native extraction:
    # "1 40393469 12.00EA ..." or "3 BAC014PB-A 1 24.00EA ..."
    $parseCompactRow = {
        param([string]$text)
        if ([string]::IsNullOrWhiteSpace($text)) { return $null }
        $t = [regex]::Replace($text.Trim(), '\s+', ' ')
        $m = [regex]::Match(
            $t,
            '^(?<line>\d{1,3})\s+(?<part>[A-Z0-9][A-Z0-9._-]{3,})(?:\s+(?<rev>(?!\d+(?:[.,]\d+)?EA\b)[A-Z0-9]+))?\s+(?<qty>\d+(?:[.,]\d+)?)EA\b',
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
        if (-not $m.Success) { return $null }
        $part = [string]$m.Groups['part'].Value.Trim().ToUpperInvariant()
        if ($part -match '^(?i)(TOTAL|SUBTOTAL|TAX|FREIGHT|CHARGES|ORDERACK)$') { return $null }
        return @{
            Line = [int]$m.Groups['line'].Value
            Part = $part
            Rev  = if ($m.Groups['rev'].Success) { [string]$m.Groups['rev'].Value.Trim().ToUpperInvariant() } else { "NA" }
            Qty  = [string]$m.Groups['qty'].Value.Trim()
        }
    }

    $findNearbyCustomerPart = {
        param([int]$startIndex, [string]$internalPart)
        $internalNorm = ([string]$internalPart).Trim().ToUpperInvariant()
        for ($j = $startIndex; $j -ge 0 -and $j -ge ($startIndex - 6); $j--) {
            $candLine = ([string]$lines[$j]).Trim()
            if ([string]::IsNullOrWhiteSpace($candLine)) { continue }
            if ($candLine -match '^(?i)(Our\s+Part|Rel\s+Need|Ship\s+By|Quantity|Line|Rev|Order\s+Qty|Unit\s+Price|Ext\.?\s*Price)$') { continue }
            if ($candLine -match '^\d+$') { continue }
            if ($candLine -match '^\d+(?:[.,]\d+)?(?:EA)?$') { continue }
            if ($candLine -match $datePattern) { continue }
            $found = & $findPart $candLine
            if ([string]::IsNullOrWhiteSpace($found)) { continue }
            $found = $found.TrimEnd('-').ToUpperInvariant()
            if ($found -eq $internalNorm) { continue }
            return $found
        }
        return ""
    }

    # Description patterns - lines that look like part descriptions (ALL CAPS text, not a part number)
    $descPattern = '^[A-Z][A-Z0-9 ,./&''\-]{2,60}$'

    # Helper: save current item to list
    $saveItem = {
        if ($null -ne $cur -and $cur.Part) {
            $parts.Add([PSCustomObject]@{
                Order = $orderNumber
                Line  = $cur.Line
                Part  = $cur.Part
                CustPart = $cur.CustPart
                Rev   = $cur.Rev
                Qty   = $cur.Qty
                Desc  = $cur.Desc
            })
            $logPart = if ($cur.CustPart -and $cur.CustPart -ne $cur.Part) { "$($cur.CustPart) -> $($cur.Part)" } else { $cur.Part }
            Write-Log "  Line $($cur.Line) : $logPart  Rev=$($cur.Rev)  Qty=$($cur.Qty)  Desc=$($cur.Desc)" "SUCCESS"
        }
    }

    # Pre-scan for order number and client name (they appear before the table)
    foreach ($ln in $lines) {
        if ($ln -match '(?i)Sales\s+Order.*?(\d{4,6})') { $orderNumber = $matches[1] }
        if (-not $clientName -and $ln -match '(?i)(Freeport|McMoran|Barrick|Anglo|Newmont|Teck|BHP|Rio\s+Tinto|Vale|Kinross|Agnico)') {
            $clientName = $ln
        }
    }

    $i = 0
    while ($i -lt $lines.Count) {
        $line = $lines[$i]

        # ---- Table start detection ----
        if (-not $inTable) {
            # Vertical header: "Line" on its own, then "Part Number" within next 4 lines
            if ($line -eq "Line") {
                $found = $false
                for ($j = $i+1; $j -lt [Math]::Min($i+5, $lines.Count); $j++) {
                    if ($lines[$j] -match '(?i)Part\s*Number') { $inTable = $true; $i = $j + 1; $found = $true; break }
                }
                if ($found) { continue }
            }
            # Horizontal header variant: "Line Part Number/Description ..."
            if ($line -match '(?i)Line\s+Part\s*Number') { $inTable = $true; $i++; continue }
            $i++; continue
        }

        # ---- Inside table ----

        # Table end markers
        if ($line -match '(?i)(Line\s+Total|Order\s+Total|Total\s+Tax|Subtotal)') {
            & $saveItem; $cur = $null; $inTable = $false; $i++; continue
        }

        # Skip pure column header labels
        if ($colHeaders -contains $line) { $i++; continue }

        # Skip dates
        if ($line -match $datePattern) { $i++; continue }

        # Non-item note lines should be ignored without discarding the current valid part.
        if ($line -match $skipOnlyKw) {
            $i++; continue
        }

        # Actual non-item rows should abandon the current parse state.
        if ($line -match $discardKw) {
            $cur = $null; $i++; continue
        }

        # Compact inline row with line / part / optional rev / qty on one line
        $compactRow = & $parseCompactRow $line
        if ($null -ne $compactRow) {
            & $saveItem
            $cur = @{
                Line     = $compactRow.Line
                Part     = $compactRow.Part
                Rev      = $compactRow.Rev
                Qty      = $compactRow.Qty
                CustPart = $compactRow.Part
                Desc     = ""
            }
            $lastLineNum = [int]$compactRow.Line
            $i++; continue
        }

        # --- Standalone integer: new line item OR revision column value ---
        # In vertical Crystal Reports format the Rev column appears BEFORE Qty.
        # If we already have a customer part but no rev/qty yet, this integer is the rev.
        if ($line -match '^\d+$') {
            $n = [int]$line
            # Priority 1: if the current item still has no revision, a standalone integer
            # is usually the Rev column unless the next line clearly starts a new part row.
            $nextLine = if ($i + 1 -lt $lines.Count) { ([string]$lines[$i + 1]).Trim() } else { "" }
            $nextPart = if (-not [string]::IsNullOrWhiteSpace($nextLine)) { & $findPart $nextLine } else { "" }
            $nextStartsNewRow = (-not [string]::IsNullOrWhiteSpace($nextPart)) -and ($nextPart.ToUpperInvariant() -ne ([string]$cur.Part).Trim().ToUpperInvariant())
            $nextLooksLikeReleaseRow = (-not [string]::IsNullOrWhiteSpace($nextLine)) -and ($nextLine -match $datePattern)
            if ($cur -and ($cur.Part -or $cur.CustPart) -and $cur.Rev -eq "NA" -and -not $nextStartsNewRow) {
                if ($nextLooksLikeReleaseRow) {
                    $i++; continue
                }
                $cur.Rev = $line
                $i++; continue
            }
            # Priority 2: start a new line item (integer greater than last line seen)
            if ($n -gt $lastLineNum -and $n -le 999) {
                & $saveItem
                $cur         = @{ Line = $n; Part = ""; Rev = "NA"; Qty = "0"; CustPart = ""; Desc = "" }
                $lastLineNum = $n
                $i++; continue
            }
            # Otherwise it's a release number, sub-line, or other numeric noise - skip
            $i++; continue
        }

        # No current item yet - keep scanning
        if (-not $cur) { $i++; continue }

        # --- "Our Part:" label ---
        # Vertical: "Our Part:" on one line, part number on next, rev on the line after
        if ($line -match '(?i)^Our\s+Part\s*:?\s*$') {
            if ($i+1 -lt $lines.Count) {
                $nextLine = $lines[$i+1]
                if ($nextLine -match '^[A-Z0-9][-A-Z0-9._]{1,30}$') {
                    $cur.Part = $nextLine
                    if (-not $cur.CustPart) { $cur.CustPart = & $findNearbyCustomerPart ($i - 1) $cur.Part }
                    Write-Log "  Cross-ref Our Part -> $nextLine" "SUCCESS"
                    if ($i+2 -lt $lines.Count -and $lines[$i+2] -match '^\d{1,4}$') {
                        $cur.Rev = $lines[$i+2]
                        $i += 3; continue
                    }
                    $i += 2; continue
                }
            }
        }
        # Inline: "Our Part: 1206-P404" (possibly with trailing rev token)
        elseif ($line -match '(?i)Our\s+Part\s*:\s*([A-Z0-9][-A-Z0-9._]{1,30})') {
            $cur.Part = $matches[1]
            if (-not $cur.CustPart) { $cur.CustPart = & $findNearbyCustomerPart ($i - 1) $cur.Part }
            Write-Log "  Cross-ref Our Part -> $($cur.Part)" "SUCCESS"
            # Check if rev trails on same line
            $rest = ($line -replace "(?i).*Our\s+Part\s*:\s*$([regex]::Escape($matches[1]))", '').Trim()
            if ($rest -match '^(\d{1,4})$') {
                $cur.Rev = $matches[1]; $i++; continue
            }
            # Or on next line
            if ($i+1 -lt $lines.Count -and $lines[$i+1] -match '^\d{1,4}$') {
                $cur.Rev = $lines[$i+1]; $i += 2; continue
            }
            $i++; continue
        }

        # --- Explicit "NA" rev token ---
        if ($line -eq "NA" -and $cur.Rev -eq "NA") {
            $cur.Rev = "NA"   # already NA; just consume so it doesn't match as a part
            $i++; continue
        }

        # --- Quantity: "3.00" with 2 decimal places, followed by "EA" on same or next line ---
        if ($cur.Qty -eq "0") {
            if ($line -match '^(\d+\.\d{1,2})$') {
                $qtyVal = $matches[1]  # save before next -match clobbers $matches
                if ($i+1 -lt $lines.Count -and $lines[$i+1] -match '^EA$') {
                    $cur.Qty = $qtyVal; $i += 2; continue
                }
            }
            if ($line -match '^(\d+(?:\.\d{1,2})?)\s*EA$') {
                $cur.Qty = $matches[1]; $i++; continue
            }
        }

        # --- Customer part number (first part-like token, before "Our Part:" overrides) ---
        if (-not $cur.CustPart) {
            $found = & $findPart $line
            if ($found) {
                # Trim trailing dash (e.g. "FHCS-0.625-11x2.250-ZP-" -> "FHCS-0.625-11x2.250-ZP")
                $found = $found.TrimEnd('-')
                # Multi-line part: if this token ends with '-' before trimming, try concatenating next line
                if ($line -match '-$' -and ($i+1 -lt $lines.Count)) {
                    $nextTok = $lines[$i+1].Trim()
                    if ($nextTok -match '^[A-Z0-9][-A-Z0-9._]{1,20}$') {
                        $combined = & $findPart ($line.TrimEnd('-') + $nextTok)
                        if ($combined -and $combined.Length -gt $found.Length) {
                            $found = $combined
                            $i++   # consume the continuation line
                        }
                    }
                }
                $cur.CustPart = $found
                if (-not $cur.Part) { $cur.Part = $found }
            }
        }

        # --- Description line: ALL CAPS text after part has been identified ---
        # Must have a CustPart already, no desc yet, and line must look like a description
        if ($cur.CustPart -and -not $cur.Desc -and $line -match $descPattern) {
            # Exclude lines that are just "Rel Need By Ship By", "Quantity", or noise
            if ($line -notmatch '(?i)^(Rel\s+Need|Ship\s+By|Quantity|Canadian|Ext\.|Unit\s+Price|Order\s+Qty)') {
                $cur.Desc = $line
            }
        }

        $i++
    }

    & $saveItem

    foreach ($p in @($parts)) {
        if ($null -eq $p) { continue }
        $custPart = ([string]$p.CustPart).Trim()
        $internalPart = ([string]$p.Part).Trim()
        if (-not [string]::IsNullOrWhiteSpace($custPart) -or [string]::IsNullOrWhiteSpace($internalPart)) { continue }
        $lineNo = [string]$p.Line
        if ([string]::IsNullOrWhiteSpace($lineNo)) { continue }

        $rx = '(?is)\b' + [regex]::Escape($lineNo) + '\b\s+(?<cust>[A-Z0-9][A-Z0-9._-]{3,})\b.*?Our\s+Part\s*:\s*' + [regex]::Escape($internalPart) + '\b'
        $m = [regex]::Match($Text, $rx, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($m.Success) {
            $foundCust = ([string]$m.Groups['cust'].Value).Trim().TrimEnd('-').ToUpperInvariant()
            if (-not [string]::IsNullOrWhiteSpace($foundCust) -and $foundCust -ne $internalPart.ToUpperInvariant()) {
                $p.CustPart = $foundCust
            }
        }
    }

    return @{ Parts = $parts; OrderNumber = $orderNumber; Client = $clientName }
}

# ==============================================================================
# Epicor REST API
# ==============================================================================

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

# Returns @{ Exists=$true/false; Revisions=@(...); LatestRev="4"; Approved=$true; Description="..." }
# Uses PartSvc/GetByID which returns the full dataset including PartRev child rows.
# Note: the OData entity is "PartRevs" (plural) but the dataset table is "PartRev" (singular).
function Get-EpicorPartInfo {
    param([string]$PartNumber)
    if (-not $epicorEnabled) { return $null }
    try {
        $h    = Get-EpicorHeaders
        $pnEnc = [Uri]::EscapeDataString($PartNumber)
        $url  = "$epicorApiUrl/api/v1/Erp.BO.PartSvc/GetByID?partNum=$pnEnc"
        $resp = Invoke-RestMethod -Uri $url -Headers $h -Method Get -TimeoutSec 15 -ErrorAction Stop

        $partRows = $resp.returnObj.Part
        if (-not $partRows -or @($partRows).Count -eq 0) {
            return @{ Exists = $false; Revisions = @(); LatestRev = ""; Approved = $false; Description = "" }
        }

        $desc    = $partRows[0].PartDescription
        $revRows = @($resp.returnObj.PartRev)   # singular "PartRev" = dataset table name

        if ($revRows.Count -eq 0) {
            return @{ Exists = $true; Revisions = @(); LatestRev = ""; Approved = $false; Description = $desc }
        }

        # Pick latest approved revision; fall back to any revision if none approved
        $approved = @($revRows | Where-Object { $_.Approved -eq $true })
        $pool     = if ($approved.Count -gt 0) { $approved } else { $revRows }
        $latest   = $pool | Sort-Object {
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
        $code = $_.Exception.Response.StatusCode.value__
        if ($code -eq 401) { Write-Log "Epicor API: 401 Unauthorized - check username/password in config.json" "WARN" }
        elseif ($code -eq 403) { Write-Log "Epicor API: 403 Forbidden - REST API may require an API key" "WARN" }
        else { Write-Log "Epicor API error [$code] for ${PartNumber}: $($_.Exception.Message)" "WARN" }
        return $null   # null = API unreachable, falls back to drawing index only
    }
}

# ==============================================================================
# Epicor / PDF Index Validation
# ==============================================================================

$script:pdfIndex = $null
$script:dxfIndex = $null

function Initialize-PdfIndex {
    if ($script:pdfIndex) { return }   # already loaded
    $csvPath = Join-Path $indexFolder "pdf_index_clean.csv"
    if (Test-Path $csvPath) {
        try {
            $script:pdfIndex = @(Import-Csv $csvPath)
            Write-Log "PDF index loaded: $($script:pdfIndex.Count) entries" "INFO"
        } catch {
            Write-Log "Could not load PDF index: $($_.Exception.Message)" "WARN"
            $script:pdfIndex = @()
        }
    } else {
        Write-Log "PDF index not found at $csvPath" "WARN"
        $script:pdfIndex = @()
    }

    # Also load DXF index
    if (-not $script:dxfIndex) {
        $dxfCsvPath = Join-Path $indexFolder "dxf_index_clean.csv"
        if (Test-Path $dxfCsvPath) {
            try {
                $script:dxfIndex = @(Import-Csv $dxfCsvPath)
                Write-Log "DXF index loaded: $($script:dxfIndex.Count) entries" "INFO"
            } catch {
                Write-Log "Could not load DXF index: $($_.Exception.Message)" "WARN"
                $script:dxfIndex = @()
            }
        } else {
            Write-Log "DXF index not found at $dxfCsvPath" "WARN"
            $script:dxfIndex = @()
        }
    }
}

function Find-PartInIndex {
    param([string]$PartNumber)
    if (-not $script:pdfIndex) { return @() }
    $pn = $PartNumber.Trim().ToUpperInvariant()
    return @($script:pdfIndex | Where-Object { ([string]$_.BasePart).Trim().ToUpperInvariant() -eq $pn })
}

function Find-PartInDxfIndex {
    param([string]$PartNumber)
    if (-not $script:dxfIndex) { return @() }
    $pn = $PartNumber.Trim().ToUpperInvariant()
    return @($script:dxfIndex | Where-Object { ([string]$_.BasePart).Trim().ToUpperInvariant() -eq $pn })
}

function Normalize-Rev {
    param([string]$Rev)
    if ([string]::IsNullOrWhiteSpace($Rev)) { return "NA" }
    return ($Rev.Trim() -replace '^[Rr][Ee][Vv]', '').Trim()
}

function Compare-OrderRevToIndex {
    param([string]$OrderRev, [string]$IndexRev)
    $a = Normalize-Rev $OrderRev
    $b = Normalize-Rev $IndexRev
    if ($a -eq "NA" -or $b -eq "NA") { return "Unknown" }
    if ($a -ieq $b)                  { return "Match" }
    return "Mismatch"
}

# ==============================================================================
# Drawing Collection via SimpleCollector.ps1
# ==============================================================================

function Invoke-DrawingCollection {
    param([string[]]$PartNumbers, [string]$OutputFolder, [string]$Mode = "BOTH")

    if (-not (Test-Path $simpleCollectorPath)) {
        Write-Log "SimpleCollector not found at: $simpleCollectorPath" "WARN"
        return @{ PdfsCollected = 0; DxfsCollected = 0 }
    }
    if ($PartNumbers.Count -eq 0) {
        return @{ PdfsCollected = 0; DxfsCollected = 0 }
    }

    $bomFile = Join-Path $env:TEMP "so_bom_$(Get-Random).txt"
    $PartNumbers | Set-Content $bomFile -Encoding UTF8
    if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }

    try {
        & $simpleCollectorPath $bomFile $OutputFolder $Mode $configPath 2>&1 | ForEach-Object {
            Write-Log "  [Collector] $_" "INFO"
        }
        Write-Log "Drawing collection complete" "SUCCESS"
    } catch {
        Write-Log "SimpleCollector error: $($_.Exception.Message)" "WARN"
    } finally {
        Remove-Item $bomFile -Force -ErrorAction SilentlyContinue
    }

    $pdfs = @(Get-ChildItem $OutputFolder -Filter "*.pdf" -ErrorAction SilentlyContinue)
    $dxfs = @(Get-ChildItem $OutputFolder -Filter "*.dxf" -ErrorAction SilentlyContinue)
    return @{ PdfsCollected = $pdfs.Count; DxfsCollected = $dxfs.Count }
}

# ==============================================================================
# Outlook Transmittal Draft
# ==============================================================================

function Send-TransmittalDraft {
    param(
        [string]$OrderNumber,
        [string]$Client,
        [object[]]$Parts,
        [string[]]$FlaggedParts,
        [string[]]$MissingParts,
        [string]$OutputFolder,
        [object[]]$CollectedPdfs = @(),
        [object[]]$CollectedDxfs = @(),
        [string]$SourcePdfPath = ""
    )
    try {
        # --- Try to get existing Outlook instance first (like EmailOrderMonitor) ---
        $outlook = $null
        try {
            $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        } catch {
            try { $outlook = New-Object -ComObject Outlook.Application } catch {
                Write-Log "Failed to get Outlook COM app for transmittal: $($_.Exception.Message)" "ERROR"
                throw
            }
        }

        # --- Try to load F02 Document Transmittal template ---
        $f02Template = ""
        $templateCandidates = @(
            (Join-Path (Split-Path $scriptDir -Parent) "F02 Document Transmittal v4.0.oft"),
            (Join-Path $scriptDir "F02 Document Transmittal v4.0.oft"),
            "U:\30-Common\Forms\NMT Blank Forms\F02 Document Transmittal v4.0.oft"
        )
        foreach ($cand in $templateCandidates) {
            if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path $cand)) {
                $f02Template = $cand
                break
            }
        }

        $mail = $null
        if (-not [string]::IsNullOrWhiteSpace($f02Template) -and (Test-Path $f02Template)) {
            try {
                $mail = $outlook.CreateItemFromTemplate($f02Template)
                Write-Log "Loaded F02 Document Transmittal template from $f02Template" "INFO"
            } catch {
                Write-Log "Template load failed ($($_.Exception.Message)); using blank email." "WARN"
            }
        } else {
            Write-Log "F02 template not found; using blank email" "WARN"
        }
        if ($null -eq $mail) {
            $mail = $outlook.CreateItem(0)
        }

        $mail.To      = $transmittalTo

        # --- Transmittal numbering (same logic as EmailOrderMonitor) ---
        $effectiveJob = if ([string]::IsNullOrWhiteSpace($OrderNumber)) { "UNKNOWN" } else { $OrderNumber.Trim().ToUpperInvariant() }
        $transmittalNo = "T01"
        $nextTransmittalNum = 1
        try {
            if (-not $TestMode) {
                $historyFileForNo = Join-Path $indexFolder "transmittal_history.json"
                if (Test-Path $historyFileForNo) {
                    $parsedHist = $null
                    try { $parsedHist = (Get-Content -Path $historyFileForNo -Raw).Trim() | ConvertFrom-Json } catch {}
                    $hist = [System.Collections.Generic.List[object]]::new()
                    if ($null -ne $parsedHist) {
                        if ($parsedHist -is [System.Array]) {
                            foreach ($pi in $parsedHist) { $hist.Add($pi) }
                        } else { $hist.Add($parsedHist) }
                    }
                    $sameJob = @($hist | Where-Object { $_ -and ([string]$_.JobNumber).Trim().ToUpperInvariant() -eq $effectiveJob })
                    $maxExistingNo = 0
                    foreach ($hrow in $sameJob) {
                        $noText = ""
                        if ($hrow.PSObject.Properties['TransmittalNo']) { $noText = [string]$hrow.TransmittalNo }
                        if ([string]::IsNullOrWhiteSpace($noText)) { continue }
                        $mNo = [regex]::Match($noText.ToUpperInvariant(), '^T?(\d{1,3})$')
                        if ($mNo.Success) {
                            $n = [int]$mNo.Groups[1].Value
                            if ($n -gt $maxExistingNo) { $maxExistingNo = $n }
                        }
                    }
                    if ($maxExistingNo -gt 0) { $nextTransmittalNum = $maxExistingNo + 1 }
                    elseif ($sameJob.Count -eq 1) { $nextTransmittalNum = 2 }
                    else { $nextTransmittalNum = 1 }
                    if ($nextTransmittalNum -lt 1) { $nextTransmittalNum = 1 }
                }
            }
        } catch {
            Write-Log "Could not derive next transmittal number: $($_.Exception.Message)" "WARN"
        }
        $transmittalNo = ("T{0:D2}" -f $nextTransmittalNum)
        Write-Log "Transmittal number: $transmittalNo" "INFO"

        # Set subject in Document Transmittal format (matches F02 template style)
        $testSuffix = if ($TestMode) { " [TEST]" } else { "" }
        $mail.Subject = "Document Transmittal ($effectiveJob-$transmittalNo)$testSuffix"

        # --- Derive job description from source PDF path ---
        # e.g. "J:\Epicor\Orders\Spare Parts\27000-27199\27131 - Alamos Gold YD - ..." -> "Spare Parts"
        $jobCategory = "Spare Parts"
        if ($SourcePdfPath) {
            if ($SourcePdfPath -match '(?i)\\(Capital|Spare\s*Parts|Service|Warranty)\\') {
                $jobCategory = $matches[1]
            }
        }

        # --- Helper: format qty as integer when whole number ---
        function Format-Qty {
            param([string]$Q)
            if ([string]::IsNullOrWhiteSpace($Q) -or $Q -eq "0") { return "" }
            if ($Q -match '^\d+\.0+$') { return ($Q -replace '\.0+$', '') }
            return $Q
        }

        # --- Helper: title-case a description ---
        function Format-Desc {
            param([string]$D)
            if ([string]::IsNullOrWhiteSpace($D)) { return "" }
            # Title case: capitalize first letter of each word, lowercase rest
            $words = $D.Trim() -split '\s+'
            $result = $words | ForEach-Object {
                $w = $_.ToLower()
                # Keep short prepositions/conjunctions lowercase unless first word
                if ($w.Length -le 2 -and $w -match '^(x|of|to|or|in|on|at|by|an|a)$') { return $w }
                # Capitalize first letter
                $w.Substring(0,1).ToUpper() + $w.Substring(1)
            }
            return ($result -join ' ')
        }

        # --- Build drawing map from collected files ---
        $drawMap = @{}
        $extractPartRev = {
            param([string]$LeafName)
            $name = [System.IO.Path]::GetFileNameWithoutExtension($LeafName).ToUpperInvariant()
            $base = $name; $rev = ""
            if ($name -match '^(?<base>.+?)_REV(?<rev>[A-Z0-9]+)$') { $base = $Matches['base']; $rev = $Matches['rev'] }
            if ($name -match '^(?<base>.+?)-REV(?<rev>[A-Z0-9]+)$') { $base = $Matches['base']; if (-not $rev) { $rev = $Matches['rev'] } }
            return @{ Base = $base; Rev = $rev }
        }
        foreach ($f in $CollectedPdfs) {
            $pr = & $extractPartRev $f.Name
            if ([string]::IsNullOrWhiteSpace($pr.Base)) { continue }
            if (-not $drawMap.ContainsKey($pr.Base)) { $drawMap[$pr.Base] = @{ Part = $pr.Base; HasPdf = $false; HasDxf = $false; PdfRev = ""; DxfRev = ""; PdfPath = ""; DxfPath = "" } }
            $drawMap[$pr.Base].HasPdf = $true
            $drawMap[$pr.Base].PdfPath = $f.FullName
            if ($pr.Rev) { $drawMap[$pr.Base].PdfRev = $pr.Rev }
        }
        foreach ($f in $CollectedDxfs) {
            $pr = & $extractPartRev $f.Name
            if ([string]::IsNullOrWhiteSpace($pr.Base)) { continue }
            if (-not $drawMap.ContainsKey($pr.Base)) { $drawMap[$pr.Base] = @{ Part = $pr.Base; HasPdf = $false; HasDxf = $false; PdfRev = ""; DxfRev = ""; PdfPath = ""; DxfPath = "" } }
            $drawMap[$pr.Base].HasDxf = $true
            $drawMap[$pr.Base].DxfPath = $f.FullName
            if ($pr.Rev) { $drawMap[$pr.Base].DxfRev = $pr.Rev }
        }

        # --- Build notes lines from collected drawings + order parts ---
        # Format: "17069-10-P06 Rev.3 Upper Drawbar Qty: 4"
        $noteLines = @()
        $releasedDrawings = @($drawMap.Values | Where-Object { $_.HasPdf -or $_.HasDxf })
        if ($releasedDrawings.Count -gt 0) {
            $noteLines = @($releasedDrawings | ForEach-Object {
                $drawEntry = $_   # save outer $_ before Where-Object shadows it
                # Find matching order line for description and qty
                $orderPart = $Parts | Where-Object { ([string]$_.Part).Trim().ToUpperInvariant() -eq $drawEntry.Part } | Select-Object -First 1
                $displayPart = $drawEntry.Part
                if ($null -ne $orderPart) {
                    $custPart = ([string]$orderPart.CustPart).Trim()
                    if (-not [string]::IsNullOrWhiteSpace($custPart) -and $custPart.ToUpperInvariant() -ne $drawEntry.Part.ToUpperInvariant()) {
                        $displayPart = "$custPart our part # $($drawEntry.Part)"
                    }
                }
                $line = $displayPart
                $rev = if ($drawEntry.PdfRev) { $drawEntry.PdfRev } elseif ($drawEntry.DxfRev) { $drawEntry.DxfRev } else { "" }
                if ($rev) { $line += " Rev.$rev" }
                if ($null -ne $orderPart) {
                    $desc = Format-Desc $orderPart.Desc
                    if ($desc) { $line += " $desc" }
                    $fq = Format-Qty $orderPart.Qty
                    if ($fq) { $line += " Qty: $fq" }
                }
                $line
            })
        } else {
            # No collected files - list ALL parts from order (not just those with drawings)
            $noteLines = @($Parts | ForEach-Object {
                $displayPart = [string]$_.Part
                $custPart = ([string]$_.CustPart).Trim()
                if (-not [string]::IsNullOrWhiteSpace($custPart) -and $custPart.ToUpperInvariant() -ne $displayPart.ToUpperInvariant()) {
                    $displayPart = "$custPart our part # $displayPart"
                }
                $line = $displayPart
                $rev = if ($_.Rev -and $_.Rev -ne "NA") { $_.Rev } elseif ($_.IndexRev -and $_.IndexRev -ne "N/A") { $_.IndexRev } else { "" }
                if ($rev) { $line += " Rev.$rev" }
                $desc = Format-Desc $_.Desc
                if ($desc) { $line += " $desc" }
                $fq = Format-Qty $_.Qty
                if ($fq) { $line += " Qty: $fq" }
                $line
            })
        }

        # --- Build body text with part table + paths ---
        $partTable = ($Parts | ForEach-Object {
            $epicorTag = if ($_.EpicorRev) { "Epicor Rev $($_.EpicorRev)" } else { "" }
            $drawTag   = if ($_.HasDrawing) { "PDF Rev $($_.IndexRev)" } else { "no PDF" }
            $dxfTag    = if ($_.HasDxf) { "DXF found" } else { "no DXF" }
            $pathTag   = if ($_.PdfPath) { "  -> $($_.PdfPath)" } else { "" }
            $revTag    = if ($_.RevMatch -eq "Mismatch") { " [REV MISMATCH]" } elseif ($_.RevMatch -eq "Match") { " [rev OK]" } else { "" }
            $tags      = (@($epicorTag, $drawTag, $dxfTag) | Where-Object { $_ }) -join " | "
            "  Line $($_.Line): $($_.Part)  OrderRev=$($_.Rev)  Qty=$($_.Qty)  [$tags]$revTag$pathTag"
        }) -join "`n"

        $flagSection = ""
        if ($FlaggedParts.Count -gt 0) {
            $flagSection = "`n`nREV MISMATCHES - REVIEW REQUIRED:`n" + ($FlaggedParts -join "`n")
        }
        $missingSection = ""
        if ($MissingParts.Count -gt 0) {
            $missingSection = "`n`nPARTS WITHOUT DRAWINGS (purchased/standard - no collection needed):`n" + ($MissingParts -join "`n")
        }
        $testNote = if ($TestMode) { "`n`n*** TEST MODE - draft only, not sent to production ***" } else { "" }

        $notesText = if ($noteLines.Count -gt 0) {
            "The following drawings have been released for $jobCategory job # $effectiveJob`n" + ($noteLines -join "`n")
        } else {
            "No drawing files collected for $jobCategory job # $effectiveJob"
        }

        # --- Try to populate F02 template HTML if loaded ---
        $htmlTemplate = ""
        try { $htmlTemplate = [string]$mail.HTMLBody } catch { $htmlTemplate = "" }
        if (-not [string]::IsNullOrWhiteSpace($htmlTemplate) -and $htmlTemplate.Length -gt 200) {
            $h = $htmlTemplate
            # Replace job number placeholder
            $h = $h.Replace("265??", $effectiveJob)
            $h = $h.Replace("T0?", $transmittalNo)

            # Path placeholders - derive project path from source PDF path
            # e.g. "J:\Epicor\Orders\Spare Parts\27000-27199\27131 - ...\50 - Sales\SO.pdf"
            # -> project root = "...\27131 - Alamos Gold YD - Upper and Lower Drawbars and Pins"
            $projectPath = ""
            $burnProfilePath = ""
            $cadLinkPath = ""
            if ($SourcePdfPath) {
                # Walk up from PDF to find the job folder (contains the job number)
                $parentDir = Split-Path $SourcePdfPath -Parent
                # If we're in a subfolder like "50 - Sales", go up one level
                if ((Split-Path $parentDir -Leaf) -match '^\d+\s*-') {
                    $projectPath = Split-Path $parentDir -Parent
                } else {
                    $projectPath = $parentDir
                }
                # Look for NMT_PDM project path with matching job number
                $pdmBase = "C:\NMT_PDM\Projects"
                if (Test-Path $pdmBase) {
                    $pdmMatch = Get-ChildItem $pdmBase -Directory -Recurse -Depth 2 -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -match "^$effectiveJob\b" } | Select-Object -First 1
                    if ($pdmMatch) {
                        $designDir = Join-Path $pdmMatch.FullName "3 - Design\Drawings"
                        $burnDir = Join-Path $designDir "Burn Profiles"
                        if (Test-Path $designDir) { $cadLinkPath = $designDir }
                        if (Test-Path $burnDir)   { $burnProfilePath = $burnDir }
                    }
                }
            }
            $pathPlaceholder = "J:\Epicor\Orders\Capital\??"
            # First occurrence: Burn Profiles path, Second: CADLink/Design\Drawings path
            $burnPathHtml = if ($burnProfilePath) { [System.Security.SecurityElement]::Escape($burnProfilePath) } else { [System.Security.SecurityElement]::Escape($OutputFolder) }
            $cadPathHtml  = if ($cadLinkPath) { [System.Security.SecurityElement]::Escape($cadLinkPath) } else { [System.Security.SecurityElement]::Escape($OutputFolder) }
            $idx1 = $h.IndexOf($pathPlaceholder)
            if ($idx1 -ge 0) {
                $h = $h.Substring(0, $idx1) + $burnPathHtml + $h.Substring($idx1 + $pathPlaceholder.Length)
            }
            $idx2 = $h.IndexOf($pathPlaceholder)
            if ($idx2 -ge 0) {
                $h = $h.Substring(0, $idx2) + $cadPathHtml + $h.Substring($idx2 + $pathPlaceholder.Length)
            }

            # Notes section - inject into the cell after "Notes / Reason for Change" header
            # The F02 template has a table row with empty content after the header row.
            # Find the content cell (after Reason for Change row) and replace its &nbsp; placeholder.
            $rcMarker = "Reason for Change"
            $rcIdx = $h.IndexOf($rcMarker)
            if ($rcIdx -ge 0) {
                # Find the next </tr> (end of header row), then the content cell's <b><o:p>&nbsp;
                $trEnd = $h.IndexOf("</tr>", $rcIdx)
                if ($trEnd -ge 0) {
                    # Look for the empty placeholder in the content row
                    $nbspMarker = "<b><o:p>&nbsp;</o:p></b>"
                    $nbspIdx = $h.IndexOf($nbspMarker, $trEnd)
                    $burnIdx = $h.IndexOf("Burn Profiles", $trEnd)
                    # Only replace if the &nbsp; is before "Burn Profiles" (i.e., it's in the notes cell)
                    if ($nbspIdx -ge 0 -and ($burnIdx -lt 0 -or $nbspIdx -lt $burnIdx)) {
                        # Build HTML notes: each line as a paragraph
                        $notesHtmlLines = ($notesText -split "`n") | ForEach-Object {
                            $escaped = [System.Security.SecurityElement]::Escape($_)
                            "<span style='font-size:10.0pt'>$escaped</span><br>"
                        }
                        $notesHtml = $notesHtmlLines -join "`n"
                        $h = $h.Substring(0, $nbspIdx) + $notesHtml + $h.Substring($nbspIdx + $nbspMarker.Length)
                    }
                }
            }

            $mail.HTMLBody = $h
            Write-Log "F02 template populated with job $effectiveJob, transmittal $transmittalNo" "INFO"
        } else {
            # Fallback: plain text body
            $mail.Body = @"
Sales Order Drawing Package - Transmittal $transmittalNo
Order: $OrderNumber
Client: $Client
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')$testNote

$notesText

PARTS LIST:
$partTable$flagSection$missingSection

Drawings output folder: $OutputFolder

-- NMT Drawing Hub (Automated - EpicorOrderMonitor)
"@
        }

        # Attach collected drawings
        if (Test-Path $OutputFolder) {
            foreach ($f in (Get-ChildItem $OutputFolder -File -ErrorAction SilentlyContinue)) {
                try { $mail.Attachments.Add($f.FullName) | Out-Null } catch { }
            }
            $dxfSub = Join-Path $OutputFolder "DXFs"
            if (Test-Path $dxfSub) {
                foreach ($f in (Get-ChildItem $dxfSub -File -ErrorAction SilentlyContinue)) {
                    try { $mail.Attachments.Add($f.FullName) | Out-Null } catch { }
                }
            }
        }

        $mail.Save()   # Saves to Drafts - does NOT auto-send
        Write-Log "Transmittal $transmittalNo draft saved for SO $OrderNumber -> To: $transmittalTo ($($CollectedPdfs.Count) PDFs, $($CollectedDxfs.Count) DXFs)" "SUCCESS"
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail)    | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
        return $transmittalNo
    } catch {
        Write-Log "Transmittal draft error: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# ==============================================================================
# History (transmittal_history.json - same format as main Hub)
# ==============================================================================

function Add-ToHistory {
    param([object]$Entry)
    $hPath   = Join-Path $indexFolder "transmittal_history.json"
    $history = [System.Collections.Generic.List[object]]::new()
    if (Test-Path $hPath) {
        try {
            $raw = Get-Content $hPath -Raw | ConvertFrom-Json
            if ($raw -is [System.Array]) { foreach ($i in $raw) { $history.Add($i) } }
            elseif ($null -ne $raw)      { $history.Add($raw) }
        } catch { }
    }
    $history.Insert(0, $Entry)
    while ($history.Count -gt 100) { $history.RemoveAt($history.Count - 1) }
    $lines = foreach ($h in $history) { $h | ConvertTo-Json -Depth 8 -Compress }
    [System.IO.File]::WriteAllText($hPath, "[`n" + ($lines -join ",`n") + "`n]", [System.Text.Encoding]::UTF8)
}

# ==============================================================================
# Processed Email Cache
# ==============================================================================

function Get-ProcessedCache {
    if (Test-Path $processedCachePath) {
        try { return [System.Collections.Generic.HashSet[string]](@(Get-Content $processedCachePath -Raw | ConvertFrom-Json)) }
        catch { }
    }
    return [System.Collections.Generic.HashSet[string]]::new()
}

function Add-ToProcessedCache {
    param([string]$EntryId)
    $cache = Get-ProcessedCache
    $cache.Add($EntryId) | Out-Null
    $arr   = @($cache)
    # Keep last 2000 entries
    if ($arr.Count -gt 2000) { $arr = $arr | Select-Object -Last 2000 }
    $arr | ConvertTo-Json | Set-Content $processedCachePath -Encoding UTF8 -Force
}

# ==============================================================================
# Core Processor - one sales order PDF
# ==============================================================================

function Invoke-ProcessSalesOrder {
    param(
        [string]$PdfPath,
        [string]$OcrTextDirect = "",   # Pass text directly (test/RunTest mode)
        [string]$EmailId       = "",
        [string]$EmailSubject  = ""
    )

    Write-Log "--- Processing Sales Order ---" "INFO"

    # Step 1: Get text from PDF
    $ocrText = $OcrTextDirect
    if ([string]::IsNullOrWhiteSpace($ocrText)) {
        Write-Log "Step 1/5: Extracting text from PDF..." "INFO"
        Write-EmailProgress "processing" "" "Extracting text from $([System.IO.Path]::GetFileName($PdfPath))"
        $ocrText = Get-PdfText -PdfPath $PdfPath
    } else {
        Write-Log "Step 1/5: Using pre-supplied text ($($ocrText.Length) chars)" "INFO"
    }
    if ([string]::IsNullOrWhiteSpace($ocrText)) {
        Write-Log "No text extracted - skipping" "WARN"
        return $null
    }
    Write-Log "Step 1/5 done: $($ocrText.Length) chars of text ready" "SUCCESS"

    # Step 2: Parse parts
    Write-Log "Step 2/5: Parsing parts list from order text..." "INFO"
    Write-EmailProgress "processing" "" "Extracting parts list"
    $extraction  = Get-SalesOrderParts -Text $ocrText
    $parts       = @($extraction.Parts)
    $orderNumber = $extraction.OrderNumber
    $client      = $extraction.Client

    # Fallback: if order number not found in text, try to parse from filename
    # e.g. "Sales Order Acknowledgment_27122_REV_0.pdf" or "SO_27122.pdf"
    if ($orderNumber -eq "UNKNOWN" -and $PdfPath) {
        $fname = [System.IO.Path]::GetFileNameWithoutExtension($PdfPath)
        if ($fname -match '(?:_|SO|SO[-_]?)(\d{4,6})(?:_|$|-|\.)')  {
            $orderNumber = $matches[1]
            Write-Log "Order number from filename: $orderNumber" "INFO"
        } elseif ($fname -match '(?<!\d)(\d{4,6})(?!\d)') {
            $orderNumber = $matches[1]
            Write-Log "Order number from filename fallback: $orderNumber" "INFO"
        }
    }

    Write-Log "Order $orderNumber : $($parts.Count) parts found, Client: $client" "INFO"

    Write-Log "Step 2/5 done: $($parts.Count) parts extracted from order $orderNumber" "SUCCESS"

    if ($parts.Count -eq 0) {
        Write-Log "No parts extracted - check PDF format" "WARN"
        Write-Log "Raw text sample (first 500 chars):`n$($ocrText.Substring(0, [Math]::Min(500,$ocrText.Length)))" "INFO"
        return @{ OrderNumber = $orderNumber; Client = $client; Parts = @(); Success = $false }
    }

    # Step 3: Validate parts
    # Priority:  (A) Epicor REST API  ->  (B) Drawing index (pdf_index_clean.csv)
    # Epicor API: confirms part exists + gets approved revisions
    # Drawing index: confirms we have a drawing file to collect + its rev on file
    $apiMode = if ($epicorEnabled) { "Epicor API + Drawing Index" } else { "Drawing Index only (set epicor.password in config.json to enable API)" }
    Write-Log "Step 3/5: Validating $($parts.Count) parts  [$apiMode]..." "INFO"
    Write-EmailProgress "processing" "SO $orderNumber" "Validating parts"
    Initialize-PdfIndex

    $validated     = [System.Collections.Generic.List[object]]::new()
    $flagged       = [System.Collections.Generic.List[string]]::new()
    $noDrawing     = [System.Collections.Generic.List[string]]::new()
    $apiWorked     = $false

    foreach ($p in $parts) {

        # --- (A) Try Epicor API first ---
        $epicorInfo = Get-EpicorPartInfo -PartNumber $p.Part
        if ($null -ne $epicorInfo) { $apiWorked = $true }

        $inEpicor   = if ($null -ne $epicorInfo) { $epicorInfo.Exists } else { $null }
        $epicorRev  = if ($null -ne $epicorInfo -and $epicorInfo.Exists) { $epicorInfo.LatestRev } else { "" }

        if ($null -ne $inEpicor -and -not $inEpicor) {
            # API confirmed this part does NOT exist in Epicor at all
            Write-Log "  [NOT IN EPICOR] $($p.Part) - part not found in Epicor system" "WARN"
        } elseif ($inEpicor) {
            Write-Log "  [IN EPICOR] $($p.Part)  Latest approved rev: $epicorRev" "SUCCESS"
        }

        # --- (B) Check drawing index for file collection + rev on file ---
        $hits       = Find-PartInIndex -PartNumber $p.Part
        $dxfHits    = Find-PartInDxfIndex -PartNumber $p.Part
        $hasDrawing = ($hits.Count -gt 0)
        $hasDxf     = ($dxfHits.Count -gt 0)
        $indexRev   = "N/A"
        $revStatus  = "N/A"
        $pdfPath    = ""
        $dxfPath    = ""

        if ($hasDrawing) {
            $bestHit  = $hits | Sort-Object {
                $rv = Normalize-Rev $_.Rev
                try { [int]$rv } catch { try { [double]$rv } catch { 0 } }
            } -Descending | Select-Object -First 1
            $indexRev = Normalize-Rev $bestHit.Rev
            $pdfPath  = [string]$bestHit.FullPath

            # Rev comparison: prefer Epicor API rev, fall back to index rev
            $compareRev = if ($epicorRev) { $epicorRev } else { $indexRev }
            $revStatus  = Compare-OrderRevToIndex -OrderRev $p.Rev -IndexRev $compareRev

            if ($revStatus -eq "Mismatch") {
                $flag = "Line $($p.Line) - $($p.Part): Order Rev $($p.Rev) vs System Rev $compareRev"
                $flagged.Add($flag)
                Write-Log "  [REV MISMATCH] $flag" "WARN"
            } else {
                Write-Log "  [OK] $($p.Part)  Rev $compareRev - matches order rev $($p.Rev)  PDF: $pdfPath" "SUCCESS"
            }
        } else {
            # No drawing file - check if Epicor rev matches anyway
            if ($epicorRev -and $p.Rev -ne "NA") {
                $revStatus = Compare-OrderRevToIndex -OrderRev $p.Rev -IndexRev $epicorRev
                if ($revStatus -eq "Mismatch") {
                    $flag = "Line $($p.Line) - $($p.Part): Order Rev $($p.Rev) vs Epicor Rev $epicorRev"
                    $flagged.Add($flag)
                    Write-Log "  [REV MISMATCH] $flag" "WARN"
                } else {
                    Write-Log "  [OK] $($p.Part)  Rev $epicorRev - Epicor rev matches order (no drawing file)" "SUCCESS"
                }
            } else {
                $noDrawing.Add($p.Part)
                Write-Log "  [NO DRAWING] $($p.Part) - no drawing file in index" "INFO"
            }
        }

        if ($hasDxf) {
            $bestDxf = $dxfHits | Sort-Object {
                $rv = Normalize-Rev $_.Rev
                try { [int]$rv } catch { try { [double]$rv } catch { 0 } }
            } -Descending | Select-Object -First 1
            $dxfPath = [string]$bestDxf.FullPath
            Write-Log "  [DXF] $($p.Part) -> $dxfPath" "INFO"
        }

        # Use Epicor description if available, fall back to PDF-extracted description
        $partDesc = if ($null -ne $epicorInfo -and $epicorInfo.Exists -and $epicorInfo.Description) {
            $epicorInfo.Description
        } elseif ($p.Desc) { $p.Desc } else { "" }

        $validated.Add([PSCustomObject]@{
            Line       = $p.Line
            Part       = $p.Part
            CustPart   = $p.CustPart
            Rev        = $p.Rev
            Qty        = $p.Qty
            Desc       = $partDesc
            InEpicor   = if ($null -ne $inEpicor) { $inEpicor } else { $hasDrawing }  # best guess if API unavailable
            EpicorRev  = $epicorRev
            HasDrawing = $hasDrawing
            HasDxf     = $hasDxf
            IndexRev   = $indexRev
            RevMatch   = $revStatus
            PdfPath    = $pdfPath
            DxfPath    = $dxfPath
        })
    }

    $partsWithDrawing = @($validated | Where-Object { $_.HasDrawing -or $_.HasDxf } | ForEach-Object { $_.Part })
    # Send ALL non-hardware parts to SimpleCollector - it does its own index matching
    # with parent/child variant logic that this script doesn't replicate
    $allPartNumbers = @($validated | ForEach-Object { $_.Part })
    Write-Log "Step 3/5 done: $($partsWithDrawing.Count) with drawings in index, $($noDrawing.Count) no drawing, $($flagged.Count) rev flags$(if ($apiWorked) { ' [Epicor API OK]' } else { ' [no API]' })" "SUCCESS"

    # Step 4: Collect drawings - send ALL parts to SimpleCollector
    # SimpleCollector has its own multi-strategy lookup (BasePart, FileName, parent/child variants)
    # that is more thorough than our simple index check above.
    $orderOutputDir = Join-Path $outputRoot "SO_$orderNumber"
    Write-Log "Step 4/5: Collecting drawings for $($allPartNumbers.Count) part(s) (sending all to SimpleCollector)..." "INFO"
    Write-EmailProgress "processing" "SO $orderNumber" "Collecting PDFs and DXFs"
    $collected = @{ PdfsCollected = 0; DxfsCollected = 0 }
    if ($allPartNumbers.Count -gt 0) {
        $collected = Invoke-DrawingCollection -PartNumbers $allPartNumbers -OutputFolder $orderOutputDir -Mode $collectMode
    }

    # Post-collection: scan output folder for actual collected files and build path data
    $collectedPdfFiles = @(Get-ChildItem $orderOutputDir -Filter "*.pdf" -ErrorAction SilentlyContinue)
    $dxfSubDir = Join-Path $orderOutputDir "DXFs"
    $collectedDxfFiles = @(Get-ChildItem $dxfSubDir -Filter "*.dxf" -ErrorAction SilentlyContinue)
    Write-Log "Step 4/5 done: $($collectedPdfFiles.Count) PDFs, $($collectedDxfFiles.Count) DXFs collected to $orderOutputDir" "SUCCESS"

    # Step 5: Create transmittal draft
    Write-Log "Step 5/5: Creating Outlook transmittal draft..." "INFO"
    Write-EmailProgress "processing" "SO $orderNumber" "Creating Outlook transmittal draft"
    $transmittalResult = Send-TransmittalDraft `
        -OrderNumber $orderNumber `
        -Client      $client `
        -Parts       @($validated) `
        -FlaggedParts @($flagged) `
        -MissingParts @($noDrawing) `
        -OutputFolder $orderOutputDir `
        -CollectedPdfs $collectedPdfFiles `
        -CollectedDxfs $collectedDxfFiles `
        -SourcePdfPath $PdfPath
    $transmittalSent = ($transmittalResult -ne $false)
    $transmittalNo   = if ($transmittalResult -is [string]) { $transmittalResult } else { "T01" }

    # Step 6: Build summary and write history
    $status = if ($flagged.Count -gt 0) { "Rev Mismatches Found" } else { "Complete" }

    $histEntry = [PSCustomObject]@{
        Timestamp        = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        OrderName        = "Sales Order $orderNumber"
        JobNumber        = $orderNumber
        Client           = $client
        TransmittalNo    = $transmittalNo
        PartsExtracted   = $validated.Count
        PdfsCollected    = $collectedPdfFiles.Count
        DxfsCollected    = $collectedDxfFiles.Count
        TransmittalSent  = $transmittalSent
        Status           = $status
        FlaggedParts     = @($flagged)
        PartsNoDrawing   = @($noDrawing)
        MissingParts     = @($flagged)      # dashboard compat: "missing" = rev mismatches worth attention
        NotFoundParts    = @($noDrawing)    # dashboard compat: parts without drawings (informational)
        Parts            = @($validated)
        OutputFolder     = $orderOutputDir
        SourcePdf        = if ($PdfPath) { [System.IO.Path]::GetFileName($PdfPath) } else { "(test data)" }
        TestMode         = $TestMode.IsPresent
    }
    Add-ToHistory -Entry $histEntry

    $hubMsg = "$($validated.Count) parts | $($collectedPdfFiles.Count) PDFs, $($collectedDxfFiles.Count) DXFs collected | $transmittalNo | $($flagged.Count) rev flags"
    Push-HubNotification -Title "SO $orderNumber Processed" -Message $hubMsg -FolderPath $orderOutputDir

    Write-Log "Order $orderNumber done - $($validated.Count) parts, $($collectedPdfFiles.Count) PDFs, $($collectedDxfFiles.Count) DXFs, $transmittalNo, $($flagged.Count) rev flags" "SUCCESS"
    Write-EmailProgress "complete" "SO $orderNumber" "$status"

    return $histEntry
}

# ==============================================================================
# Outlook Email Scanning
# ==============================================================================

function Get-SalesOrderEmails {
    $result = @{ Emails = @(); OutlookObj = $null; Namespace = $null; InboxFolderObj = $null }
    try {
        $outlook   = New-Object -ComObject Outlook.Application
        $ns        = $outlook.GetNamespace("MAPI")

        # Optional: Force Send/Receive before scanning to ensure latest emails are present
        if ($soCfg.forceSendReceiveBeforeScan) {
            Write-Log "Forcing Outlook Send/Receive..." "INFO"
            try {
                $syncs = $ns.SyncObjects
                for ($i = 1; $i -le $syncs.Count; $i++) { $syncs.Item($i).Start() }
                $waitSec = if ($soCfg.sendReceiveWaitSeconds) { [int]$soCfg.sendReceiveWaitSeconds } else { 5 }
                Start-Sleep -Seconds $waitSec
            } catch { Write-Log "Send/Receive trigger failed: $($_.Exception.Message)" "WARN" }
        }

        $folder    = $ns.GetDefaultFolder(6)   # 6 = olFolderInbox
        
        # Navigate to sub-folder. If the config starts with 'Inbox', start from Inbox.
        # Otherwise, start from the Store Root to allow finding top-level folders.
        $parts = ($emailInboxFolder -split '[/\\]') | Where-Object { $_ -and $_ -ne "Inbox" }
        if ($emailInboxFolder -notmatch '^Inbox') {
            $folder = $ns.GetDefaultFolder(6).Parent # Move up to Store Root
        }

        Write-Log "Scanning folder: $($folder.FolderPath)" "INFO"
        foreach ($part in $parts) {
            try {
                $folder = $folder.Folders.Item($part)
                Write-Log "  Navigated to folder: $($folder.FolderPath)" "INFO"
            } catch {
                Write-Log "Could not navigate to Outlook folder: $part" "WARN"
            }
        }

        $processed = Get-ProcessedCache
        $matching  = [System.Collections.Generic.List[object]]::new()
        $totalItems = $folder.Items.Count
        Write-Log "  Checking $totalItems total items in folder..." "INFO"

        # Collect sample of non-matching subjects for diagnostics
        $subjectSamples   = [System.Collections.Generic.List[string]]::new()
        $subjectSampleMax = 30

        foreach ($item in $folder.Items) {
            try {
                if ($item.Class -ne 43) { continue }  # 43 = olMailItem

                # Subject keyword check
                $subjectOk = $false
                foreach ($kw in $subjectKeywords) {
                    if ($item.Subject -match [regex]::Escape($kw)) { $subjectOk = $true; break }
                }
                if (-not $subjectOk) {
                    if ($subjectSamples.Count -lt $subjectSampleMax -and $item.Subject) {
                        $subjectSamples.Add($item.Subject)
                    }
                    continue
                }

                Write-Log "  Matching subject found: $($item.Subject)" "INFO"

                # Not already processed
                if ($processed.Contains($item.EntryID)) {
                    Write-Log "    Skip: Already processed (ID in cache)" "INFO"
                    continue
                }

                # Check for PDF attachment OR F80 docx reference in email body
                $hasPdf = $false
                $pdfNames = @()
                foreach ($att in $item.Attachments) {
                    if ($att.FileName -match '\.pdf$') {
                        $hasPdf = $true
                        $pdfNames += $att.FileName
                    }
                }

                $hasDocxRef = $false
                if (-not $hasPdf) {
                    # PDM notifications are HTML - check HTMLBody for F80 docx reference
                    try {
                        $htmlBody = [string]$item.HTMLBody
                        if ($htmlBody -match '(?i)F80[a]?\.docx') { $hasDocxRef = $true }
                    } catch { }
                    if (-not $hasDocxRef) {
                        try {
                            $plainBody = [string]$item.Body
                            if ($plainBody -match '(?i)F80[a]?\.docx') { $hasDocxRef = $true }
                        } catch { }
                    }
                }

                if (-not $hasPdf -and -not $hasDocxRef) {
                    Write-Log "    Skip: No PDF attachment or F80 docx reference" "INFO"
                    continue
                }

                if ($hasPdf) {
                    Write-Log "    Found $($pdfNames.Count) PDF(s): $($pdfNames -join ', ')" "SUCCESS"
                } else {
                    Write-Log "    Found F80 docx reference in email body" "SUCCESS"
                }
                $matching.Add($item)
            } catch {
                Write-Log "    Error checking email '$($item.Subject)': $($_.Exception.Message)" "ERROR"
            }
        }

        Write-Log "Found $($matching.Count) unprocessed sales order email(s)" "INFO"

        # Write subject samples to debug file so mismatches are visible
        if ($matching.Count -eq 0 -and $subjectSamples.Count -gt 0) {
            $debugPath = Join-Path $indexFolder "subject_debug.json"
            @{
                Timestamp   = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
                FolderPath  = $folder.FolderPath
                Keywords    = $subjectKeywords
                SampleCount = $subjectSamples.Count
                Subjects    = @($subjectSamples)
            } | ConvertTo-Json -Depth 3 | Set-Content $debugPath -Encoding UTF8 -Force
            Write-Log "  No matches - wrote $($subjectSamples.Count) sample subject(s) to $debugPath" "WARN"
        }

        $result.Emails          = @($matching)
        $result.OutlookObj      = $outlook
        $result.Namespace       = $ns
        $result.InboxFolderObj  = $folder
    } catch {
        Write-Log "Outlook connection error: $($_.Exception.Message)" "ERROR"
    }
    return $result
}

# ==============================================================================
# Built-in Test Data (same orders as Test-Extraction.ps1)
# ==============================================================================

function Get-TestOrderText {
    param([int]$Index = 0)

    $order1 = @"
Sales Order Acknowledgement 1 of 2
Sales Order: 26816
Sold To: Freeport-McMoran/P.T. Freeport Indonesia
333 North Central Ave.
Phoenix AZ 85004-2121 United States
Order Date:09-Jun-2025   Need By: 21-Aug-2025
Terms:Net 15   Ship Via:Land
Line Part Number/Description Rev Order Qty Unit Price Ext. Price
1 4823-P4-34 11 1.00EA 2,770.080 2,770.08
RUBBER
Rel Date Quantity
1 21-Aug-2025 1.00
Supplied w Hardware
2 FHCS-0.625-11x2.000-ZP F 6.00EA 0.000 0.00
FLAT HEAD 5/8-11 X 2 Lg - ZINC PLATED
Rel Date Quantity
1 21-Aug-2025 6.00
Hardware for 4823-P4-34.
3 FW-0.625-SAE-ZP NA 6.00EA 0.000 0.00
FLAT WASHER 5/8 - SAE - ZINC PLATED
4 HN-0.625-11-G5-ZP 1 6.00EA 0.000 0.00
HEX NUT 5/8-11 GR 5 - ZINC PLATED
"@

    $order2 = @"
Sales Order Acknowledgement 1 of 1
Sales Order: 27111
Sold To: Freeport-McMoran/P.T. Freeport Indonesia
333 North Central Ave.
Phoenix AZ 85004-2121 United States
Order Date:12-Mar-2026   Need By: 08-Jun-2026
Terms:Net 60
Line Part Number/Description Rev Order Qty Unit Price Ext. Price
1 BRG03 1 10.00EA 237.750 2,377.50
TAPERED ROLLER BRG
Rel Need By Ship By Quantity
1 08-Jun-2026 01-Jun-2026 10.00
2 4823-P6 3 4.00EA 4,910.000 19,640.00
HINGE ASSEMBLY
Rel Need By Ship By Quantity
1 08-Jun-2026 01-Jun-2026 4.00
3 JOB CHARGES
JOB/ ORDER ADDITIONAL CHARGES
"@

    return @($order1, $order2)[$Index % 2]
}

# ==============================================================================
# MAIN
# ==============================================================================

Write-Log "=== EpicorOrderMonitor v1.0 starting ===" "INFO"
Write-Log "Mode: $(if ($TestMode) { 'TEST' } else { 'PRODUCTION' })  Config: $configPath" "INFO"

# --- RunTest mode: process built-in sample orders without email or PDF ---
if ($RunTest) {
    Write-Log "RunTest mode - using built-in sample orders" "INFO"
    Write-EmailProgress "processing" "" "Running built-in test orders"
    foreach ($idx in 0,1) {
        $testText = Get-TestOrderText -Index $idx
        Invoke-ProcessSalesOrder -OcrTextDirect $testText
    }
    Write-EmailProgress "idle" "" ""
    exit 0
}

# --- Direct PDF mode: single PDF specified on command line ---
if ($PdfPath -and (Test-Path $PdfPath)) {
    Write-Log "Direct PDF mode: $PdfPath" "INFO"
    Write-EmailProgress "processing" "" "Direct PDF: $([System.IO.Path]::GetFileName($PdfPath))"
    Invoke-ProcessSalesOrder -PdfPath $PdfPath
    Write-EmailProgress "idle" "" ""
    exit 0
}

# --- Normal email scan mode ---
Write-EmailProgress "scanning" "" "Scanning Outlook for sales order emails"
$emailResult = Get-SalesOrderEmails
$emails      = $emailResult.Emails

if ($emails.Count -eq 0) {
    Write-Log "No new sales order emails found" "INFO"
    @{ Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); Status = "No orders found" } |
        ConvertTo-Json | Set-Content (Join-Path $indexFolder "last_email_summary.json") -Encoding UTF8 -Force
    Write-EmailProgress "idle" "" ""
    exit 0
}

Write-EmailProgress "found" "" "$($emails.Count) sales order email(s) to process"
Write-Log "Processing $($emails.Count) email(s)" "INFO"

# Navigate / create the processed folder in Outlook (production only)
$outlookProcessedFolder = $null
if (-not $TestMode -and $emailResult.Namespace) {
    try {
        $outlookProcessedFolder = $emailResult.Namespace.GetDefaultFolder(6)
        $pfParts = ($processedFolderName -split '[/\\]') | Where-Object { $_ -and $_ -ne "Inbox" }
        foreach ($part in $pfParts) {
            try {
                $outlookProcessedFolder = $outlookProcessedFolder.Folders.Item($part)
            } catch {
                try { $outlookProcessedFolder = $outlookProcessedFolder.Folders.Add($part) }
                catch { $outlookProcessedFolder = $null; break }
            }
        }
    } catch { $outlookProcessedFolder = $null }
}

foreach ($email in $emails) {
    $entryId = ""
    try {
        $entryId = $email.EntryID
        Write-Log "Processing email: $($email.Subject)" "INFO"

        # Try PDF attachment first, then fall back to docx path -> Sales Order PDF on disk
        $tmpDir   = Join-Path $env:TEMP "so_pdf_$(Get-Random)"
        New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
        $savedPdf = $null

        # Strategy A: PDF attachment on the email
        foreach ($att in $email.Attachments) {
            if ($att.FileName -match '\.pdf$') {
                $savedPdf = Join-Path $tmpDir $att.FileName
                $att.SaveAsFile($savedPdf)
                Write-Log "  Saved PDF attachment: $($att.FileName)" "INFO"
                break
            }
        }

        # Strategy B: Extract docx path from PDM notification -> find Sales Order PDF on file server
        if (-not $savedPdf -or -not (Test-Path $savedPdf)) {
            Write-Log "  No PDF attachment - trying PDM docx path flow..." "INFO"
            $bodyText = ""
            try {
                $htmlBody = [string]$email.HTMLBody
                $bodyText = $htmlBody -replace '<[^>]+>', ' ' -replace '&nbsp;', ' ' -replace '&amp;', '&' -replace '&#92;', '\' -replace '\s+', ' '
            } catch {
                try { $bodyText = [string]$email.Body } catch { $bodyText = "" }
            }

            # Extract docx path from email body
            $docxPath = $null
            $docxPatterns = @(
                '(?mi)(\\[^\r\n]+F80[a]?\.docx)',
                '(?mi)(\\[^\r\n]+\.docx)',
                '(?mi)([A-Z]:\\[^\r\n]+\.docx)'
            )
            foreach ($pat in $docxPatterns) {
                $m = [regex]::Match($bodyText, $pat)
                if ($m.Success) { $docxPath = $m.Groups[1].Value.Trim(); break }
            }

            if ($docxPath) {
                Write-Log "  F80 docx path: $docxPath" "INFO"
                $docxFileName = Split-Path $docxPath -Leaf
                $jobMatch = [regex]::Match($docxFileName, '(\d{4,6})')
                if ($jobMatch.Success) {
                    $jobNum = $jobMatch.Groups[1].Value
                    Write-Log "  Job number: $jobNum" "SUCCESS"

                    # Search Epicor file server for Sales Order PDF
                    foreach ($root in @($crawlRoots)) {
                        $searchPaths = @(
                            (Join-Path $root "Orders\Spare Parts"),
                            (Join-Path $root "Epicor\Orders\Spare Parts")
                        )
                        foreach ($sp in $searchPaths) {
                            if (-not (Test-Path $sp)) { continue }
                            $jobFolders = @(Get-ChildItem -Path $sp -Directory -Recurse -Depth 1 -ErrorAction SilentlyContinue |
                                Where-Object { $_.Name -match "^$jobNum\b" })
                            foreach ($jf in $jobFolders) {
                                $salesDir = Join-Path $jf.FullName "50 - Sales"
                                if (Test-Path $salesDir) {
                                    $soPdfs = @(Get-ChildItem -Path $salesDir -Filter "Sales Order Acknowledgment*.pdf" -ErrorAction SilentlyContinue)
                                    if ($soPdfs.Count -eq 0) {
                                        $soPdfs = @(Get-ChildItem -Path $salesDir -Filter "Sales Order*.pdf" -ErrorAction SilentlyContinue)
                                    }
                                    if ($soPdfs.Count -gt 0) {
                                        $savedPdf = ($soPdfs | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
                                        Write-Log "  Found Sales Order PDF: $savedPdf" "SUCCESS"
                                        break
                                    }
                                }
                            }
                            if ($savedPdf -and (Test-Path $savedPdf)) { break }
                        }
                        if ($savedPdf -and (Test-Path $savedPdf)) { break }
                    }
                } else {
                    Write-Log "  Could not extract job number from: $docxFileName" "WARN"
                }
            } else {
                Write-Log "  No docx path found in email body" "WARN"
            }
        }

        if ($savedPdf -and (Test-Path $savedPdf)) {
            $result = Invoke-ProcessSalesOrder -PdfPath $savedPdf -EmailId $entryId -EmailSubject $email.Subject
            if ($result) {
                Add-ToProcessedCache -EntryId $entryId
                if (-not $TestMode -and $outlookProcessedFolder) {
                    try {
                        $email.Move($outlookProcessedFolder) | Out-Null
                        Write-Log "  Email moved to processed folder" "INFO"
                    } catch {
                        Write-Log "  Could not move email: $($_.Exception.Message)" "WARN"
                    }
                }
            }
        } else {
            Write-Log "  No Sales Order PDF found for: $($email.Subject)" "WARN"
        }

        # Cleanup temp
        if ($tmpDir -and (Test-Path $tmpDir)) {
            Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Log "Error processing email (EntryID=$entryId): $($_.Exception.Message)" "ERROR"
    }
}

# Release Outlook COM objects
try {
    if ($emailResult.OutlookObj) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($emailResult.OutlookObj) | Out-Null
    }
} catch { }

Write-EmailProgress "idle" "" ""
Write-Log "=== EpicorOrderMonitor done ===" "INFO"
