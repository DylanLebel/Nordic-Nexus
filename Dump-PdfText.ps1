param([string]$PdfPath)
# Source the file but avoid running the MAIN block if possible
# Since we can't easily stop the MAIN block in the sourced script without modifying it,
# we'll just redefine Get-NativePdfText here or call it carefully.

$scriptDir = $PSScriptRoot
$monitorScript = Join-Path $scriptDir "EpicorOrderMonitor.ps1"

# Extract just the function we need or source carefully
# For now, let's just run a separate small script that only has the extraction logic.

$code = @"
    function Get-TextFromPdfStream {
        param([byte[]]`$StreamBytes, [bool]`$IsFlate)
        try {
            `$raw = `$StreamBytes
            if (`$IsFlate -and `$raw.Length -gt 2) {
                `$ms      = [System.IO.MemoryStream]::new(`$raw, 2, `$raw.Length - 2)
                `$deflate = [System.IO.Compression.DeflateStream]::new(`$ms, [System.IO.Compression.CompressionMode]::Decompress)
                `$out     = [System.IO.MemoryStream]::new()
                `$deflate.CopyTo(`$out)
                `$deflate.Dispose(); `$ms.Dispose()
                `$raw = `$out.ToArray()
            }
            `$cs  = [System.Text.Encoding]::GetEncoding('iso-8859-1').GetString(`$raw)
            `$sb  = [System.Text.StringBuilder]::new()
            `$lastCharWasNL = `$false
            `$ops = [regex]::Matches(`$cs, '\(([^\)\\]*(?:\\.[^\)\\]*)*)\)\s*(?:Tj|TJ)|\[([^\]]*)\]\s*TJ|[-\d.]+\s+[-\d.]+\s+Td\b|\bTD\b|\bT\*\b|\bET\b')
            foreach (`$op in `$ops) {
                if (-not `$op.Groups[1].Success -and -not `$op.Groups[2].Success) {
                    if (`$sb.Length -gt 0) { `$sb.Append("`n") | Out-Null }
                    `$lastCharWasNL = `$true
                    continue
                }
                `$chunks = @()
                if (`$op.Groups[1].Success -and `$op.Groups[1].Value) { `$chunks = @(`$op.Groups[1].Value) }
                elseif (`$op.Groups[2].Success -and `$op.Groups[2].Value) {
                    `$chunks = [regex]::Matches(`$op.Groups[2].Value, '\(([^\)\\]*(?:\\.[^\)\\]*)*)\)') | ForEach-Object { `$_.Groups[1].Value }
                }
                foreach (`$chunk in `$chunks) {
                    `$t = `$chunk -replace '\\n',' ' -replace '\\r',' ' -replace '\\t',' '
                    `$t = [regex]::Replace(`$t, '\\([0-7]{3})', { [char][Convert]::ToInt32(`$args[0].Groups[1].Value, 8) })
                    `$t = `$t -replace '\\(.)', '`$1'
                    if (`$t.Trim()) {
                        if (`$sb.Length -gt 0 -and -not `$lastCharWasNL) {
                            `$last = `$sb.ToString()[-1]
                            if (`$last -ne ' ' -and `$last -ne "`n" -and `$t[0] -ne ' ') { `$sb.Append(' ') | Out-Null }
                        }
                        `$sb.Append(`$t) | Out-Null
                        `$lastCharWasNL = `$false
                    }
                }
            }
            return `$sb.ToString()
        } catch { return "" }
    }

    function Format-EpicorPdfText {
        param([string]`$Text)
        `$t = `$Text -replace "`r`n", "`n" -replace "`r", "`n"
        `$breaks = @('Sales Order\s*[:#]','Sold To\s*:','Ship To\s*:','Bill To\s*:','Our Part\s*:','Line\s*Part\s*Number','Line Total\s*:','Order Total\s*:','Total Tax','Subtotal\s*:','JOB CHARGES','Rel Need By','Rel Date','Ship By','Payment Terms','Sales Representative','Acknowledgement\b')
        foreach (`$kw in `$breaks) { `$t = [regex]::Replace(`$t, "(?<!\n)(`$kw)", "`n`$1") }
        `$t = `$t -replace '\bLinePart\b', 'Line Part'
        `$t = [regex]::Replace(`$t, '\n{3,}', "`n`n")
        return `$t.Trim()
    }

    function Get-NativePdfText {
        param([string]`$PdfPath)
        try {
            `$enc   = [System.Text.Encoding]::GetEncoding('iso-8859-1')
            `$bytes = [System.IO.File]::ReadAllBytes(`$PdfPath)
            `$pdf   = `$enc.GetString(`$bytes)
            `$sb    = [System.Text.StringBuilder]::new()
            `$pos   = 0
            while (`$true) {
                `$si = `$pdf.IndexOf('stream', `$pos)
                if (`$si -lt 0) { break }
                `$ei = `$pdf.IndexOf('endstream', `$si + 6)
                if (`$ei -lt 0) { break }
                `$di = `$si + 6
                while (`$di -lt `$pdf.Length -and (`$pdf[`$di] -eq "`r" -or `$pdf[`$di] -eq "`n")) { `$di++ }
                `$ds = `$pdf.LastIndexOf('<<', `$si)
                `$dict = $(if (`$ds -ge 0) { `$pdf.Substring(`$ds, [Math]::Min(`$si - `$ds, 600)) } else { "" })
                `$isImage = `$dict -match '/Subtype\s*/Image'
                `$isFlate  = `$dict -match '/FlateDecode'
                if (-not `$isImage) {
                    `$rawBytes = `$enc.GetBytes(`$pdf.Substring(`$di, `$ei - `$di))
                    `$text = Get-TextFromPdfStream -StreamBytes `$rawBytes -IsFlate `$isFlate
                    if (`$text.Trim()) { `$sb.Append(`$text).Append("`n") | Out-Null }
                }
                `$pos = `$ei + 9
            }
            `$raw = `$sb.ToString().Trim()
            return Format-EpicorPdfText -Text `$raw
        } catch { return "" }
    }

    `$text = Get-NativePdfText -PdfPath "$PdfPath"
    Write-Host "=== RAW EXTRACTED TEXT ===" -ForegroundColor Cyan
    `$lines = (`$text -split "`n")
    for (`$i = 0; `$i -lt `$lines.Count; `$i++) {
        Write-Host ("{0,3}: [{1}]" -f `$i, `$lines[`$i]) -ForegroundColor Gray
    }
    Write-Host "=== `$(`$lines.Count) lines, `$(`$text.Length) chars ===" -ForegroundColor Cyan
"@

$code | powershell -NoProfile -ExecutionPolicy Bypass
