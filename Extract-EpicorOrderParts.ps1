# ==============================================================================
#  Extract-EpicorOrderParts.ps1 - Nordic Minesteel Technologies
#  Extracts Part Numbers from Sales Order Acknowledgement PDFs
# ==============================================================================

param(
    [string]$OcrText,
    [string]$SourcePath,
    [switch]$CheckEpicor  # If set, verify against the "system" (Placeholder for now)
)

# --- Logging ---
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $fg = switch ($Level) {
        "ERROR"   { "Red" }
        "WARN"    { "Yellow" }
        "SUCCESS" { "Green" }
        default   { "Gray" }
    }
    Write-Host "$ts [$Level] $Message" -ForegroundColor $fg
}

# --- Part number patterns from NMT logic (Ordered by specificity) ---
$PartPatterns = @(
    '\b\d{4,6}[-_][A-Z0-9]{1,8}(?:[-_][A-Z0-9]{1,8}){0,4}\b', # NMT style: 4823-P4-34, 1202-WM-440
    '\b[A-Z]{2,5}-\d+[.,]\d{1,3}(?:[-x][A-Z0-9.]+){1,5}\b', # Hardware: FHCS-0.625-11x2.000-ZP
    '\b\d{5}-\d{2}-[A-Z]\d{2,3}(?:-[A-Z])?\b',              # NMT job part: 17141-10-P67
    '\b\d{8,10}\b',                                           # Long Customer Part (e.g. 40617527)
    '\b[A-Z]{1,4}\d{1,5}[A-Z]{0,2}\b',                        # Short multi: MHB100EP, W010A, BRG03
    '\b[A-Z]{1,6}\d{1,4}(?:-[A-Z0-9]{1,6}){1,8}\b',          # Alpha-dash: BF12C11-P
    '\b\d{5}[A-Z]\d{3,4}\b',                                  # McMaster: 90107A030
    '\b[A-Z]{2,4}\d{2,6}(?:-[A-Z]{1,3})?\b',                  # Prefix+digits: PB23056-FS, BRG03
    '\b\d{2}-\d{2}-\d{2,3}(?:-[A-Z0-9]{1,4}){0,3}\b',        # Short assy: 02-05-00-C30-PU
    '\b\d{5}\b'                                               # Plain 5-digit: 76359, 10569
)

function Get-SalesOrderParts {
    param([string]$Text)
    $lines = $Text -split "`n"
    $parts = @()
    $inTable = $false
    $orderNumber = "UNKNOWN"

    foreach ($line in $lines) {
        $line = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        if ($line -match '(?i)Sales Order:\s*(\d+)') {
            $orderNumber = $matches[1]
        }

        # Detect table start
        if ($line -match '(?i)Line\s+Part\s+Number/Description\s+Rev\s+Order\s+Qty') {
            $inTable = $true
            continue
        }

        # Detect table end
        if ($line -match '(?i)(Line Total:|Total Tax|Order Total)') {
            $inTable = $false
            continue
        }

        if ($inTable) {
            # Check for 'Our Part' cross-reference on the current line (description line)
            if ($line -match '(?i)Our Part:\s*([A-Z0-9.-]+)(?:\s+(\d+))?') {
                if ($parts.Count -gt 0) {
                    $parts[-1].Part = $matches[1]
                    if ($matches[2]) { $parts[-1].Rev = $matches[2] }
                    Write-Log "  -> Cross-referenced to NMT Part: $($matches[1]) (Rev: $($parts[-1].Rev))" "SUCCESS"
                }
                continue
            }

            # Skip noise lines
            if ($line -match '(?i)(JOB CHARGES|RUBBER|Rel Date|Quantity|Hardware for|Supplied w|Rel Need By|Ship By|TAPERED ROLLER|HINGE ASSEMBLY|Y-BEARING)') { continue }
            
            # Line items start with a digit (Line Number) followed by a space
            if ($line -match '^\d+\s+') {
                $lineParts = $line -split '\s+'
                $lineNum = $lineParts[0]
                $foundPart = $null
                $pIdx = -1

                # Strategy: Search the entire line for the most specific pattern match
                $bestMatch = $null
                
                # Pre-clean line for better matching (remove line number)
                $cleanLine = $line -replace '^\d+\s+', ''

                for($p=0; $p -lt $PartPatterns.Count; $p++) {
                    if ($cleanLine -match $PartPatterns[$p]) {
                        $match = $matches[0]
                        if ($bestMatch -eq $null -or ($match.Length -gt $bestMatch.Length)) {
                            $bestMatch = $match
                        }
                    }
                }

                if ($bestMatch) {
                    $foundPart = $bestMatch
                    # Find where this part is in the tokens to help find Rev/Qty
                    for($i=1; $i -lt $lineParts.Count; $i++) {
                        if ($lineParts[$i] -contains $foundPart -or $foundPart -contains $lineParts[$i]) {
                            $pIdx = $i
                            break
                        }
                    }

                    $rev = "NA"
                    $qty = "0"

                    # Find Rev (the first token after Part that is NOT a quantity)
                    for($j=$pIdx+1; $j -lt $lineParts.Count; $j++) {
                        $t = $lineParts[$j]
                        if ($t -match '(\d+(?:\.\d+)?)(?:EA|ea)') {
                            $qty = $matches[1]
                            break
                        }
                        # Rev cleanup: ignore single letters like 'F', 'P', 'X' if they aren't 'NA'
                        if ($rev -eq "NA" -and $t -match '^[A-Z0-9/.-]{1,5}$') {
                            if ($t.Length -eq 1 -and $t -match '[A-Z]') { continue }
                            $rev = $t
                        }
                    }

                    $parts += [PSCustomObject]@{
                        Order = $orderNumber
                        Line  = $lineNum
                        Part  = $foundPart
                        Rev   = $rev
                        Qty   = $qty
                    }
                }
            }
        }
    }
    return $parts
}

# --- EPICOR SYSTEM CHECK (Placeholder) ---
# Currently checks if the part exists in our PDF Index as a proxy.
# Returns $true if found, $false if missing.
function Check-EpicorSystem {
    param([string]$PartNumber)
    
    # In the future, this could be:
    # return (Invoke-Sql "SELECT Count(*) FROM Parts WHERE PartNum = '$PartNumber'") -gt 0
    
    # FOR NOW: Check if we have a drawing for it in the PDF Index
    # (Since we're building this in a new folder, we'll try to find the index folder)
    $indexDir = "C:\\Users\\dlebel\\Documents\\PDFIndex"
    $cleanCSV = Join-Path $indexDir "pdf_index_clean.csv"
    
    if (Test-Path $cleanCSV) {
        $found = Import-Csv $cleanCSV | Where-Object { $_.PartNumber -eq $PartNumber }
        if ($found) { return $true }
    }
    
    # If not in index, simulate a check against a hypothetical "System List"
    # (Example: Hard-coded known-missing part for testing)
    $systemKnownParts = @("4823-P4-34", "HN-0.625-11-G5-ZP", "BRG03", "4823-P6", "1206-P404")
    if ($systemKnownParts -contains $PartNumber) { return $true }

    return $false
}

# --- Execution Logic ---
if ($OcrText) {
    $results = Get-SalesOrderParts -Text $OcrText
    if ($results.Count -eq 0) {
        Write-Log "No parts extracted from OCR text." "WARN"
        return
    }
    
    Write-Host "`n=== Extraction Summary for Order: $($results[0].Order) ===" -ForegroundColor Cyan
    
    foreach ($r in $results) {
        $exists = Check-EpicorSystem -PartNumber $r.Part
        $status = if ($exists) { "[OK]" } else { "[MISSING]" }
        $color  = if ($exists) { "Green" } else { "Red" }
        
        Write-Host "$status Line $($r.Line): Part '$($r.Part)' (Rev: $($r.Rev), Qty: $($r.Qty))" -ForegroundColor $color
        
        if (-not $exists) {
            Write-Log "FLAG: Part '$($r.Part)' not found in Epicor system! Manual intervention required." "ERROR"
        }
    }
}
