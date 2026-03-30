# ==============================================================================
#  Process-SalesOrder.ps1 - Nordic Minesteel Technologies
#  Main CLI Driver for Epicor Sales Order Acknowledgment Processing
# ==============================================================================

param(
    [string]$PdfPath,
    [switch]$Force # Ignore already-processed status
)

$scriptDir = Split-Path $PSCommandPath -Parent
$cfgPath   = Join-Path $scriptDir "config.json"
$cfg       = Get-Content $cfgPath | ConvertFrom-Json
$extractor = Join-Path $scriptDir "Extract-EpicorOrderParts.ps1"

# --- Setup Directories ---
$logDir = Join-Path $scriptDir $cfg.LogFolder
$outDir = Join-Path $scriptDir $cfg.OutputFolder
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir | Out-Null }
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

function Write-HubLog {
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

# --- STEP 1: OCR THE PDF ---
Write-HubLog "Processing: $PdfPath" "INFO"
if (-not (Test-Path $PdfPath)) {
    Write-HubLog "ERROR: PDF file not found at $PdfPath" "ERROR"
    return
}

$tesseract = $cfg.TesseractPath
if (-not (Test-Path $tesseract)) {
    Write-HubLog "ERROR: Tesseract OCR not found at $tesseract" "ERROR"
    return
}

# Direct OCR from PDF (requires Tesseract 3.03+ or specialized setup)
# If Tesseract doesn't handle the PDF directly, you'll need Ghostscript.
$ocrFile = Join-Path $logDir "temp_ocr.txt"
Write-HubLog "Running OCR..." "INFO"
try {
    # Tesseract 4+ can handle multipage PDFs directly (requires 'pdf' in args or standard input)
    & $tesseract $PdfPath ($ocrFile -replace '\.txt','') --psm 6 2>$null
    $ocrText = Get-Content $ocrFile -Raw
} catch {
    Write-HubLog "OCR Failed: $($_.Exception.Message)" "ERROR"
    return
}

# --- STEP 2: EXTRACT PARTS ---
if (-not [string]::IsNullOrWhiteSpace($ocrText)) {
    Write-HubLog "Extracting parts from OCR text..." "SUCCESS"
    # Call the logic script (this will print extraction summary)
    powershell.exe -File $extractor -OcrText $ocrText -SourcePath $PdfPath
} else {
    Write-HubLog "ERROR: OCR returned no text. PDF might be encrypted or scanned poorly." "ERROR"
}

# --- CLEANUP ---
if (Test-Path $ocrFile) { Remove-Item $ocrFile -Force }
