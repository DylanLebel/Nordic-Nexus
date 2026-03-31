param(
    [Parameter(Mandatory = $true)][string]$BomFile,
    [Parameter(Mandatory = $true)][string]$ConfigPath,
    [Parameter(Mandatory = $true)][string]$OutputFolder,
    [string]$ExpectedSummaryPath = "",
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path $PSCommandPath -Parent
$collectorScript = Join-Path $scriptDir "SimpleCollector.ps1"
if (-not (Test-Path $collectorScript)) {
    throw "SimpleCollector.ps1 not found at $collectorScript"
}

function Convert-CollectorSummaryToComparisonView {
    param([object]$Summary)

    $partViews = @(
        @($Summary.partResults) |
        ForEach-Object {
            [ordered]@{
                Part                   = [string]$_.Part
                PdfFound               = [bool]$_.PdfFound
                DxfFound               = [bool]$_.DxfFound
                PdfCopiedCount         = @($_.PdfCopied).Count
                DxfCopiedCount         = @($_.DxfCopied).Count
                PrimaryPdfMatchCount   = @($_.PrimaryPdfMatches).Count
                PrimaryDxfMatchCount   = @($_.PrimaryDxfMatches).Count
                ExtraPdfCandidateCount = @($_.ExtraPdfCandidates).Count
                ExtraDxfCandidateCount = @($_.ExtraDxfCandidates).Count
                VariantOnly            = [bool]$_.VariantOnly
            }
        } |
        Sort-Object Part
    )

    return [ordered]@{
        RequestedParts  = @(@($Summary.RequestedParts) | Sort-Object)
        PdfsFound       = [int]$Summary.pdfsFound
        DxfsFound       = [int]$Summary.dxfsFound
        NotFound        = @(@($Summary.notFound) | Sort-Object)
        ExtraPdfCount   = @($Summary.extraPdfs).Count
        ExtraDxfCount   = @($Summary.extraDxfs).Count
        Warnings        = @(@($Summary.warnings) | Sort-Object)
        PartResults     = $partViews
    }
}

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$summaryPath = Join-Path $OutputFolder "collector_summary.json"
if (Test-Path $summaryPath) {
    Remove-Item -Path $summaryPath -Force -ErrorAction SilentlyContinue
}

& powershell -NoProfile -ExecutionPolicy Bypass -File $collectorScript $BomFile $OutputFolder "BOTH" $ConfigPath
if ($LASTEXITCODE -ne 0) {
    throw "SimpleCollector exited with code $LASTEXITCODE"
}

if (-not (Test-Path $summaryPath)) {
    throw "Collector summary was not written to $summaryPath"
}

$summary = Get-Content -Path $summaryPath -Raw | ConvertFrom-Json
$view = Convert-CollectorSummaryToComparisonView -Summary $summary

if (-not $Quiet) {
    Write-Host "Replay summary:"
    Write-Host ($view | ConvertTo-Json -Depth 6)
}

if (-not [string]::IsNullOrWhiteSpace($ExpectedSummaryPath)) {
    if (-not (Test-Path $ExpectedSummaryPath)) {
        throw "Expected summary not found: $ExpectedSummaryPath"
    }
    $expectedSummary = Get-Content -Path $ExpectedSummaryPath -Raw | ConvertFrom-Json
    $expectedView = Convert-CollectorSummaryToComparisonView -Summary $expectedSummary
    $actualJson = $view | ConvertTo-Json -Depth 6
    $expectedJson = $expectedView | ConvertTo-Json -Depth 6
    if ($actualJson -ne $expectedJson) {
        Write-Host "Replay mismatch." -ForegroundColor Red
        Write-Host "Expected:" -ForegroundColor Yellow
        Write-Host $expectedJson
        Write-Host "Actual:" -ForegroundColor Yellow
        Write-Host $actualJson
        exit 1
    }
    if (-not $Quiet) {
        Write-Host "Replay matched expected summary." -ForegroundColor Green
    }
}
