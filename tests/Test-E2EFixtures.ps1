param(
    [string]$FixtureRoot = (Join-Path $PSScriptRoot "e2e-fixtures"),
    [string]$ScratchRoot = (Join-Path $PSScriptRoot "e2e-scratch")
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path $PSScriptRoot -Parent
$replayScript = Join-Path $repoRoot "Replay-OrderRun.ps1"
if (-not (Test-Path $replayScript)) {
    throw "Replay-OrderRun.ps1 not found at $replayScript"
}

$fixtures = @(Get-ChildItem -Path $FixtureRoot -Directory -ErrorAction Stop)
if ($fixtures.Count -eq 0) {
    throw "No fixture directories found in $FixtureRoot"
}

$failed = $false
foreach ($fixtureDir in $fixtures) {
    $fixtureJsonPath = Join-Path $fixtureDir.FullName "fixture.json"
    $bomPath = Join-Path $fixtureDir.FullName "order_bom.txt"
    $configPath = Join-Path $fixtureDir.FullName "config.fixture.json"
    $expectedSummaryPath = Join-Path $fixtureDir.FullName "expected_collector_summary.json"
    if (-not (Test-Path $fixtureJsonPath)) { throw "Missing fixture.json in $($fixtureDir.FullName)" }
    if (-not (Test-Path $bomPath)) { throw "Missing order_bom.txt in $($fixtureDir.FullName)" }
    if (-not (Test-Path $configPath)) { throw "Missing config.fixture.json in $($fixtureDir.FullName)" }
    if (-not (Test-Path $expectedSummaryPath)) { throw "Missing expected_collector_summary.json in $($fixtureDir.FullName)" }

    $scratchOut = Join-Path $ScratchRoot $fixtureDir.Name
    if (Test-Path $scratchOut) {
        Remove-Item -Path $scratchOut -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Path $scratchOut -Force | Out-Null

    Write-Host "Testing fixture $($fixtureDir.Name)..." -ForegroundColor Cyan
    & powershell -NoProfile -ExecutionPolicy Bypass -File $replayScript `
        -BomFile $bomPath `
        -ConfigPath $configPath `
        -OutputFolder $scratchOut `
        -ExpectedSummaryPath $expectedSummaryPath `
        -Quiet
    if ($LASTEXITCODE -ne 0) {
        $failed = $true
    } else {
        Write-Host "Fixture passed: $($fixtureDir.Name)" -ForegroundColor Green
    }
}

if ($failed) {
    exit 1
}

Write-Host "All E2E fixtures passed." -ForegroundColor Green
