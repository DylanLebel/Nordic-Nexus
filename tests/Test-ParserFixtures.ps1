param(
    [string]$FixtureRoot = (Join-Path $PSScriptRoot "parser-fixtures")
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $repoRoot "SalesOrderParsing.ps1")

$fixtureNames = @("26905", "27131", "27129", "27134")
$failed = $false

foreach ($name in $fixtureNames) {
    $textPath = Join-Path $FixtureRoot "$name.txt"
    $expectedPath = Join-Path $FixtureRoot "$name.expected.json"
    if (-not (Test-Path $textPath)) { throw "Missing fixture text: $textPath" }
    if (-not (Test-Path $expectedPath)) { throw "Missing fixture expectation: $expectedPath" }

    $text = Get-Content -Path $textPath -Raw
    $actual = @(Parse-SalesOrderText -Text $text)
    $expected = @((Get-Content -Path $expectedPath -Raw | ConvertFrom-Json))

    $actualJson = $actual | ConvertTo-Json -Depth 5
    $expectedJson = $expected | ConvertTo-Json -Depth 5

    if ($actualJson -ne $expectedJson) {
        $failed = $true
        Write-Host ""
        Write-Host "Fixture failed: $name" -ForegroundColor Red
        Write-Host "Expected:" -ForegroundColor Yellow
        Write-Host $expectedJson
        Write-Host "Actual:" -ForegroundColor Yellow
        Write-Host $actualJson
    } else {
        Write-Host "Fixture passed: $name ($($actual.Count) part(s))" -ForegroundColor Green
    }
}

if ($failed) {
    exit 1
}

Write-Host "All parser fixtures passed." -ForegroundColor Green
