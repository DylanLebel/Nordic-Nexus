param(
    [string]$FixtureRoot = (Join-Path $PSScriptRoot "e2e-fixtures")
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path $PSScriptRoot -Parent
. (Join-Path $repoRoot "SalesOrderParsing.ps1")
$replayScript = Join-Path $repoRoot "Replay-OrderRun.ps1"

function Get-CanonicalPartKey {
    param([string]$NameOrPath)
    if ([string]::IsNullOrWhiteSpace($NameOrPath)) { return "" }
    $value = ([string]$NameOrPath).Trim().ToUpperInvariant()
    try {
        if ($value -match '[\\/]') {
            $value = [System.IO.Path]::GetFileNameWithoutExtension($value).Trim().ToUpperInvariant()
        }
    } catch { }
    $value = $value -replace '(?i)[ _-]*REV[ _-]*[A-Z0-9]+$', ''
    return $value.Trim()
}

function Get-AllowedChildParentKey {
    param([string]$CandidatePart)
    if ([string]::IsNullOrWhiteSpace($CandidatePart)) { return "" }
    $cand = $CandidatePart.Trim().ToUpperInvariant()
    $dashPos = $cand.LastIndexOf('-')
    if ($dashPos -le 0 -or $dashPos -ge ($cand.Length - 1)) { return "" }
    $parent = $cand.Substring(0, $dashPos)
    $suffix = $cand.Substring($dashPos + 1)
    if ($suffix -match '^(?:\d{1,3}|[LR]|\d{1,3}[A-Z]?|CL\d{1,3}|PL\d{1,3}|[A-Z]\d{1,3})$') {
        return $parent
    }
    return ""
}

function New-FixtureConfig {
    param(
        [string]$IndexFolder,
        [string]$OutputFolder
    )

    return [ordered]@{
        indexFolder = $IndexFolder
        outputFolder = $OutputFolder
        pathFilters = [ordered]@{
            disallowedModelFolderNames = @("Obsolete", "Archive", "Old", "Deprecated", "QA", "20 - QA", "Quotes")
            disallowedModelNamePatterns = @("*-OBS*", "* OBS")
        }
        collector = [ordered]@{
            strictMatching = $true
            copyChildVariants = $false
            indexStaleWarningHours = 24
        }
    }
}

function Convert-IndexRowForFixture {
    param([object]$Row)
    $fileName = [string]$Row.FileName
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = [System.IO.Path]::GetFileName([string]$Row.FullPath)
    }
    return [pscustomobject]@{
        BasePart = [string]$Row.BasePart
        FileName = $fileName
        FullPath = [string]$Row.FullPath
        Rev = [string]$Row.Rev
        FileType = [string]$Row.FileType
        IsObsolete = [string]$Row.IsObsolete
    }
}

function New-MockIndexedRows {
    param(
        [object[]]$Rows,
        [string]$Root,
        [string]$Kind
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $seenNames = @{}
    $kindRoot = Join-Path $Root $Kind
    New-Item -ItemType Directory -Path $kindRoot -Force | Out-Null

    foreach ($row in @($Rows)) {
        $baseRow = Convert-IndexRowForFixture -Row $row
        $fileName = [string]$baseRow.FileName
        if ([string]::IsNullOrWhiteSpace($fileName)) { continue }
        $nameKey = $fileName.ToUpperInvariant()
        $bucket = 1
        if ($seenNames.ContainsKey($nameKey)) {
            $bucket = [int]$seenNames[$nameKey] + 1
        }
        $seenNames[$nameKey] = $bucket
        $subdir = Join-Path $kindRoot ("item_{0:D3}" -f $bucket)
        New-Item -ItemType Directory -Path $subdir -Force | Out-Null
        $mockPath = Join-Path $subdir $fileName
        [System.IO.File]::WriteAllText($mockPath, "$Kind fixture placeholder for $fileName", [System.Text.Encoding]::UTF8)

        $results.Add([pscustomobject]@{
            BasePart = [string]$baseRow.BasePart
            FileName = $fileName
            FullPath = $mockPath
            Rev = [string]$baseRow.Rev
            FileType = [string]$baseRow.FileType
            IsObsolete = [string]$baseRow.IsObsolete
        })
    }

    return @($results)
}

$configTest = Get-Content -Path (Join-Path $repoRoot "config.test.json") -Raw | ConvertFrom-Json
$liveIndexFolder = [string]$configTest.indexFolder
$livePdfIndexPath = Join-Path $liveIndexFolder "pdf_index_clean.csv"
$liveDxfIndexPath = Join-Path $liveIndexFolder "dxf_index_clean.csv"
if (-not (Test-Path $livePdfIndexPath)) { throw "Live PDF index not found: $livePdfIndexPath" }
if (-not (Test-Path $liveDxfIndexPath)) { throw "Live DXF index not found: $liveDxfIndexPath" }

$livePdfRows = @(Import-Csv -Path $livePdfIndexPath)
$liveDxfRows = @(Import-Csv -Path $liveDxfIndexPath)

$fixtureSpecs = @(
    [ordered]@{
        Name = "26905"
        JobNumber = "26905"
        Description = "Expanded assembly order with strict-matching extras"
        SourceBomPath = "C:\Temp\NMT_HubTest\AssemblyPDFs\Orders\20260331_105916_Order\order_bom.txt"
    },
    [ordered]@{
        Name = "27131"
        JobNumber = "27131"
        Description = "Multi-line order with straightforward collected drawings"
        SourceParserFixture = (Join-Path $PSScriptRoot "parser-fixtures\27131.txt")
    },
    [ordered]@{
        Name = "27129"
        JobNumber = "27129"
        Description = "Customer part cross-reference and split revision parsing"
        SourceParserFixture = (Join-Path $PSScriptRoot "parser-fixtures\27129.txt")
    },
    [ordered]@{
        Name = "27134"
        JobNumber = "27134"
        Description = "Clean hardware-heavy order"
        SourceParserFixture = (Join-Path $PSScriptRoot "parser-fixtures\27134.txt")
    }
)

foreach ($spec in $fixtureSpecs) {
    $fixtureDir = Join-Path $FixtureRoot $spec.Name
    if (Test-Path $fixtureDir) {
        Remove-Item -Path $fixtureDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Path $fixtureDir -Force | Out-Null

    $orderBomPath = Join-Path $fixtureDir "order_bom.txt"
    $requestedParts = @()
    if ($spec.Contains("SourceBomPath")) {
        if (-not (Test-Path $spec.SourceBomPath)) { throw "Missing source BOM for fixture $($spec.Name): $($spec.SourceBomPath)" }
        $requestedParts = @(
            Get-Content -Path $spec.SourceBomPath |
            Where-Object { $_ -match '\w' } |
            ForEach-Object { Get-CanonicalPartKey $_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
        )
    } else {
        $parserText = Get-Content -Path $spec.SourceParserFixture -Raw
        $parsed = @(Parse-SalesOrderText -Text $parserText)
        $requestedParts = @(
            $parsed |
            ForEach-Object {
                $_.Part
                if (-not [string]::IsNullOrWhiteSpace($_.InternalPart)) { $_.InternalPart }
            } |
            ForEach-Object { Get-CanonicalPartKey $_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
        )
        $parsed | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $fixtureDir "expected_order_lines.json") -Encoding UTF8
    }
    $requestedParts | Set-Content -Path $orderBomPath -Encoding UTF8

    $requestedSet = @{}
    foreach ($part in $requestedParts) {
        $requestedSet[$part] = $true
    }

    $filterRows = {
        param([object[]]$Rows)
        @(
            $Rows |
            Where-Object {
                $base = Get-CanonicalPartKey ([string]$_.BasePart)
                $fileBase = Get-CanonicalPartKey ([string]$_.FileName)
                $parentBase = Get-AllowedChildParentKey $base
                $parentFile = Get-AllowedChildParentKey $fileBase
                $requestedSet.ContainsKey($base) -or
                $requestedSet.ContainsKey($fileBase) -or
                ((-not [string]::IsNullOrWhiteSpace($parentBase)) -and $requestedSet.ContainsKey($parentBase)) -or
                ((-not [string]::IsNullOrWhiteSpace($parentFile)) -and $requestedSet.ContainsKey($parentFile))
            }
        )
    }

    $fixturePdfRows = & $filterRows $livePdfRows
    $fixtureDxfRows = & $filterRows $liveDxfRows

    $mockRoot = Join-Path $fixtureDir "mock_store"
    $mockPdfRows = New-MockIndexedRows -Rows $fixturePdfRows -Root $mockRoot -Kind "pdfs"
    $mockDxfRows = New-MockIndexedRows -Rows $fixtureDxfRows -Root $mockRoot -Kind "dxfs"
    $mockPdfRows | Export-Csv -Path (Join-Path $fixtureDir "pdf_index_clean.csv") -NoTypeInformation -Encoding UTF8
    $mockDxfRows | Export-Csv -Path (Join-Path $fixtureDir "dxf_index_clean.csv") -NoTypeInformation -Encoding UTF8

    $fixtureOutput = Join-Path $fixtureDir "baseline_output"
    New-Item -ItemType Directory -Path $fixtureOutput -Force | Out-Null
    $fixtureConfig = New-FixtureConfig -IndexFolder $fixtureDir -OutputFolder $fixtureOutput
    $fixtureConfig | ConvertTo-Json -Depth 6 | Set-Content -Path (Join-Path $fixtureDir "config.fixture.json") -Encoding UTF8

    $fixtureMeta = [ordered]@{
        Name = $spec.Name
        JobNumber = $spec.JobNumber
        Description = $spec.Description
        RequestedParts = @($requestedParts)
        PdfIndexRows = @($mockPdfRows).Count
        DxfIndexRows = @($mockDxfRows).Count
    }
    $fixtureMeta | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $fixtureDir "fixture.json") -Encoding UTF8

    & powershell -NoProfile -ExecutionPolicy Bypass -File $replayScript `
        -BomFile $orderBomPath `
        -ConfigPath (Join-Path $fixtureDir "config.fixture.json") `
        -OutputFolder $fixtureOutput `
        -Quiet
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to build baseline summary for fixture $($spec.Name)"
    }
    Copy-Item -Path (Join-Path $fixtureOutput "collector_summary.json") -Destination (Join-Path $fixtureDir "expected_collector_summary.json") -Force
}

Write-Host "Built fixture skeletons under $FixtureRoot"
