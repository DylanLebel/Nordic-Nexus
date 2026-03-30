# SimpleCollector.ps1
# Background collector for NMT Service Hub
# Arguments: [string]BOMFile, [string]OutputFolder, [string]CollectMode

param(
    [Parameter(Mandatory=$true, Position=0)][string]$BOMFile,
    [Parameter(Mandatory=$true, Position=1)][string]$OutputFolder,
    [Parameter(Mandatory=$false, Position=2)][string]$CollectMode = "BOTH",
    [Parameter(Mandatory=$false, Position=3)][string]$ConfigPath = ""
)

$ErrorActionPreference = "Continue"

# Define index folder (should match HubService / config.json)
# We try to find it from the script directory's config.json if it exists, otherwise use default.
$scriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$cfgPathResolved = ""
if (-not [string]::IsNullOrWhiteSpace($ConfigPath)) {
    if ([System.IO.Path]::IsPathRooted($ConfigPath)) {
        $cfgPathResolved = $ConfigPath
    } else {
        $cfgPathResolved = Join-Path $scriptDir $ConfigPath
    }
}
if ([string]::IsNullOrWhiteSpace($cfgPathResolved)) {
    $cfgPathResolved = Join-Path $scriptDir "config.json"
}
$indexFolder = "C:\Users\dlebel\Documents\PDFIndex"

if (Test-Path $cfgPathResolved) {
    try {
        $cfg = Get-Content $cfgPathResolved -Raw | ConvertFrom-Json
        if ($cfg.indexFolder) { $indexFolder = $cfg.indexFolder }
    } catch { }
}

$pdfIndexClean = Join-Path $indexFolder "pdf_index_clean.csv"
$dxfIndexClean = Join-Path $indexFolder "dxf_index_clean.csv"

$summaryPath = Join-Path $env:TEMP "collector_summary.json"
$summary = @{
    pdfsFound = 0
    dxfsFound = 0
    notFound = @()
}
$copiedPdf = @{}
$copiedDxf = @{}

function Write-CollectorLog {
    param([string]$Msg)
    $ts = Get-Date -Format "HH:mm:ss"
    Write-Host "$ts $Msg"
}

function Get-FileBaseUpper {
    param([string]$NameOrPath)
    if ([string]::IsNullOrWhiteSpace($NameOrPath)) { return "" }
    try { return [System.IO.Path]::GetFileNameWithoutExtension($NameOrPath).Trim().ToUpperInvariant() } catch { return "" }
}

function Test-AllowedChildVariant {
    param(
        [string]$ParentPart,
        [string]$CandidatePart
    )
    if ([string]::IsNullOrWhiteSpace($ParentPart) -or [string]::IsNullOrWhiteSpace($CandidatePart)) { return $false }
    $parent = $ParentPart.Trim().ToUpperInvariant()
    $child = $CandidatePart.Trim().ToUpperInvariant()
    if ($child -eq $parent) { return $true }
    if (-not $child.StartsWith($parent + "-")) { return $false }

    $suffix = $child.Substring($parent.Length + 1)
    # Allow practical cutlist/hand suffixes while blocking broad variants like "-SP".
    if ($suffix -match '^(?:\d{1,3}|[LR]|\d{1,3}[A-Z]?|CL\d{1,3}|PL\d{1,3}|[A-Z]\d{1,3})$') { return $true }
    return $false
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

function Add-ToLookup {
    param(
        [hashtable]$Lookup,
        [string]$Key,
        [object]$Row
    )
    if ([string]::IsNullOrWhiteSpace($Key) -or $null -eq $Row) { return }
    $k = $Key.Trim().ToUpperInvariant()
    if (-not $Lookup.ContainsKey($k)) {
        $Lookup[$k] = New-Object System.Collections.ArrayList
    }
    [void]$Lookup[$k].Add($Row)
}

function Get-PdfIndexRowsCached {
    param([string]$PdfIndexPath)
    if (-not (Test-Path $PdfIndexPath)) { return @() }
    
    $fastCacheDir = Join-Path (Split-Path $PdfIndexPath -Parent) "FastCache"
    if (-not (Test-Path $fastCacheDir)) { New-Item -ItemType Directory -Path $fastCacheDir -Force | Out-Null }
    $cacheFile = Join-Path $fastCacheDir ("$(Split-Path $PdfIndexPath -Leaf)_cache.xml")
    
    if (Test-Path $cacheFile) {
        $csvTime = (Get-Item $PdfIndexPath).LastWriteTime
        $xmlTime = (Get-Item $cacheFile).LastWriteTime
        if ($xmlTime -gt $csvTime) {
            try {
                return Import-Clixml -Path $cacheFile
            } catch { }
        }
    }
    $rows = @(Import-Csv -Path $PdfIndexPath)
    try { $rows | Export-Clixml -Path $cacheFile } catch { }
    return $rows
}

function Copy-FileIfChanged {
    param([string]$Source, [string]$Dest)
    if (-not (Test-Path $Source)) { return $false }
    if ([string]::IsNullOrWhiteSpace($Dest)) { return $false }
    $destDir = ""
    try { $destDir = Split-Path -Path $Dest -Parent } catch { $destDir = "" }
    if (-not [string]::IsNullOrWhiteSpace($destDir) -and -not (Test-Path $destDir)) {
        try { New-Item -ItemType Directory -Path $destDir -Force | Out-Null } catch { return $false }
    }
    if (Test-Path $Dest) {
        try {
            $sInfo = Get-Item $Source
            $dInfo = Get-Item $Dest
            if ($sInfo.Length -eq $dInfo.Length -and [Math]::Abs(($sInfo.LastWriteTime - $dInfo.LastWriteTime).TotalSeconds) -lt 2) {
                return $true
            }
        } catch { }
    }
    try {
        Copy-Item -Path $Source -Destination $Dest -Force -ErrorAction Stop
        return $true
    } catch { return $false }
}

Write-CollectorLog "Starting collection..."
Write-CollectorLog "BOM: $BOMFile"
Write-CollectorLog "Output: $OutputFolder"
Write-CollectorLog "Mode: $CollectMode"

if (-not (Test-Path $BOMFile)) {
    Write-CollectorLog "ERROR: BOM file not found: $BOMFile"
    $summary | ConvertTo-Json | Set-Content -Path $summaryPath -Force
    exit 1
}

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$collectPDFs = ($CollectMode -eq "BOTH" -or $CollectMode -eq "PDF")
$collectDXFs = ($CollectMode -eq "BOTH" -or $CollectMode -eq "DXF")

# Load Indexes
$pdfIndex = @()
if ($collectPDFs -and (Test-Path $pdfIndexClean)) {
    Write-CollectorLog "Loading PDF index..."
    $pdfIndex = Get-PdfIndexRowsCached -PdfIndexPath $pdfIndexClean
}

$dxfIndex = @()
if ($collectDXFs -and (Test-Path $dxfIndexClean)) {
    Write-CollectorLog "Loading DXF index..."
    $dxfIndex = Get-PdfIndexRowsCached -PdfIndexPath $dxfIndexClean
}

$pdfByBase = @{}
$pdfByFileBase = @{}
$pdfByParentChild = @{}
$pdfByFileParentChild = @{}
if ($collectPDFs -and $pdfIndex.Count -gt 0) {
    foreach ($row in $pdfIndex) {
        $rowBaseU = if ([string]::IsNullOrWhiteSpace([string]$row.BasePart)) { "" } else { ([string]$row.BasePart).Trim().ToUpperInvariant() }
        $rowFileBaseU = Get-FileBaseUpper -NameOrPath ([string]$row.FileName)

        Add-ToLookup -Lookup $pdfByBase -Key $rowBaseU -Row $row
        Add-ToLookup -Lookup $pdfByFileBase -Key $rowFileBaseU -Row $row

        $parentBase = Get-AllowedChildParentKey -CandidatePart $rowBaseU
        if (-not [string]::IsNullOrWhiteSpace($parentBase)) {
            Add-ToLookup -Lookup $pdfByParentChild -Key $parentBase -Row $row
        }
        $parentFile = Get-AllowedChildParentKey -CandidatePart $rowFileBaseU
        if (-not [string]::IsNullOrWhiteSpace($parentFile)) {
            Add-ToLookup -Lookup $pdfByFileParentChild -Key $parentFile -Row $row
        }
    }
}

$dxfByBase = @{}
$dxfByFileBase = @{}
$dxfByParentChild = @{}
$dxfByFileParentChild = @{}
if ($collectDXFs -and $dxfIndex.Count -gt 0) {
    foreach ($row in $dxfIndex) {
        $rowBaseU = if ([string]::IsNullOrWhiteSpace([string]$row.BasePart)) { "" } else { ([string]$row.BasePart).Trim().ToUpperInvariant() }
        $rowFileBaseU = Get-FileBaseUpper -NameOrPath ([string]$row.FileName)

        Add-ToLookup -Lookup $dxfByBase -Key $rowBaseU -Row $row
        Add-ToLookup -Lookup $dxfByFileBase -Key $rowFileBaseU -Row $row

        $parentBase = Get-AllowedChildParentKey -CandidatePart $rowBaseU
        if (-not [string]::IsNullOrWhiteSpace($parentBase)) {
            Add-ToLookup -Lookup $dxfByParentChild -Key $parentBase -Row $row
        }
        $parentFile = Get-AllowedChildParentKey -CandidatePart $rowFileBaseU
        if (-not [string]::IsNullOrWhiteSpace($parentFile)) {
            Add-ToLookup -Lookup $dxfByFileParentChild -Key $parentFile -Row $row
        }
    }
}

# Read BOM
$partNumbers = Get-Content $BOMFile | Where-Object { $_ -match '\w' } | ForEach-Object { $_.Trim().ToUpper() } | Select-Object -Unique
Write-CollectorLog "Processing $($partNumbers.Count) unique parts."

$dxfSubFolder = Join-Path $OutputFolder "DXFs"
if ($collectDXFs -and -not (Test-Path $dxfSubFolder)) {
    New-Item -ItemType Directory -Path $dxfSubFolder -Force | Out-Null
}

foreach ($partNum in $partNumbers) {
    $pdfFoundForPart = $false
    $dxfFoundForPart = $false
    
    # Strip common revision suffixes for matching
    $partNumBase = $partNum -replace '_REV\w+$','' -replace '-REV\w+$','' -replace '\s+REV\w+$',''
    
    # PDF Collection
    if ($collectPDFs -and $pdfIndex.Count -gt 0) {
        $pdfMatches = New-Object System.Collections.ArrayList
        $pdfMatchSeen = @{}
        foreach ($k in @($partNum, $partNumBase) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) {
            foreach ($lookup in @($pdfByBase, $pdfByFileBase, $pdfByParentChild, $pdfByFileParentChild)) {
                if (-not $lookup.ContainsKey($k)) { continue }
                foreach ($m in @($lookup[$k])) {
                    if ($null -eq $m) { continue }
                    $mk = (([string]$m.FullPath).Trim().ToUpperInvariant() + "|" + ([string]$m.FileName).Trim().ToUpperInvariant())
                    if ($pdfMatchSeen.ContainsKey($mk)) { continue }
                    $pdfMatchSeen[$mk] = $true
                    [void]$pdfMatches.Add($m)
                }
            }
        }
        foreach ($match in $pdfMatches) {
            if (-not $match -or -not (Test-Path $match.FullPath)) { continue }
            $pdfSrcKey = ([string]$match.FullPath).Trim().ToUpperInvariant()
            if ($copiedPdf.ContainsKey($pdfSrcKey)) {
                $pdfFoundForPart = $true
                continue
            }
            $destFileName = [string]$match.FileName
            if ([string]::IsNullOrWhiteSpace($destFileName)) {
                try { $destFileName = [System.IO.Path]::GetFileName([string]$match.FullPath) } catch { $destFileName = "" }
            }
            if ([string]::IsNullOrWhiteSpace($destFileName)) {
                Write-CollectorLog "WARN: Skipping PDF match with empty filename for part '$partNum' (path='$($match.FullPath)')."
                continue
            }
            $dest = Join-Path $OutputFolder $destFileName
            if (Copy-FileIfChanged -Source $match.FullPath -Dest $dest) {
                $copiedPdf[$pdfSrcKey] = $true
                $summary.pdfsFound++
                $pdfFoundForPart = $true
            }
        }
    }
    
    # DXF Collection
    if ($collectDXFs -and $dxfIndex.Count -gt 0) {
        # Include weldment/cutlist style children for this base part (e.g. PART-01, PART-02).
        $dxfMatches = New-Object System.Collections.ArrayList
        $dxfMatchSeen = @{}
        foreach ($k in @($partNum, $partNumBase) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) {
            foreach ($lookup in @($dxfByBase, $dxfByFileBase, $dxfByParentChild, $dxfByFileParentChild)) {
                if (-not $lookup.ContainsKey($k)) { continue }
                foreach ($m in @($lookup[$k])) {
                    if ($null -eq $m) { continue }
                    $mk = (([string]$m.FullPath).Trim().ToUpperInvariant() + "|" + ([string]$m.FileName).Trim().ToUpperInvariant())
                    if ($dxfMatchSeen.ContainsKey($mk)) { continue }
                    $dxfMatchSeen[$mk] = $true
                    [void]$dxfMatches.Add($m)
                }
            }
        }
        foreach ($match in $dxfMatches) {
            if (-not $match -or -not (Test-Path $match.FullPath)) { continue }
            $dxfSrcKey = ([string]$match.FullPath).Trim().ToUpperInvariant()
            if ($copiedDxf.ContainsKey($dxfSrcKey)) {
                $dxfFoundForPart = $true
                continue
            }
            $destFileName = [string]$match.FileName
            if ([string]::IsNullOrWhiteSpace($destFileName)) {
                try { $destFileName = [System.IO.Path]::GetFileName([string]$match.FullPath) } catch { $destFileName = "" }
            }
            if ([string]::IsNullOrWhiteSpace($destFileName)) {
                Write-CollectorLog "WARN: Skipping DXF match with empty filename for part '$partNum' (path='$($match.FullPath)')."
                continue
            }
            $dest = Join-Path $dxfSubFolder $destFileName
            if (Copy-FileIfChanged -Source $match.FullPath -Dest $dest) {
                $copiedDxf[$dxfSrcKey] = $true
                $summary.dxfsFound++
                $dxfFoundForPart = $true
            }
        }
    }
    
    if (-not $pdfFoundForPart -and -not $dxfFoundForPart) {
        $summary.notFound += $partNum
    }
}

Write-CollectorLog "Done. PDFs: $($summary.pdfsFound), DXFs: $($summary.dxfsFound), Missing: $($summary.notFound.Count)"

# Write summary for EmailOrderMonitor to read
$summary | ConvertTo-Json | Set-Content -Path $summaryPath -Force
exit 0
