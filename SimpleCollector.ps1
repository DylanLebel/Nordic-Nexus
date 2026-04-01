# SimpleCollector.ps1
# Background collector for NMT Service Hub
# Arguments: [string]BOMFile, [string]OutputFolder, [string]CollectMode

param(
    [Parameter(Mandatory = $true, Position = 0)][string]$BOMFile,
    [Parameter(Mandatory = $true, Position = 1)][string]$OutputFolder,
    [Parameter(Mandatory = $false, Position = 2)][string]$CollectMode = "BOTH",
    [Parameter(Mandatory = $false, Position = 3)][string]$ConfigPath = ""
)

$ErrorActionPreference = "Continue"
$scriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
$summaryPath = Join-Path $env:TEMP "collector_summary.json"

function Resolve-ConfiguredPath {
    param(
        [string]$PathValue,
        [string]$BasePath,
        [string]$DefaultValue = ""
    )

    $candidate = if (-not [string]::IsNullOrWhiteSpace($PathValue)) { $PathValue } else { $DefaultValue }
    if ([string]::IsNullOrWhiteSpace($candidate)) { return "" }
    if ([System.IO.Path]::IsPathRooted($candidate)) { return [System.IO.Path]::GetFullPath($candidate) }
    if ([string]::IsNullOrWhiteSpace($BasePath)) { return [System.IO.Path]::GetFullPath($candidate) }
    return [System.IO.Path]::GetFullPath((Join-Path $BasePath $candidate))
}

function Write-CollectorLog {
    param([string]$Msg)
    $ts = Get-Date -Format "HH:mm:ss"
    Write-Host "$ts $Msg"
}

function Write-JsonUtf8 {
    param(
        [string]$Path,
        [object]$Value,
        [int]$Depth = 8
    )

    $parent = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $json = $Value | ConvertTo-Json -Depth $Depth
    [System.IO.File]::WriteAllText($Path, $json, [System.Text.Encoding]::UTF8)
}

function Get-FileBaseUpper {
    param([string]$NameOrPath)
    if ([string]::IsNullOrWhiteSpace($NameOrPath)) { return "" }
    try { return [System.IO.Path]::GetFileNameWithoutExtension($NameOrPath).Trim().ToUpperInvariant() } catch { return "" }
}

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

function Get-IndexRowsCached {
    param([string]$IndexPath)
    if (-not (Test-Path $IndexPath)) { return @() }

    $fastCacheDir = Join-Path (Split-Path $IndexPath -Parent) "FastCache"
    if (-not (Test-Path $fastCacheDir)) { New-Item -ItemType Directory -Path $fastCacheDir -Force | Out-Null }
    $cacheFile = Join-Path $fastCacheDir ("$(Split-Path $IndexPath -Leaf)_cache.xml")

    if (Test-Path $cacheFile) {
        $csvTime = (Get-Item $IndexPath).LastWriteTime
        $xmlTime = (Get-Item $cacheFile).LastWriteTime
        if ($xmlTime -gt $csvTime) {
            try {
                return Import-Clixml -Path $cacheFile
            } catch { }
        }
    }
    $rows = @(Import-Csv -Path $IndexPath)
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
    } catch {
        return $false
    }
}

function Test-DisallowedIndexedPath {
    param(
        [string]$FullPath,
        [string[]]$DisallowedFolderNames = @(),
        [string[]]$DisallowedNamePatterns = @()
    )

    if ([string]::IsNullOrWhiteSpace($FullPath)) { return $false }
    $segments = @($FullPath -split '[\\/]')
    foreach ($segment in $segments) {
        if ([string]::IsNullOrWhiteSpace($segment)) { continue }
        foreach ($name in $DisallowedFolderNames) {
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            if ($segment.Trim().ToUpperInvariant() -eq $name.Trim().ToUpperInvariant()) { return $true }
        }
    }
    $leaf = ""
    try { $leaf = [System.IO.Path]::GetFileNameWithoutExtension($FullPath) } catch { $leaf = "" }
    foreach ($pattern in $DisallowedNamePatterns) {
        if ([string]::IsNullOrWhiteSpace($pattern)) { continue }
        if ($leaf -like $pattern) { return $true }
    }
    return $false
}

function Convert-IndexRow {
    param([object]$Row)
    if ($null -eq $Row) { return $null }

    $fullPath = [string]$Row.FullPath
    $fileName = [string]$Row.FileName
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        try { $fileName = [System.IO.Path]::GetFileName($fullPath) } catch { $fileName = "" }
    }

    $basePart = ([string]$Row.BasePart).Trim().ToUpperInvariant()
    $canonicalBasePart = Get-CanonicalPartKey $basePart
    $canonicalFileBase = Get-CanonicalPartKey $fileName
    $rev = ([string]$Row.Rev).Trim().ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($rev) -and $fileName -match '(?i)[ _-]REV[ _-]*(?<rev>[A-Z0-9]+)\.') {
        $rev = $Matches['rev'].Trim().ToUpperInvariant()
    }
    $lastWriteText = [string]$Row.LastWriteTime
    $lastWriteValue = $null
    if (-not [string]::IsNullOrWhiteSpace($lastWriteText)) {
        try { $lastWriteValue = [datetime]$lastWriteText } catch { $lastWriteValue = $null }
    }

    return [pscustomobject]@{
        BasePart          = $basePart
        CanonicalBasePart = $canonicalBasePart
        CanonicalFileBase = $canonicalFileBase
        FileName          = $fileName
        FullPath          = $fullPath
        Rev               = $rev
        FileType          = [string]$Row.FileType
        IsObsolete        = [string]$Row.IsObsolete
        RootFolder        = [string]$Row.RootFolder
        LastWriteTime     = $lastWriteText
        LastWriteTimeValue = $lastWriteValue
        SourceIndex       = "clean"
    }
}

function Convert-RawIndexRow {
    param(
        [object]$Row,
        [string]$DefaultFileType = ""
    )
    if ($null -eq $Row) { return $null }

    $fullPath = [string]$Row.FullPath
    $fileName = [string]$Row.FileName
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        try { $fileName = [System.IO.Path]::GetFileName($fullPath) } catch { $fileName = "" }
    }

    $baseName = ([string]$Row.BaseName).Trim().ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($baseName)) {
        $baseName = Get-FileBaseUpper $fileName
    }
    $canonicalBasePart = Get-CanonicalPartKey $baseName
    $canonicalFileBase = Get-CanonicalPartKey $fileName
    $rev = ""
    foreach ($candidate in @($baseName, $fileName)) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        if ($candidate -match '(?i)[ _-]REV[ _-]*(?<rev>[A-Z0-9]+)(?:\.|$)') {
            $rev = $Matches['rev'].Trim().ToUpperInvariant()
            break
        }
    }
    $lastWriteText = [string]$Row.LastWriteTime
    $lastWriteValue = $null
    if (-not [string]::IsNullOrWhiteSpace($lastWriteText)) {
        try { $lastWriteValue = [datetime]$lastWriteText } catch { $lastWriteValue = $null }
    }

    return [pscustomobject]@{
        BasePart          = $canonicalBasePart
        CanonicalBasePart = $canonicalBasePart
        CanonicalFileBase = $canonicalFileBase
        FileName          = $fileName
        FullPath          = $fullPath
        Rev               = $rev
        FileType          = $DefaultFileType
        IsObsolete        = ""
        RootFolder        = [string]$Row.RootFolder
        LastWriteTime     = $lastWriteText
        LastWriteTimeValue = $lastWriteValue
        SourceIndex       = "raw"
        BaseNameRaw       = $baseName
    }
}

function New-MatchRecord {
    param(
        [string]$RequestedPart,
        [string]$Strategy,
        [string]$Kind,
        [object]$Row
    )
    if ($null -eq $Row) { return $null }
    $directory = ""
    try { $directory = Split-Path -Path ([string]$Row.FullPath) -Parent } catch { $directory = "" }
    return [ordered]@{
        RequestedPart      = $RequestedPart
        Kind               = $Kind
        Strategy           = $Strategy
        BasePart           = [string]$Row.BasePart
        CanonicalBasePart  = [string]$Row.CanonicalBasePart
        CanonicalFileBase  = [string]$Row.CanonicalFileBase
        FileName           = [string]$Row.FileName
        FullPath           = [string]$Row.FullPath
        Directory          = $directory
        Rev                = [string]$Row.Rev
    }
}

function Get-IndexMetadata {
    param(
        [string]$Path,
        [double]$StaleWarningHours
    )

    $exists = Test-Path $Path
    $lastWrite = $null
    $ageHours = $null
    if ($exists) {
        try {
            $lastWrite = (Get-Item $Path).LastWriteTime
            $ageHours = [math]::Round(((Get-Date) - $lastWrite).TotalHours, 2)
        } catch { }
    }
    $stale = ($exists -and $null -ne $ageHours -and $ageHours -gt $StaleWarningHours)
    return [ordered]@{
        Path          = $Path
        Exists        = $exists
        LastWriteTime = $(if ($null -ne $lastWrite) { $lastWrite.ToString("yyyy-MM-dd HH:mm:ss") } else { "" })
        AgeHours      = $ageHours
        IsStale       = $stale
    }
}

function Add-MatchSet {
    param(
        [System.Collections.ArrayList]$Target,
        [hashtable]$Seen,
        [object[]]$Rows = @(),
        [string]$RequestedPart,
        [string]$Strategy,
        [string]$Kind
    )

    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $key = (([string]$row.FullPath).Trim().ToUpperInvariant() + "|" + ([string]$row.FileName).Trim().ToUpperInvariant())
        if ($Seen.ContainsKey($key)) { continue }
        $Seen[$key] = $true
        [void]$Target.Add((New-MatchRecord -RequestedPart $RequestedPart -Strategy $Strategy -Kind $Kind -Row $row))
    }
}

function Add-RowSet {
    param(
        [System.Collections.ArrayList]$Target,
        [hashtable]$Seen,
        [object[]]$Rows = @()
    )

    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $key = ([string]$row.FullPath).Trim().ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        if ($Seen.ContainsKey($key)) { continue }
        $Seen[$key] = $true
        [void]$Target.Add($row)
    }
}

function Test-PathUnderRoots {
    param(
        [string]$Path,
        [string[]]$Roots = @()
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or @($Roots).Count -eq 0) { return $false }

    $pathNorm = ""
    try { $pathNorm = [System.IO.Path]::GetFullPath($Path).TrimEnd('\').ToUpperInvariant() } catch { $pathNorm = $Path.Trim().TrimEnd('\').ToUpperInvariant() }
    if ([string]::IsNullOrWhiteSpace($pathNorm)) { return $false }

    foreach ($root in @($Roots)) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        $rootNorm = ""
        try { $rootNorm = [System.IO.Path]::GetFullPath($root).TrimEnd('\').ToUpperInvariant() } catch { $rootNorm = $root.Trim().TrimEnd('\').ToUpperInvariant() }
        if ([string]::IsNullOrWhiteSpace($rootNorm)) { continue }
        if ($pathNorm -eq $rootNorm -or $pathNorm.StartsWith($rootNorm + "\")) { return $true }
    }
    return $false
}

function Get-DistinctMatchRevisions {
    param([object[]]$Matches = @())

    $vals = @(
        @($Matches) |
        ForEach-Object { ([string]$_.Rev).Trim().ToUpperInvariant() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
    )
    return $vals
}

function Get-ExactLookupRows {
    param(
        [hashtable]$ByCanonicalBase,
        [hashtable]$ByCanonicalFileBase,
        [string[]]$RequestedKeys = @()
    )

    $rows = New-Object System.Collections.ArrayList
    $seen = @{}
    foreach ($k in @($RequestedKeys | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        Add-RowSet -Target $rows -Seen $seen -Rows @($ByCanonicalBase[$k])
        Add-RowSet -Target $rows -Seen $seen -Rows @($ByCanonicalFileBase[$k])
    }
    return @($rows)
}

function Select-PreferredLiveRows {
    param(
        [object[]]$Rows = @(),
        [string[]]$DesiredRevs = @(),
        [string[]]$ProtectedRoots = @(),
        [bool]$AllowMultipleWithoutRev = $false,
        [bool]$RequireProtectedRoot = $false
    )

    $candidates = @(
        @($Rows) |
        Where-Object {
            $null -ne $_ -and
            -not [string]::IsNullOrWhiteSpace($_.FullPath) -and
            (Test-Path ([string]$_.FullPath))
        }
    )
    if (@($candidates).Count -eq 0) { return @() }

    $desiredRevSet = @{}
    foreach ($rev in @($DesiredRevs)) {
        $revKey = ([string]$rev).Trim().ToUpperInvariant()
        if (-not [string]::IsNullOrWhiteSpace($revKey)) { $desiredRevSet[$revKey] = $true }
    }
    if ($desiredRevSet.Count -gt 0) {
        $candidates = @(
            @($candidates) |
            Where-Object {
                $revKey = ([string]$_.Rev).Trim().ToUpperInvariant()
                $desiredRevSet.ContainsKey($revKey)
            }
        )
        if (@($candidates).Count -eq 0) { return @() }
    }

    $protected = @()
    if (@($ProtectedRoots).Count -gt 0) {
        $protected = @(
            @($candidates) |
            Where-Object { Test-PathUnderRoots -Path ([string]$_.FullPath) -Roots $ProtectedRoots }
        )
    }
    if (@($protected).Count -gt 0) {
        $candidates = $protected
    } elseif ($RequireProtectedRoot -and @($ProtectedRoots).Count -gt 0) {
        return @()
    }

    $candidates = @(
        @($candidates) |
        Sort-Object `
            @{Expression={ if ($_.LastWriteTimeValue) { [datetime]$_.LastWriteTimeValue } else { [datetime]::MinValue } };Descending=$true},`
            @{Expression={ [string]$_.FileName };Descending=$false},`
            @{Expression={ [string]$_.FullPath };Descending=$false}
    )

    if (-not $AllowMultipleWithoutRev -and $desiredRevSet.Count -eq 0 -and @($candidates).Count -gt 1) {
        return @()
    }

    return $candidates
}

function Test-ModelEntryExtension {
    param(
        [object]$Row,
        [string]$ExpectedExtension
    )
    if ($null -eq $Row -or [string]::IsNullOrWhiteSpace($ExpectedExtension)) { return $false }
    $fileName = [string]$Row.FileName
    if ([string]::IsNullOrWhiteSpace($fileName)) {
        $fileName = [string]$Row.FullPath
    }
    if ([string]::IsNullOrWhiteSpace($fileName)) { return $false }
    return $fileName.Trim().ToUpperInvariant().EndsWith($ExpectedExtension.Trim().ToUpperInvariant())
}

function Copy-MatchRecordsForPart {
    param(
        [object[]]$Matches = @(),
        [hashtable]$AlreadyCopied,
        [string]$DestinationFolder,
        [System.Collections.ArrayList]$CollectedFiles,
        [System.Collections.IDictionary]$Summary,
        [string]$RequestedPart,
        [string]$Kind
    )

    $foundForPart = $false
    foreach ($match in @($Matches)) {
        if ($null -eq $match -or -not (Test-Path ([string]$match.FullPath))) { continue }
        $srcKey = ([string]$match.FullPath).Trim().ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($srcKey)) { continue }
        if ($AlreadyCopied.ContainsKey($srcKey)) {
            $foundForPart = $true
            continue
        }
        $destFileName = [string]$match.FileName
        if ([string]::IsNullOrWhiteSpace($destFileName)) {
            try { $destFileName = [System.IO.Path]::GetFileName([string]$match.FullPath) } catch { $destFileName = "" }
        }
        if ([string]::IsNullOrWhiteSpace($destFileName)) { continue }
        $dest = Join-Path $DestinationFolder $destFileName
        if (-not (Copy-FileIfChanged -Source ([string]$match.FullPath) -Dest $dest)) { continue }

        $AlreadyCopied[$srcKey] = $true
        $foundForPart = $true
        $entry = [ordered]@{
            RequestedPart = $RequestedPart
            FileName      = $destFileName
            SourcePath    = [string]$match.FullPath
            DestPath      = $dest
            Strategy      = [string]$match.Strategy
            Rev           = [string]$match.Rev
        }
        [void]$CollectedFiles.Add($entry)
        if ($Kind -eq "PDF") {
            $Summary["pdfsFound"] = [int]$Summary["pdfsFound"] + 1
            $Summary["CollectedPdfFiles"] = @($Summary["CollectedPdfFiles"]) + @($entry)
        } else {
            $Summary["dxfsFound"] = [int]$Summary["dxfsFound"] + 1
            $Summary["CollectedDxfFiles"] = @($Summary["CollectedDxfFiles"]) + @($entry)
        }
    }
    return [bool]$foundForPart
}

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

$cfg = $null
if (Test-Path $cfgPathResolved) {
    try { $cfg = Get-Content $cfgPathResolved -Raw | ConvertFrom-Json } catch { $cfg = $null }
}

$configBaseDir = if ([string]::IsNullOrWhiteSpace($cfgPathResolved)) { $scriptDir } else { Split-Path $cfgPathResolved -Parent }
$indexFolder = Resolve-ConfiguredPath -PathValue $cfg.indexFolder -BasePath $configBaseDir -DefaultValue "C:\Users\dlebel\Documents\PDFIndex"
$pdfIndexClean = Join-Path $indexFolder "pdf_index_clean.csv"
$dxfIndexClean = Join-Path $indexFolder "dxf_index_clean.csv"
$pdfIndexRaw = Join-Path $indexFolder "pdf_index_raw.csv"
$dxfIndexRaw = Join-Path $indexFolder "dxf_index_raw.csv"
$modelIndexClean = Join-Path $indexFolder "model_index_clean.csv"

$collectorCfg = if ($cfg -and $cfg.collector) { $cfg.collector } else { $null }
$strictMatching = $true
if ($collectorCfg -and $null -ne $collectorCfg.strictMatching) {
    try { $strictMatching = [bool]$collectorCfg.strictMatching } catch { $strictMatching = $true }
}
$copyChildVariants = $false
if ($collectorCfg -and $null -ne $collectorCfg.copyChildVariants) {
    try { $copyChildVariants = [bool]$collectorCfg.copyChildVariants } catch { $copyChildVariants = $false }
}
$indexStaleWarningHours = 24.0
if ($collectorCfg -and $collectorCfg.indexStaleWarningHours) {
    try { $indexStaleWarningHours = [double]$collectorCfg.indexStaleWarningHours } catch { $indexStaleWarningHours = 24.0 }
}
$enableLiveExactRawFallback = $false
if ($collectorCfg -and $null -ne $collectorCfg.enableLiveExactRawFallback) {
    try { $enableLiveExactRawFallback = [bool]$collectorCfg.enableLiveExactRawFallback } catch { $enableLiveExactRawFallback = $false }
}
$enableAssemblyParentPdfFallback = $false
if ($collectorCfg -and $null -ne $collectorCfg.enableAssemblyParentPdfFallback) {
    try { $enableAssemblyParentPdfFallback = [bool]$collectorCfg.enableAssemblyParentPdfFallback } catch { $enableAssemblyParentPdfFallback = $false }
}
$disallowedFolderNames = @()
$disallowedNamePatterns = @()
if ($cfg -and $cfg.pathFilters) {
    if ($cfg.pathFilters.disallowedModelFolderNames) { $disallowedFolderNames = @($cfg.pathFilters.disallowedModelFolderNames) }
    if ($cfg.pathFilters.disallowedModelNamePatterns) { $disallowedNamePatterns = @($cfg.pathFilters.disallowedModelNamePatterns) }
}
$protectedRoots = @()
if ($cfg -and $cfg.protectedRoots) {
    $protectedRoots = @(
        @($cfg.protectedRoots) |
        ForEach-Object { Resolve-ConfiguredPath -PathValue ([string]$_) -BasePath $configBaseDir } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
    )
}
$copyExtraMatches = ((-not $strictMatching) -or $copyChildVariants)

$summary = [ordered]@{
    Timestamp        = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    BomFile          = $BOMFile
    OutputFolder     = $OutputFolder
    PdfOutputFolder  = $OutputFolder
    DxfOutputFolder  = (Join-Path $OutputFolder "DXFs")
    CollectMode      = $CollectMode
    ConfigPath       = $cfgPathResolved
    StrictMatching   = [bool]$strictMatching
    CopyChildVariants = [bool]$copyChildVariants
    EnableLiveExactRawFallback = [bool]$enableLiveExactRawFallback
    EnableAssemblyParentPdfFallback = [bool]$enableAssemblyParentPdfFallback
    RequestedParts   = @()
    pdfsFound        = 0
    dxfsFound        = 0
    notFound         = @()
    extraPdfs        = @()
    extraDxfs        = @()
    CollectedPdfFiles = @()
    CollectedDxfFiles = @()
    partResults      = @()
    warnings         = @()
    indexInfo        = [ordered]@{
        PdfIndex = Get-IndexMetadata -Path $pdfIndexClean -StaleWarningHours $indexStaleWarningHours
        DxfIndex = Get-IndexMetadata -Path $dxfIndexClean -StaleWarningHours $indexStaleWarningHours
        PdfRawIndex = $(if ($enableLiveExactRawFallback -or $enableAssemblyParentPdfFallback) { Get-IndexMetadata -Path $pdfIndexRaw -StaleWarningHours $indexStaleWarningHours } else { $null })
        DxfRawIndex = $(if ($enableLiveExactRawFallback) { Get-IndexMetadata -Path $dxfIndexRaw -StaleWarningHours $indexStaleWarningHours } else { $null })
        ModelIndex = $(if ($enableAssemblyParentPdfFallback) { Get-IndexMetadata -Path $modelIndexClean -StaleWarningHours $indexStaleWarningHours } else { $null })
        WarningHours = $indexStaleWarningHours
        IsStale = $false
    }
}

Write-CollectorLog "Starting collection..."
Write-CollectorLog "BOM: $BOMFile"
Write-CollectorLog "Output: $OutputFolder"
Write-CollectorLog "Mode: $CollectMode"
Write-CollectorLog "Strict matching: $strictMatching | Copy child variants: $copyChildVariants"
Write-CollectorLog "Live exact fallback: $enableLiveExactRawFallback | Assembly parent fallback: $enableAssemblyParentPdfFallback"

if (-not (Test-Path $BOMFile)) {
    Write-CollectorLog "ERROR: BOM file not found: $BOMFile"
    Write-JsonUtf8 -Path $summaryPath -Value $summary -Depth 8
    exit 1
}

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$collectPDFs = ($CollectMode -eq "BOTH" -or $CollectMode -eq "PDF")
$collectDXFs = ($CollectMode -eq "BOTH" -or $CollectMode -eq "DXF")

if (-not $summary.indexInfo.PdfIndex.Exists -and $collectPDFs) {
    $summary.warnings += "PDF index not found: $pdfIndexClean"
}
if (-not $summary.indexInfo.DxfIndex.Exists -and $collectDXFs) {
    $summary.warnings += "DXF index not found: $dxfIndexClean"
}
if ($summary.indexInfo.PdfIndex.IsStale) {
    $summary.warnings += "PDF index is stale ($($summary.indexInfo.PdfIndex.AgeHours) hours old)"
}
if ($summary.indexInfo.DxfIndex.IsStale) {
    $summary.warnings += "DXF index is stale ($($summary.indexInfo.DxfIndex.AgeHours) hours old)"
}
if ($enableLiveExactRawFallback -and $collectPDFs -and $summary.indexInfo.PdfRawIndex -and -not $summary.indexInfo.PdfRawIndex.Exists) {
    $summary.warnings += "PDF raw index not found: $pdfIndexRaw"
}
if ($enableLiveExactRawFallback -and $collectDXFs -and $summary.indexInfo.DxfRawIndex -and -not $summary.indexInfo.DxfRawIndex.Exists) {
    $summary.warnings += "DXF raw index not found: $dxfIndexRaw"
}
if ($enableAssemblyParentPdfFallback -and $summary.indexInfo.ModelIndex -and -not $summary.indexInfo.ModelIndex.Exists) {
    $summary.warnings += "Model index not found: $modelIndexClean"
}
$summary.indexInfo.IsStale = [bool]($summary.indexInfo.PdfIndex.IsStale -or $summary.indexInfo.DxfIndex.IsStale)

$pdfIndex = @()
if ($collectPDFs -and (Test-Path $pdfIndexClean)) {
    Write-CollectorLog "Loading PDF index..."
    $pdfIndex = @(
        Get-IndexRowsCached -IndexPath $pdfIndexClean |
        ForEach-Object { Convert-IndexRow -Row $_ } |
        Where-Object {
            $null -ne $_ -and
            -not [string]::IsNullOrWhiteSpace($_.FullPath) -and
            -not (Test-DisallowedIndexedPath -FullPath $_.FullPath -DisallowedFolderNames $disallowedFolderNames -DisallowedNamePatterns $disallowedNamePatterns)
        }
    )
}

$dxfIndex = @()
if ($collectDXFs -and (Test-Path $dxfIndexClean)) {
    Write-CollectorLog "Loading DXF index..."
    $dxfIndex = @(
        Get-IndexRowsCached -IndexPath $dxfIndexClean |
        ForEach-Object { Convert-IndexRow -Row $_ } |
        Where-Object {
            $null -ne $_ -and
            -not [string]::IsNullOrWhiteSpace($_.FullPath) -and
            -not (Test-DisallowedIndexedPath -FullPath $_.FullPath -DisallowedFolderNames $disallowedFolderNames -DisallowedNamePatterns $disallowedNamePatterns)
        }
    )
}

$pdfRawIndexRows = @()
if (($enableLiveExactRawFallback -or $enableAssemblyParentPdfFallback) -and $collectPDFs -and (Test-Path $pdfIndexRaw)) {
    Write-CollectorLog "Loading PDF raw index fallback data..."
    $pdfRawIndexRows = @(
        Get-IndexRowsCached -IndexPath $pdfIndexRaw |
        ForEach-Object { Convert-RawIndexRow -Row $_ -DefaultFileType "PDF" } |
        Where-Object {
            $null -ne $_ -and
            -not [string]::IsNullOrWhiteSpace($_.FullPath) -and
            -not (Test-DisallowedIndexedPath -FullPath $_.FullPath -DisallowedFolderNames $disallowedFolderNames -DisallowedNamePatterns $disallowedNamePatterns)
        }
    )
}

$dxfRawIndexRows = @()
if ($enableLiveExactRawFallback -and $collectDXFs -and (Test-Path $dxfIndexRaw)) {
    Write-CollectorLog "Loading DXF raw index fallback data..."
    $dxfRawIndexRows = @(
        Get-IndexRowsCached -IndexPath $dxfIndexRaw |
        ForEach-Object { Convert-RawIndexRow -Row $_ -DefaultFileType "DXF" } |
        Where-Object {
            $null -ne $_ -and
            -not [string]::IsNullOrWhiteSpace($_.FullPath) -and
            -not (Test-DisallowedIndexedPath -FullPath $_.FullPath -DisallowedFolderNames $disallowedFolderNames -DisallowedNamePatterns $disallowedNamePatterns)
        }
    )
}

$modelIndexRows = @()
if ($enableAssemblyParentPdfFallback -and (Test-Path $modelIndexClean)) {
    Write-CollectorLog "Loading model index for assembly fallback..."
    $modelIndexRows = @(
        Get-IndexRowsCached -IndexPath $modelIndexClean |
        Where-Object {
            $null -ne $_ -and
            -not [string]::IsNullOrWhiteSpace([string]$_.BasePart) -and
            -not [string]::IsNullOrWhiteSpace([string]$_.FullPath)
        }
    )
}

$pdfByCanonicalBase = @{}
$pdfByCanonicalFileBase = @{}
$pdfByParentChildBase = @{}
$pdfByParentChildFileBase = @{}
foreach ($row in @($pdfIndex)) {
    Add-ToLookup -Lookup $pdfByCanonicalBase -Key ([string]$row.CanonicalBasePart) -Row $row
    Add-ToLookup -Lookup $pdfByCanonicalFileBase -Key ([string]$row.CanonicalFileBase) -Row $row
    $parentBase = Get-AllowedChildParentKey -CandidatePart ([string]$row.CanonicalBasePart)
    if (-not [string]::IsNullOrWhiteSpace($parentBase)) { Add-ToLookup -Lookup $pdfByParentChildBase -Key $parentBase -Row $row }
    $parentFile = Get-AllowedChildParentKey -CandidatePart ([string]$row.CanonicalFileBase)
    if (-not [string]::IsNullOrWhiteSpace($parentFile)) { Add-ToLookup -Lookup $pdfByParentChildFileBase -Key $parentFile -Row $row }
}

$dxfByCanonicalBase = @{}
$dxfByCanonicalFileBase = @{}
$dxfByParentChildBase = @{}
$dxfByParentChildFileBase = @{}
foreach ($row in @($dxfIndex)) {
    Add-ToLookup -Lookup $dxfByCanonicalBase -Key ([string]$row.CanonicalBasePart) -Row $row
    Add-ToLookup -Lookup $dxfByCanonicalFileBase -Key ([string]$row.CanonicalFileBase) -Row $row
    $parentBase = Get-AllowedChildParentKey -CandidatePart ([string]$row.CanonicalBasePart)
    if (-not [string]::IsNullOrWhiteSpace($parentBase)) { Add-ToLookup -Lookup $dxfByParentChildBase -Key $parentBase -Row $row }
    $parentFile = Get-AllowedChildParentKey -CandidatePart ([string]$row.CanonicalFileBase)
    if (-not [string]::IsNullOrWhiteSpace($parentFile)) { Add-ToLookup -Lookup $dxfByParentChildFileBase -Key $parentFile -Row $row }
}

$pdfRawByCanonicalBase = @{}
$pdfRawByCanonicalFileBase = @{}
foreach ($row in @($pdfRawIndexRows)) {
    Add-ToLookup -Lookup $pdfRawByCanonicalBase -Key ([string]$row.CanonicalBasePart) -Row $row
    Add-ToLookup -Lookup $pdfRawByCanonicalFileBase -Key ([string]$row.CanonicalFileBase) -Row $row
}

$dxfRawByCanonicalBase = @{}
$dxfRawByCanonicalFileBase = @{}
foreach ($row in @($dxfRawIndexRows)) {
    Add-ToLookup -Lookup $dxfRawByCanonicalBase -Key ([string]$row.CanonicalBasePart) -Row $row
    Add-ToLookup -Lookup $dxfRawByCanonicalFileBase -Key ([string]$row.CanonicalFileBase) -Row $row
}

$modelByBasePart = @{}
foreach ($row in @($modelIndexRows)) {
    $basePartKey = ([string]$row.BasePart).Trim().ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($basePartKey)) { continue }
    Add-ToLookup -Lookup $modelByBasePart -Key $basePartKey -Row $row
}

$partNumbers = @(
    Get-Content $BOMFile |
    Where-Object { $_ -match '\w' } |
    ForEach-Object { Get-CanonicalPartKey $_ } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Select-Object -Unique
)
$summary.RequestedParts = @($partNumbers)
Write-CollectorLog "Processing $($partNumbers.Count) unique parts."

$dxfSubFolder = Join-Path $OutputFolder "DXFs"
if ($collectDXFs -and -not (Test-Path $dxfSubFolder)) {
    New-Item -ItemType Directory -Path $dxfSubFolder -Force | Out-Null
}

$copiedPdf = @{}
$copiedDxf = @{}

foreach ($partNum in $partNumbers) {
    $partNumBase = Get-CanonicalPartKey $partNum
    $requestedKeys = @($partNum, $partNumBase) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    $primaryPdfMatches = New-Object System.Collections.ArrayList
    $extraPdfMatches = New-Object System.Collections.ArrayList
    $pdfMatchSeen = @{}
    foreach ($k in $requestedKeys) {
        Add-MatchSet -Target $primaryPdfMatches -Seen $pdfMatchSeen -Rows @($pdfByCanonicalBase[$k]) -RequestedPart $partNum -Strategy "exact-base" -Kind "PDF"
        Add-MatchSet -Target $primaryPdfMatches -Seen $pdfMatchSeen -Rows @($pdfByCanonicalFileBase[$k]) -RequestedPart $partNum -Strategy "exact-file-base" -Kind "PDF"
        Add-MatchSet -Target $extraPdfMatches -Seen $pdfMatchSeen -Rows @($pdfByParentChildBase[$k]) -RequestedPart $partNum -Strategy "parent-child-base" -Kind "PDF"
        Add-MatchSet -Target $extraPdfMatches -Seen $pdfMatchSeen -Rows @($pdfByParentChildFileBase[$k]) -RequestedPart $partNum -Strategy "parent-child-file-base" -Kind "PDF"
    }

    $primaryDxfMatches = New-Object System.Collections.ArrayList
    $extraDxfMatches = New-Object System.Collections.ArrayList
    $dxfMatchSeen = @{}
    foreach ($k in $requestedKeys) {
        Add-MatchSet -Target $primaryDxfMatches -Seen $dxfMatchSeen -Rows @($dxfByCanonicalBase[$k]) -RequestedPart $partNum -Strategy "exact-base" -Kind "DXF"
        Add-MatchSet -Target $primaryDxfMatches -Seen $dxfMatchSeen -Rows @($dxfByCanonicalFileBase[$k]) -RequestedPart $partNum -Strategy "exact-file-base" -Kind "DXF"
        Add-MatchSet -Target $extraDxfMatches -Seen $dxfMatchSeen -Rows @($dxfByParentChildBase[$k]) -RequestedPart $partNum -Strategy "parent-child-base" -Kind "DXF"
        Add-MatchSet -Target $extraDxfMatches -Seen $dxfMatchSeen -Rows @($dxfByParentChildFileBase[$k]) -RequestedPart $partNum -Strategy "parent-child-file-base" -Kind "DXF"
    }

    $copiedPdfFiles = New-Object System.Collections.ArrayList
    $copiedDxfFiles = New-Object System.Collections.ArrayList
    $fallbackPdfMatches = New-Object System.Collections.ArrayList
    $fallbackDxfMatches = New-Object System.Collections.ArrayList
    $fallbackNotes = New-Object System.Collections.ArrayList
    $pdfFoundForPart = $false
    $dxfFoundForPart = $false
    $pdfFallbackSeen = @{}
    $dxfFallbackSeen = @{}

    $pdfMatchesToCopy = @($primaryPdfMatches)
    if ($copyExtraMatches) { $pdfMatchesToCopy += @($extraPdfMatches) }
    $pdfFoundForPart = Copy-MatchRecordsForPart -Matches @($pdfMatchesToCopy) -AlreadyCopied $copiedPdf -DestinationFolder $OutputFolder -CollectedFiles $copiedPdfFiles -Summary $summary -RequestedPart $partNum -Kind "PDF"

    $dxfMatchesToCopy = @($primaryDxfMatches)
    if ($copyExtraMatches) { $dxfMatchesToCopy += @($extraDxfMatches) }
    $dxfFoundForPart = Copy-MatchRecordsForPart -Matches @($dxfMatchesToCopy) -AlreadyCopied $copiedDxf -DestinationFolder $dxfSubFolder -CollectedFiles $copiedDxfFiles -Summary $summary -RequestedPart $partNum -Kind "DXF"

    if (-not $pdfFoundForPart -and $enableLiveExactRawFallback -and $collectPDFs) {
        $desiredPdfRevs = Get-DistinctMatchRevisions -Matches @($primaryPdfMatches)
        $rawPdfExactRows = Get-ExactLookupRows -ByCanonicalBase $pdfRawByCanonicalBase -ByCanonicalFileBase $pdfRawByCanonicalFileBase -RequestedKeys $requestedKeys
        $selectedRawPdfRows = Select-PreferredLiveRows -Rows @($rawPdfExactRows) -DesiredRevs @($desiredPdfRevs) -ProtectedRoots $protectedRoots -AllowMultipleWithoutRev $false
        if (@($selectedRawPdfRows).Count -gt 0) {
            Add-MatchSet -Target $fallbackPdfMatches -Seen $pdfFallbackSeen -Rows @($selectedRawPdfRows | Select-Object -First 1) -RequestedPart $partNum -Strategy "live-exact-raw-fallback" -Kind "PDF"
            if (Copy-MatchRecordsForPart -Matches @($fallbackPdfMatches) -AlreadyCopied $copiedPdf -DestinationFolder $OutputFolder -CollectedFiles $copiedPdfFiles -Summary $summary -RequestedPart $partNum -Kind "PDF") {
                $pdfFoundForPart = $true
                [void]$fallbackNotes.Add("Used live exact PDF fallback.")
            }
        }
    }

    if (-not $dxfFoundForPart -and $enableLiveExactRawFallback -and $collectDXFs) {
        $desiredDxfRevs = Get-DistinctMatchRevisions -Matches @($primaryDxfMatches)
        $rawDxfExactRows = Get-ExactLookupRows -ByCanonicalBase $dxfRawByCanonicalBase -ByCanonicalFileBase $dxfRawByCanonicalFileBase -RequestedKeys $requestedKeys
        $selectedRawDxfRows = Select-PreferredLiveRows -Rows @($rawDxfExactRows) -DesiredRevs @($desiredDxfRevs) -ProtectedRoots $protectedRoots -AllowMultipleWithoutRev $false
        if (@($selectedRawDxfRows).Count -gt 0) {
            Add-MatchSet -Target $fallbackDxfMatches -Seen $dxfFallbackSeen -Rows @($selectedRawDxfRows | Select-Object -First 1) -RequestedPart $partNum -Strategy "live-exact-raw-fallback" -Kind "DXF"
            if (Copy-MatchRecordsForPart -Matches @($fallbackDxfMatches) -AlreadyCopied $copiedDxf -DestinationFolder $dxfSubFolder -CollectedFiles $copiedDxfFiles -Summary $summary -RequestedPart $partNum -Kind "DXF") {
                $dxfFoundForPart = $true
                [void]$fallbackNotes.Add("Used live exact DXF fallback.")
            }
        }
    }

    if (-not $pdfFoundForPart -and $enableAssemblyParentPdfFallback -and $collectPDFs) {
        $parentPart = Get-AllowedChildParentKey -CandidatePart $partNumBase
        if (-not [string]::IsNullOrWhiteSpace($parentPart)) {
            $parentKeys = @($parentPart) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
            $selectedParentRows = @()
            $parentSelectionSource = ""
            $selectedParentCleanRows = Select-PreferredLiveRows -Rows @(Get-ExactLookupRows -ByCanonicalBase $pdfByCanonicalBase -ByCanonicalFileBase $pdfByCanonicalFileBase -RequestedKeys $parentKeys) -DesiredRevs @() -ProtectedRoots $protectedRoots -AllowMultipleWithoutRev $true -RequireProtectedRoot $true
            if (@($selectedParentCleanRows).Count -gt 0) {
                $selectedParentRows = @($selectedParentCleanRows)
                $parentSelectionSource = "clean"
            } else {
                $selectedParentRawRows = Select-PreferredLiveRows -Rows @(Get-ExactLookupRows -ByCanonicalBase $pdfRawByCanonicalBase -ByCanonicalFileBase $pdfRawByCanonicalFileBase -RequestedKeys $parentKeys) -DesiredRevs @() -ProtectedRoots $protectedRoots -AllowMultipleWithoutRev $true -RequireProtectedRoot $true
                if (@($selectedParentRawRows).Count -gt 0) {
                    $selectedParentRows = @($selectedParentRawRows)
                    $parentSelectionSource = "raw"
                }
            }
            if (@($selectedParentRows).Count -gt 0) {
                Add-MatchSet -Target $fallbackPdfMatches -Seen $pdfFallbackSeen -Rows @($selectedParentRows | Select-Object -First 1) -RequestedPart $partNum -Strategy "assembly-parent-pdf-fallback" -Kind "PDF"
                if (Copy-MatchRecordsForPart -Matches @($fallbackPdfMatches) -AlreadyCopied $copiedPdf -DestinationFolder $OutputFolder -CollectedFiles $copiedPdfFiles -Summary $summary -RequestedPart $partNum -Kind "PDF") {
                    $pdfFoundForPart = $true
                    $note = "Used assembly parent PDF fallback via $parentPart"
                    if (-not [string]::IsNullOrWhiteSpace($parentSelectionSource)) {
                        $note += " ($parentSelectionSource index"
                        if (@($selectedParentRows).Count -gt 1) {
                            $note += ", newest of $(@($selectedParentRows).Count) live candidates"
                        }
                        $note += ")."
                    } else {
                        $note += "."
                    }
                    [void]$fallbackNotes.Add($note)
                }
            }
        }
    }

    if (-not $pdfFoundForPart -and -not $dxfFoundForPart) {
        $summary.notFound += $partNum
    }

    foreach ($extra in @($extraPdfMatches)) {
        $summary.extraPdfs += $extra
    }
    foreach ($extra in @($extraDxfMatches)) {
        $summary.extraDxfs += $extra
    }

    $summary.partResults += [ordered]@{
        Part                    = $partNum
        CanonicalPart           = $partNumBase
        PdfFound                = [bool]$pdfFoundForPart
        DxfFound                = [bool]$dxfFoundForPart
        PdfCopied               = @($copiedPdfFiles)
        DxfCopied               = @($copiedDxfFiles)
        PrimaryPdfMatches       = @($primaryPdfMatches)
        PrimaryDxfMatches       = @($primaryDxfMatches)
        ExtraPdfCandidates      = @($extraPdfMatches)
        ExtraDxfCandidates      = @($extraDxfMatches)
        FallbackPdfMatches      = @($fallbackPdfMatches)
        FallbackDxfMatches      = @($fallbackDxfMatches)
        FallbackNotes           = @($fallbackNotes)
        VariantOnly             = [bool]((-not $pdfFoundForPart -and @($extraPdfMatches).Count -gt 0) -or (-not $dxfFoundForPart -and @($extraDxfMatches).Count -gt 0))
    }
}

$summary.notFound = @($summary.notFound | Select-Object -Unique)
$summary.extraPdfs = @(
    @($summary.extraPdfs) |
    Sort-Object `
        @{Expression={ [string]$_['RequestedPart'] };Descending=$false},`
        @{Expression={ [string]$_['FileName'] };Descending=$false},`
        @{Expression={ [string]$_['FullPath'] };Descending=$false} -Unique
)
$summary.extraDxfs = @(
    @($summary.extraDxfs) |
    Sort-Object `
        @{Expression={ [string]$_['RequestedPart'] };Descending=$false},`
        @{Expression={ [string]$_['FileName'] };Descending=$false},`
        @{Expression={ [string]$_['FullPath'] };Descending=$false} -Unique
)
$summary.CollectedPdfFiles = @(
    @($summary.CollectedPdfFiles) |
    Sort-Object `
        @{Expression={ [string]$_['FileName'] };Descending=$false},`
        @{Expression={ [string]$_['DestPath'] };Descending=$false},`
        @{Expression={ [string]$_['SourcePath'] };Descending=$false} -Unique
)
$summary.CollectedDxfFiles = @(
    @($summary.CollectedDxfFiles) |
    Sort-Object `
        @{Expression={ [string]$_['FileName'] };Descending=$false},`
        @{Expression={ [string]$_['DestPath'] };Descending=$false},`
        @{Expression={ [string]$_['SourcePath'] };Descending=$false} -Unique
)

if ((@($summary.extraPdfs).Count + @($summary.extraDxfs).Count) -gt 0 -and -not $copyExtraMatches) {
    $summary.warnings += "Extra parent-child matches were found but not copied because strict matching is enabled."
}

$summaryOutputPath = Join-Path $OutputFolder "collector_summary.json"
Write-CollectorLog "Done. PDFs: $($summary.pdfsFound), DXFs: $($summary.dxfsFound), Missing: $($summary.notFound.Count), Extras: $(@($summary.extraPdfs).Count + @($summary.extraDxfs).Count)"
Write-JsonUtf8 -Path $summaryPath -Value $summary -Depth 8
Write-JsonUtf8 -Path $summaryOutputPath -Value $summary -Depth 8
exit 0
