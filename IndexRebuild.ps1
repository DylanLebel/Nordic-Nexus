# ==============================================================================
#  IndexRebuild.ps1  v1.3 - Nordic Minesteel Technologies
#  Headless Index Rebuild (for Scheduled Tasks / Service Mode)
# ==============================================================================
#  v1.3 Update: High-frequency UI heartbeats during parallel scan.
# ==============================================================================

param(
    [switch]$PDFOnly,
    [switch]$DXFOnly,
    [switch]$ModelsOnly,
    [switch]$SkipModels,
    [string]$Config = "config.json"   # Override with "config.test.json" for test mode
)

$ErrorActionPreference = "Continue"

# --- Load config ---
$scriptDir  = Split-Path $PSCommandPath -Parent
$configPath = if ([System.IO.Path]::IsPathRooted($Config)) { $Config } else { Join-Path $scriptDir $Config }
$cfg = @{}
if (Test-Path $configPath) {
    try { $cfg = Get-Content $configPath -Raw | ConvertFrom-Json } catch {}
}

$pathFiltersCfg = if ($cfg.pathFilters) { $cfg.pathFilters } else { @{} }
$script:obsoleteFolderNames = @("Obsolete","Archive","Old","Deprecated")
if ($null -ne $pathFiltersCfg.obsoleteFolderNames) {
    $tmp = @($pathFiltersCfg.obsoleteFolderNames | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($tmp.Count -gt 0) { $script:obsoleteFolderNames = $tmp }
}
$script:skipCrawlFolderNames = @("QA","20 - QA")
if ($null -ne $pathFiltersCfg.skipCrawlFolderNames) {
    $tmpSkip = @($pathFiltersCfg.skipCrawlFolderNames | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($tmpSkip.Count -gt 0) { $script:skipCrawlFolderNames = @((@($script:skipCrawlFolderNames) + @($tmpSkip)) | Select-Object -Unique) }
}
$script:obsoletePathRegex = $null
if ($script:obsoleteFolderNames.Count -gt 0) {
    $escaped = @($script:obsoleteFolderNames | ForEach-Object { [regex]::Escape($_.Trim()) } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($escaped.Count -gt 0) {
        $script:obsoletePathRegex = '(?i)\\(' + ($escaped -join '|') + ')\\'
    }
}
$script:skipCrawlPathRegex = $null
if ($script:skipCrawlFolderNames.Count -gt 0) {
    $escapedSkip = @($script:skipCrawlFolderNames | ForEach-Object { [regex]::Escape($_.Trim()) } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($escapedSkip.Count -gt 0) {
        $script:skipCrawlPathRegex = '(?i)\\(' + ($escapedSkip -join '|') + ')\\'
    }
}

$indexFolder = if ($cfg.indexFolder)  { $cfg.indexFolder }  else { "C:\Users\dlebel\Documents\PDFIndex" }
$crawlRoots  = if ($cfg.crawlRoots)  { $cfg.crawlRoots }   else { @("C:\NMT_PDM","J:\Epicor","J:\MFL Jobs","J:\NordicMinesteel","Y:\") }
$logFolder   = if ($cfg.logFolder)   { $cfg.logFolder }     else { (Join-Path $indexFolder "Logs") }

# Index files
$pdfRawCSV   = Join-Path $indexFolder "pdf_index_raw.csv"
$pdfCleanCSV = Join-Path $indexFolder "pdf_index_clean.csv"
$pdfState    = Join-Path $indexFolder "crawl_state.json"
$dxfRawCSV   = Join-Path $indexFolder "dxf_index_raw.csv"
$dxfCleanCSV = Join-Path $indexFolder "dxf_index_clean.csv"
$dxfState    = Join-Path $indexFolder "dxf_crawl_state.json"
$modelRawCSV = Join-Path $indexFolder "model_index_raw.csv"
$modelAllCSV = Join-Path $indexFolder "model_index_all.csv"
$modelCleanCSV = Join-Path $indexFolder "model_index_clean.csv"
$modelState  = Join-Path $indexFolder "model_crawl_state.json"
$progressFile = Join-Path $indexFolder "progress.json"

# Ensure folders exist
if (-not (Test-Path $indexFolder)) { New-Item -ItemType Directory -Path $indexFolder -Force | Out-Null }
if (-not (Test-Path $logFolder))   { New-Item -ItemType Directory -Path $logFolder   -Force | Out-Null }

# --- Logging & Progress ---
$logFile = Join-Path $logFolder "rebuild_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').log"

function Write-ProgressFile {
    param([string]$Message, [int]$Count = 0)
    try {
        @{ Message = $Message; Count = $Count; Timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") } | 
            ConvertTo-Json | Set-Content -Path $progressFile -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {}
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $prefix = switch ($Level) {
        "ERROR"   { "[!]" }
        "WARN"    { "[~]" }
        "SUCCESS" { "[+]" }
        default   { "[ ]" }
    }
    $entry = "$ts $prefix $Message"
    Add-Content -Path $logFile -Value $entry -ErrorAction SilentlyContinue
    Write-Host $entry -ForegroundColor $(switch ($Level) {
        "ERROR" {"Red"} "WARN" {"Yellow"} "SUCCESS" {"Green"} default {"Gray"}
    })
}

# --- Revision helpers (same logic as PDFIndexManager) ---
function Get-RevisionValue {
    param([string]$RevString)
    if ($RevString -match '^\d+$') { return [int]$RevString }
    if ($RevString -match '^[A-Z]$') { return [int][char]$RevString - 64 }
    if ($RevString -match '^[A-Z]+$') {
        $value = 0
        for ($i = 0; $i -lt $RevString.Length; $i++) { $value = $value * 26 + ([int][char]$RevString[$i] - 64) }
        return $value
    }
    if ($RevString -match '([A-Z]+)(\d+)') {
        $lv = 0; $lp = $Matches[1]
        for ($i = 0; $i -lt $lp.Length; $i++) { $lv = $lv * 26 + ([int][char]$lp[$i] - 64) }
        return ($lv * 1000) + [int]$Matches[2]
    }
    return $RevString.GetHashCode()
}

function Get-PartNumberInfo {
    param([string]$BaseName)
    if ($BaseName -match '^(.+?)[\s_\-]?[Rr][Ee][Vv][\.\-_]?([A-Za-z0-9]+)$') {
        $basePart = $Matches[1].Trim()
        $revRaw   = $Matches[2].Trim().ToUpper()
        $revValue = Get-RevisionValue -RevString $revRaw
        return @{ BasePart = $basePart.ToUpper(); RevRaw = $revRaw; RevValue = $revValue; HasRev = $true }
    } else {
        return @{ BasePart = $BaseName.Trim().ToUpper(); RevRaw = $null; RevValue = 0; HasRev = $false }
    }
}

function Test-ObsoletePath {
    param([string]$FilePath)
    if ([string]::IsNullOrWhiteSpace($FilePath)) { return $false }
    if ([string]::IsNullOrWhiteSpace($script:obsoletePathRegex)) { return $false }
    return ($FilePath -match $script:obsoletePathRegex)
}

function Test-SkipCrawlPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if ([string]::IsNullOrWhiteSpace($script:skipCrawlPathRegex)) { return $false }
    return ($Path -match $script:skipCrawlPathRegex)
}

# ==============================================================================
#  Generic Parallel Crawl Engine (PDF, DXF, Models)
# ==============================================================================
function Start-ParallelCrawl {
    param(
        [string[]]$Patterns,
        [string]$RawCSV,
        [string]$StateFile,
        [string]$Label
    )

    Write-Log "Starting $Label crawl (parallel) across $($crawlRoots.Count) root(s)..." "INFO"
    Write-ProgressFile -Message "Starting $Label Crawl..." -Count 0
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    # --- Step 1: Dynamic folder expansion ---
    $scanUnits = [System.Collections.Generic.List[pscustomobject]]::new()
    $expandQueue = [System.Collections.Generic.Queue[pscustomobject]]::new()
    foreach ($root in $crawlRoots) {
        if ((Test-Path $root) -and -not (Test-SkipCrawlPath $root)) { $expandQueue.Enqueue([pscustomobject]@{ Path=$root; Depth=0 }) }
    }
    
    $cpuCount   = [Environment]::ProcessorCount
    $maxWorkers = [math]::Max(4, $cpuCount)
    $targetUnits = $maxWorkers * 4
    $maxExpandDepth = 4
    
    while ($expandQueue.Count -gt 0) {
        $item = $expandQueue.Dequeue()
        if (Test-SkipCrawlPath $item.Path) { continue }
        $shouldExpand = $false
        if ($item.Depth -lt $maxExpandDepth) {
            $norm = $item.Path.TrimEnd('\') + '\'
            if    ($norm -match '^[A-Za-z]:\\$') { $shouldExpand = $true }
            elseif ($item.Depth -lt 2)            { $shouldExpand = $true }
            elseif (($scanUnits.Count + $expandQueue.Count) -lt $targetUnits) { $shouldExpand = $true }
        }
        if ($shouldExpand) {
            try {
                $subs = [System.IO.Directory]::GetDirectories($item.Path)
                if ($subs.Count -gt 0) {
                    foreach ($s in $subs) {
                        if (Test-SkipCrawlPath $s) { continue }
                        $expandQueue.Enqueue([pscustomobject]@{ Path=$s; Depth=($item.Depth+1) })
                    }
                    $norm = $item.Path.TrimEnd('\') + '\'
                    if ($norm -notmatch '^[A-Za-z]:\\$') { $scanUnits.Add([pscustomobject]@{ Path=$item.Path; ShallowOnly=$true }) }
                } else { $scanUnits.Add([pscustomobject]@{ Path=$item.Path; ShallowOnly=$false }) }
            } catch { $scanUnits.Add([pscustomobject]@{ Path=$item.Path; ShallowOnly=$false }) }
        } else { $scanUnits.Add([pscustomobject]@{ Path=$item.Path; ShallowOnly=$false }) }
    }
    Write-Log "$Label crawl: $($scanUnits.Count) scan units, $maxWorkers max workers" "INFO"

    $tempDir = Join-Path $indexFolder ("crawl_tmp_" + $Label.ToLower())
    if (-not (Test-Path $tempDir)) { New-Item -ItemType Directory -Path $tempDir -Force | Out-Null }
    Get-ChildItem -Path $tempDir -Filter "chunk_*.csv" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue

    # Write Header
    $header = '"FileName","BaseName","FullPath","LastWriteTime","SizeBytes","RootFolder"'
    if ($Label -eq "MODEL") { $header += ',"FileType"' }
    [System.IO.File]::WriteAllText($RawCSV, $header + [Environment]::NewLine)

    # --- Step 2: RunspacePool ---
    $pool     = [RunspaceFactory]::CreateRunspacePool(1, $maxWorkers)
    $pool.Open()
    $counters = [hashtable]::Synchronized(@{ Total=0; Errors=0; Done=0 })
    $workers  = [System.Collections.Generic.List[object]]::new()

    $workerScript = {
        param($ScanRoot, $ShallowOnly, $ChunkFile, $Counters, $Patterns, $Label, $SkipPathRegex)
        $localCount=0; $localErrors=0
        $rootEsc=$ScanRoot.Replace('"','""')
        $buf=[System.Text.StringBuilder]::new(16384); $bufLines=0; $stream=$null
        try {
            $stream=[System.IO.StreamWriter]::new($ChunkFile,$false,[System.Text.Encoding]::UTF8,65536)
            $stack=[System.Collections.Generic.Stack[string]]::new(); $stack.Push($ScanRoot)
            
            while($stack.Count-gt 0){
                $cur=$stack.Pop()
                if ($SkipPathRegex -and $cur -match $SkipPathRegex) { continue }
                try {
                    foreach ($pat in $Patterns) {
                        foreach($fp in [System.IO.Directory]::EnumerateFiles($cur,$pat)){
                            try{
                                $fi=[System.IO.FileInfo]::new($fp)
                                if ($fi.Length -eq 0 -and $Label -ne "MODEL") { continue } # Skip 0-byte PDFs/DXFs
                                
                                $fn=[System.IO.Path]::GetFileName($fp); $bn=[System.IO.Path]::GetFileNameWithoutExtension($fp)
                                $lw=$fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                                
                                [void]$buf.Append('"').Append($fn.Replace('"','""')).Append('","').Append($bn.Replace('"','""')).Append('","').Append($fp.Replace('"','""')).Append('","').Append($lw).Append('",').Append($fi.Length).Append(',"').Append($rootEsc).Append('"')
                                
                                if ($Label -eq "MODEL") {
                                    $ft=[System.IO.Path]::GetExtension($fp).TrimStart('.').ToUpperInvariant()
                                    [void]$buf.Append(',"').Append($ft).Append('"')
                                }
                                [void]$buf.AppendLine()
                                
                                $localCount++; $bufLines++
                                
                                if($bufLines-ge 200){$stream.Write($buf.ToString());$buf.Clear();$bufLines=0}
                            }catch{$localErrors++}
                        }
                    }
                    if (-not $ShallowOnly) {
                        foreach($sd in [System.IO.Directory]::EnumerateDirectories($cur)){
                            if ($SkipPathRegex -and $sd -match $SkipPathRegex) { continue }
                            $stack.Push($sd)
                        }
                    }
                }catch{$localErrors++}
                if ($ShallowOnly) { break }
            }
        } finally {
            if($bufLines-gt 0-and $stream){$stream.Write($buf.ToString())}
            if($stream){$stream.Flush();$stream.Close();$stream.Dispose()}
        }
        [System.Threading.Monitor]::Enter($Counters)
        try { $Counters.Total += $localCount; $Counters.Errors += $localErrors; $Counters.Done += 1 }
        finally { [System.Threading.Monitor]::Exit($Counters) }
    }

    # --- Step 3: Launch workers ---
    foreach ($unit in $scanUnits) {
        $safeUnitName = ($unit.Path -replace '[\\:/\s]','_')
        if ($safeUnitName.Length -gt 80) { $safeUnitName = $safeUnitName.Substring(0,80) }
        $suffix    = if ($unit.ShallowOnly) { "_SH" } else { "" }
        $chunkFile = Join-Path $tempDir ("chunk_" + $safeUnitName + $suffix + ".csv")
        $wPs = [PowerShell]::Create(); $wPs.RunspacePool = $pool
        [void]$wPs.AddScript($workerScript)
        [void]$wPs.AddArgument($unit.Path); [void]$wPs.AddArgument($unit.ShallowOnly)
        [void]$wPs.AddArgument($chunkFile); [void]$wPs.AddArgument($counters); [void]$wPs.AddArgument($Patterns); [void]$wPs.AddArgument($Label); [void]$wPs.AddArgument($script:skipCrawlPathRegex)
        $wHandle = $wPs.BeginInvoke()
        $workers.Add([pscustomobject]@{ PS=$wPs; Handle=$wHandle; Chunk=$chunkFile; Root=$unit.Path; ShallowOnly=$unit.ShallowOnly })
    }

    # --- Step 4: Wait for all workers ---
    while ($counters.Done -lt $workers.Count) {
        $percent = [math]::Round(($counters.Done / $workers.Count) * 100, 0)
        Write-ProgressFile -Message "Scanning ${Label}: $percent% complete" -Count $counters.Total
        Start-Sleep -Milliseconds 500
    }

    foreach ($w in $workers) {
        try { $w.PS.EndInvoke($w.Handle); $w.PS.Dispose() } catch {}
    }
    try { $pool.Close(); $pool.Dispose() } catch {}

    # --- Step 5: Merge chunk files ---
    Write-Log "  Merging $($workers.Count) chunk files..." "INFO"
    Write-ProgressFile -Message "Finalizing $Label index..." -Count $counters.Total
    
    $ms=$null
    try {
        $ms=[System.IO.StreamWriter]::new($RawCSV,$true,[System.Text.Encoding]::UTF8,65536)
        foreach ($w in $workers) {
            if (Test-Path $w.Chunk) {
                $reader=$null
                try {
                    $reader=[System.IO.StreamReader]::new($w.Chunk,[System.Text.Encoding]::UTF8,$true,65536)
                    while(-not $reader.EndOfStream){$line=$reader.ReadLine();if($line){$ms.WriteLine($line)}}
                } finally { if($reader){$reader.Close();$reader.Dispose()} }
                Remove-Item $w.Chunk -Force -ErrorAction SilentlyContinue
            }
        }
    } finally { if($ms){$ms.Flush();$ms.Close();$ms.Dispose()} }
    
    try { if((Get-ChildItem $tempDir -ErrorAction SilentlyContinue|Measure-Object).Count-eq 0){Remove-Item $tempDir -Force -ErrorAction SilentlyContinue} } catch {}

    $sw.Stop()
    $elapsed = $sw.Elapsed.ToString('hh\:mm\:ss')

    # Save state
    try {
        $state = @{
            LastCrawlTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
            TotalFiles    = [int]$counters.Total
            TotalErrors   = [int]$counters.Errors
            ElapsedTime   = $elapsed
            FoldersCrawled = $crawlRoots
            Mode          = "Parallel"
        } | ConvertTo-Json
        [System.IO.File]::WriteAllText($StateFile, $state)
    } catch {}

    Write-Log "$Label CRAWL COMPLETE: $($counters.Total) files, $elapsed" "SUCCESS"
    return @{ Files = $counters.Total; Errors = $counters.Errors; Elapsed = $elapsed }
}

# ==============================================================================
#  Dedup Function (headless)
# ==============================================================================
function Start-Dedup {
    param(
        [string]$RawCSV,
        [string]$CleanCSV,
        [string]$Label
    )

    if (-not (Test-Path $RawCSV)) {
        Write-Log "No raw $Label index found at $RawCSV - cannot deduplicate" "ERROR"
        return $null
    }

    Write-Log "Starting $Label deduplication..." "INFO"
    Write-ProgressFile -Message "Deduplicating $Label..."

    $rawData = Import-Csv -Path $RawCSV
    Write-Log "Loaded $($rawData.Count) raw $Label entries" "INFO"

    $partGroups = @{}
    foreach ($row in $rawData) {
        try {
            $info = Get-PartNumberInfo -BaseName $row.BaseName
            $key  = $info.BasePart
            $fileType = ""
            try { $fileType = [string]$row.FileType } catch {}
            if ([string]::IsNullOrWhiteSpace($fileType)) {
                try { $fileType = [System.IO.Path]::GetExtension([string]$row.FileName).TrimStart('.').ToUpperInvariant() } catch {}
            }
            if (-not $partGroups.ContainsKey($key)) { $partGroups[$key] = [System.Collections.ArrayList]::new() }
            [void]$partGroups[$key].Add(@{
                FileName=$row.FileName; BaseName=$row.BaseName; FullPath=$row.FullPath
                LastWriteTime=[datetime]$row.LastWriteTime; SizeBytes=[long]$row.SizeBytes
                RootFolder=$row.RootFolder; BasePart=$info.BasePart; RevRaw=$info.RevRaw
                RevValue=$info.RevValue; HasRev=$info.HasRev
                IsObsolete=(Test-ObsoletePath $row.FullPath); FileType=$fileType
            })
        } catch { continue }
    }

    $cleanResults = [System.Collections.ArrayList]::new()
    foreach ($key in $partGroups.Keys) {
        $candidates = $partGroups[$key]
        if ($candidates.Count -eq 1) { $winner = $candidates[0] }
        else {
            $withRev = @($candidates | Where-Object { $_.HasRev -eq $true })
            if ($withRev.Count -gt 0) {
                $winner = $withRev | Sort-Object @{Expression={$_.IsObsolete};Ascending=$true},@{Expression={$_.RevValue};Descending=$true},@{Expression={$_.LastWriteTime};Descending=$true} | Select-Object -First 1
            } else {
                $winner = $candidates | Sort-Object @{Expression={$_.IsObsolete};Ascending=$true},@{Expression={$_.LastWriteTime};Descending=$true} | Select-Object -First 1
            }
        }
        $revDisplay = if ($winner.RevRaw) { $winner.RevRaw } else { "(none)" }
        $isObs = if ($winner.IsObsolete) { "Yes" } else { "No" }
        [void]$cleanResults.Add([PSCustomObject]@{
            BasePart=$winner.BasePart; Rev=$revDisplay; FileName=$winner.FileName
            FullPath=$winner.FullPath; LastWriteTime=$winner.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
            SizeKB=[math]::Round($winner.SizeBytes/1024,1); RootFolder=$winner.RootFolder; Candidates=$candidates.Count
            IsObsolete=$isObs; FileType=$winner.FileType
        })
    }

    $cleanResults = $cleanResults | Sort-Object BasePart
    $cleanResults | Export-Csv -Path $CleanCSV -NoTypeInformation -Encoding UTF8

    $dupes = $rawData.Count - $cleanResults.Count
    $obsoleteCount = @($cleanResults | Where-Object { $_.IsObsolete -eq "Yes" }).Count
    Write-Log "$Label DEDUP COMPLETE: $($cleanResults.Count) unique from $($rawData.Count) total ($dupes duplicates, $obsoleteCount obsolete)" "SUCCESS"

    return @{
        Total   = $rawData.Count
        Unique  = $cleanResults.Count
        Dupes   = $dupes
        Obsolete = $obsoleteCount
    }
}

# ==============================================================================
#  Main
# ==============================================================================

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  INDEX REBUILD (PARALLEL v1.3)" -ForegroundColor Cyan
Write-Host "  Nordic Minesteel Technologies" -ForegroundColor DarkCyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

$startTime = Get-Date
Write-Log "Index rebuild started at $($startTime.ToString('yyyy-MM-dd HH:mm:ss'))" "INFO"
Write-Log "Index folder: $indexFolder" "INFO"
Write-Log "Crawl roots: $($crawlRoots -join ', ')" "INFO"

$doPDF = (-not $DXFOnly) -and (-not $ModelsOnly)
$doDXF = (-not $PDFOnly) -and (-not $ModelsOnly)
$doModels = (-not $PDFOnly) -and (-not $DXFOnly)
if ($ModelsOnly) { $doModels = $true }
if ($SkipModels) { $doModels = $false }

$results = @{}

# PDF crawl + dedup
if ($doPDF) {
    Write-Host ""
    $results.PDFCrawl = Start-ParallelCrawl -Patterns @("*.pdf") -RawCSV $pdfRawCSV -StateFile $pdfState -Label "PDF"
    $results.PDFDedup = Start-Dedup -RawCSV $pdfRawCSV -CleanCSV $pdfCleanCSV -Label "PDF"
}

# DXF crawl + dedup
if ($doDXF) {
    Write-Host ""
    $results.DXFCrawl = Start-ParallelCrawl -Patterns @("*.dxf") -RawCSV $dxfRawCSV -StateFile $dxfState -Label "DXF"
    $results.DXFDedup = Start-Dedup -RawCSV $dxfRawCSV -CleanCSV $dxfCleanCSV -Label "DXF"
}

# Model crawl + dedup
if ($doModels) {
    Write-Host ""
    $results.ModelCrawl = Start-ParallelCrawl -Patterns @("*.sldasm","*.sldprt") -RawCSV $modelRawCSV -StateFile $modelState -Label "MODEL"
    $results.ModelDedup = Start-Dedup -RawCSV $modelRawCSV -CleanCSV $modelCleanCSV -Label "MODEL"
    # Copy model index to 'all' version for legacy support
    try { Copy-Item -Path $modelRawCSV -Destination $modelAllCSV -Force } catch {}
}

# Cleanup progress file
if (Test-Path $progressFile) { Remove-Item $progressFile -Force -ErrorAction SilentlyContinue }

$endTime = Get-Date
$totalElapsed = ($endTime - $startTime).ToString('hh\:mm\:ss')

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  REBUILD COMPLETE" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
if ($results.PDFDedup) {
    Write-Host "  PDF: $($results.PDFDedup.Unique) unique drawings" -ForegroundColor Green
}
if ($results.DXFDedup) {
    Write-Host "  DXF: $($results.DXFDedup.Unique) unique drawings" -ForegroundColor Green
}
if ($results.ModelDedup) {
    Write-Host "  Models: $($results.ModelDedup.Unique) unique model base parts" -ForegroundColor Green
}
Write-Host "  Total time: $totalElapsed" -ForegroundColor White
Write-Host "==========================================" -ForegroundColor Cyan

Write-Log "Index rebuild completed in $totalElapsed" "SUCCESS"

# Save rebuild summary for dashboard
$rebuildSummary = @{
    Timestamp   = $endTime.ToString("yyyy-MM-dd HH:mm:ss")
    Elapsed     = $totalElapsed
    PDFCrawl    = $results.PDFCrawl
    PDFDedup    = $results.PDFDedup
    DXFCrawl    = $results.DXFCrawl
    DXFDedup    = $results.DXFDedup
    ModelCrawl  = $results.ModelCrawl
    ModelDedup  = $results.ModelDedup
    ModelAllCSV = $modelAllCSV
    ModelCleanCSV = $modelCleanCSV
    LogFile     = $logFile
}
$rebuildSummary | ConvertTo-Json -Depth 3 | Set-Content -Path (Join-Path $indexFolder "last_rebuild.json") -Encoding UTF8
