# ==============================================================================
#  IndexClean.ps1  v2.0 - Nordic Minesteel Technologies
#  Drawing File Cleanup Tool
# ==============================================================================
#  Analyzes the raw index to find duplicate files and old revisions on disk,
#  then cleans them up by moving losers to a quarantine folder (or deleting).
#
#  The "winner" for each part is the same logic as IndexRebuild:
#    1. Non-obsolete beats obsolete
#    2. Highest revision wins
#    3. Newest date breaks ties
#  Everything else is a "loser" - an old rev, a duplicate copy, junk.
#
#  Usage:
#    IndexClean.ps1                     # Analyze only - show what would be cleaned
#    IndexClean.ps1 -Clean              # Move losers to quarantine folder
#    IndexClean.ps1 -Clean -Delete      # Delete losers instead of quarantine
#    IndexClean.ps1 -Type PDF           # PDF files only
#    IndexClean.ps1 -Type DXF           # DXF files only
#    IndexClean.ps1 -ExportReport       # Save full analysis to CSV
# ==============================================================================

param(
    [switch]$Clean,
    [switch]$Delete,
    [switch]$ExportReport,
    [ValidateSet("ALL","PDF","DXF")]
    [string]$Type = "ALL",
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
$script:obsoletePathRegex = $null
if ($script:obsoleteFolderNames.Count -gt 0) {
    $escaped = @($script:obsoleteFolderNames | ForEach-Object { [regex]::Escape($_.Trim()) } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($escaped.Count -gt 0) {
        $script:obsoletePathRegex = '(?i)\\(' + ($escaped -join '|') + ')\\'
    }
}

$indexFolder    = if ($cfg.indexFolder) { $cfg.indexFolder } else { "C:\Users\dlebel\Documents\PDFIndex" }
$logFolder      = if ($cfg.logFolder)  { $cfg.logFolder }  else { (Join-Path $indexFolder "Logs") }
$quarantineRoot = Join-Path $indexFolder "Quarantine"
$protectedRoots = @()
if ($cfg.protectedRoots) { $protectedRoots = @($cfg.protectedRoots) }

if (-not (Test-Path $logFolder)) { New-Item -ItemType Directory -Path $logFolder -Force | Out-Null }

# --- Log file ---
$logFile = Join-Path $logFolder "cleanup_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').log"
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $prefix = switch ($Level) {
        "ERROR"   { "[!]" }
        "WARN"    { "[~]" }
        "SUCCESS" { "[+]" }
        "CLEAN"   { "[x]" }
        default   { "[ ]" }
    }
    $entry = "$ts $prefix $Message"
    Add-Content -Path $logFile -Value $entry -ErrorAction SilentlyContinue
}

# --- Revision helpers (same as IndexRebuild) ---
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

function Test-ProtectedPath {
    param([string]$FilePath)
    foreach ($root in $protectedRoots) {
        if ($FilePath -like "$root\*" -or $FilePath -eq $root) { return $true }
    }
    return $false
}

# ==============================================================================
#  Analyze: group raw index entries, pick winners, identify losers
# ==============================================================================
function Get-CleanupPlan {
    param(
        [string]$Label,
        [string]$RawCSV
    )

    $plan = @{
        Label     = $Label
        Winners   = [System.Collections.ArrayList]::new()
        Losers    = [System.Collections.ArrayList]::new()
        Stale     = [System.Collections.ArrayList]::new()
        PartCount = 0
    }

    if (-not (Test-Path $RawCSV)) {
        Write-Host "  [!] Raw index not found: $RawCSV" -ForegroundColor Yellow
        return $plan
    }

    $rawData = Import-Csv -Path $RawCSV
    Write-Host "  Raw entries: $($rawData.Count)" -ForegroundColor Gray

    # Group by base part number (same logic as IndexRebuild)
    $partGroups = @{}
    foreach ($row in $rawData) {
        try {
            # Check if file still exists
            if (-not (Test-Path $row.FullPath)) {
                [void]$plan.Stale.Add([PSCustomObject]@{
                    BasePart = $row.BaseName
                    FileName = $row.FileName
                    FullPath = $row.FullPath
                    Reason   = "File no longer on disk"
                })
                continue
            }

            $info = Get-PartNumberInfo -BaseName $row.BaseName
            $key  = $info.BasePart
            if (-not $partGroups.ContainsKey($key)) { $partGroups[$key] = [System.Collections.ArrayList]::new() }
            [void]$partGroups[$key].Add(@{
                FileName      = $row.FileName
                BaseName      = $row.BaseName
                FullPath      = $row.FullPath
                LastWriteTime = [datetime]$row.LastWriteTime
                SizeBytes     = [long]$row.SizeBytes
                RootFolder    = $row.RootFolder
                BasePart      = $info.BasePart
                RevRaw        = $info.RevRaw
                RevValue      = $info.RevValue
                HasRev        = $info.HasRev
                IsObsolete    = (Test-ObsoletePath $row.FullPath)
                IsProtected   = (Test-ProtectedPath $row.FullPath)
            })
        } catch { continue }
    }

    $plan.PartCount = $partGroups.Count

    # For each part group, pick winner and mark losers
    foreach ($key in $partGroups.Keys) {
        $candidates = $partGroups[$key]

        # Only 1 file - it's the winner, nothing to clean
        if ($candidates.Count -eq 1) {
            [void]$plan.Winners.Add($candidates[0])
            continue
        }

        # Pick winner (same priority as IndexRebuild)
        $withRev = @($candidates | Where-Object { $_.HasRev -eq $true })
        if ($withRev.Count -gt 0) {
            $winner = $withRev | Sort-Object `
                @{Expression={$_.IsObsolete};Ascending=$true},
                @{Expression={$_.RevValue};Descending=$true},
                @{Expression={$_.LastWriteTime};Descending=$true} | Select-Object -First 1
        } else {
            $winner = $candidates | Sort-Object `
                @{Expression={$_.IsObsolete};Ascending=$true},
                @{Expression={$_.LastWriteTime};Descending=$true} | Select-Object -First 1
        }

        [void]$plan.Winners.Add($winner)

        # Everyone else is a loser (unless protected)
        foreach ($c in $candidates) {
            if ($c.FullPath -eq $winner.FullPath) { continue }

            # Never touch files in protected roots (the corner / PDM vault)
            if ($c.IsProtected) { continue }

            $reason = ""
            if ($c.IsObsolete) {
                $reason = "In obsolete folder"
            } elseif ($c.HasRev -and $winner.HasRev -and $c.RevValue -lt $winner.RevValue) {
                $reason = "Old revision (Rev $($c.RevRaw) < Rev $($winner.RevRaw))"
            } elseif (-not $c.HasRev -and $winner.HasRev) {
                $reason = "No revision tag (winner has Rev $($winner.RevRaw))"
            } else {
                $reason = "Duplicate copy"
            }

            [void]$plan.Losers.Add([PSCustomObject]@{
                BasePart   = $key
                FileName   = $c.FileName
                FullPath   = $c.FullPath
                RevRaw     = $c.RevRaw
                IsObsolete = $c.IsObsolete
                SizeBytes  = $c.SizeBytes
                Reason     = $reason
                WinnerFile = $winner.FileName
                WinnerPath = $winner.FullPath
                WinnerRev  = $winner.RevRaw
            })
        }
    }

    return $plan
}

# ==============================================================================
#  Clean: move or delete loser files
# ==============================================================================
function Invoke-Cleanup {
    param(
        [object]$Plan,
        [bool]$DeleteMode
    )

    $moved = 0; $deleted = 0; $failed = 0; $skipped = 0
    $totalSize = [long]0

    foreach ($loser in $Plan.Losers) {
        # Double-check file still exists
        if (-not (Test-Path $loser.FullPath)) {
            $skipped++
            continue
        }

        # Safety: never touch files in protected roots
        if (Test-ProtectedPath $loser.FullPath) {
            Write-Log "BLOCKED: $($loser.FullPath) is in a protected root - skipping" "WARN"
            $skipped++
            continue
        }

        try {
            if ($DeleteMode) {
                Remove-Item -Path $loser.FullPath -Force
                $deleted++
                Write-Log "$($loser.FullPath) -> DELETED ($($loser.Reason))" "CLEAN"
            } else {
                # Mirror the folder structure under quarantine
                $dateFolder = Get-Date -Format "yyyy-MM-dd"
                $quarantineSub = Join-Path $quarantineRoot $dateFolder
                if (-not (Test-Path $quarantineSub)) { New-Item -ItemType Directory -Path $quarantineSub -Force | Out-Null }

                $destName = $loser.FileName
                $destPath = Join-Path $quarantineSub $destName
                $counter = 1
                while (Test-Path $destPath) {
                    $bn  = [System.IO.Path]::GetFileNameWithoutExtension($loser.FileName)
                    $ext = [System.IO.Path]::GetExtension($loser.FileName)
                    $destPath = Join-Path $quarantineSub "${bn}_dup${counter}${ext}"
                    $counter++
                }

                Move-Item -Path $loser.FullPath -Destination $destPath -Force
                $moved++
                Write-Log "$($loser.FullPath) -> $destPath ($($loser.Reason))" "CLEAN"
            }
            $totalSize += $loser.SizeBytes
        } catch {
            $failed++
            Write-Log "FAILED: $($loser.FullPath) - $($_.Exception.Message)" "ERROR"
        }
    }

    return @{ Moved = $moved; Deleted = $deleted; Failed = $failed; Skipped = $skipped; SizeBytes = $totalSize }
}

# ==============================================================================
#  Main
# ==============================================================================

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  DRAWING FILE CLEANUP" -ForegroundColor Cyan
Write-Host "  Nordic Minesteel Technologies" -ForegroundColor DarkCyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host ""

if ($Clean -and $Delete) {
    Write-Host "  Mode: DELETE (permanently remove duplicates/old revs)" -ForegroundColor Red
} elseif ($Clean) {
    Write-Host "  Mode: CLEAN (move duplicates/old revs to quarantine)" -ForegroundColor Yellow
    Write-Host "  Quarantine: $quarantineRoot" -ForegroundColor Gray
} else {
    Write-Host "  Mode: ANALYZE (show what would be cleaned)" -ForegroundColor White
}
Write-Host "  Index folder: $indexFolder" -ForegroundColor Gray
if ($protectedRoots.Count -gt 0) {
    Write-Host "  Protected (no cleanup): $($protectedRoots -join ', ')" -ForegroundColor DarkCyan
}
Write-Host ""

$allPlans = @()
$totalLosers = 0; $totalStale = 0; $totalSizeBytes = [long]0

# PDF
if ($Type -eq "ALL" -or $Type -eq "PDF") {
    $pdfRaw = Join-Path $indexFolder "pdf_index_raw.csv"
    Write-Host "--- PDF Files ---" -ForegroundColor Cyan
    $pdfPlan = Get-CleanupPlan -Label "PDF" -RawCSV $pdfRaw
    $allPlans += $pdfPlan
    $loserSize = ($pdfPlan.Losers | Measure-Object -Property SizeBytes -Sum).Sum
    if (-not $loserSize) { $loserSize = 0 }
    Write-Host "  Unique parts:  $($pdfPlan.PartCount)" -ForegroundColor Gray
    Write-Host "  Winners:       $($pdfPlan.Winners.Count) (keep)" -ForegroundColor Green
    Write-Host "  Losers:        $($pdfPlan.Losers.Count) (old revs / duplicates)" -ForegroundColor $(if($pdfPlan.Losers.Count -gt 0){"Yellow"}else{"Green"})
    if ($pdfPlan.Stale.Count -gt 0) {
        Write-Host "  Gone from disk: $($pdfPlan.Stale.Count)" -ForegroundColor DarkGray
    }
    if ($pdfPlan.Losers.Count -gt 0) {
        Write-Host "  Reclaimable:   $([math]::Round($loserSize / 1MB, 1)) MB" -ForegroundColor Cyan
    }
    $totalLosers += $pdfPlan.Losers.Count
    $totalStale += $pdfPlan.Stale.Count
    $totalSizeBytes += $loserSize
    Write-Host ""

    # Show breakdown by reason
    if ($pdfPlan.Losers.Count -gt 0) {
        $reasons = $pdfPlan.Losers | Group-Object Reason | Sort-Object Count -Descending
        foreach ($r in $reasons) {
            Write-Host "    $($r.Count.ToString().PadLeft(5)) - $($r.Name)" -ForegroundColor DarkYellow
        }
        Write-Host ""
    }
}

# DXF
if ($Type -eq "ALL" -or $Type -eq "DXF") {
    $dxfRaw = Join-Path $indexFolder "dxf_index_raw.csv"
    Write-Host "--- DXF Files ---" -ForegroundColor Cyan
    $dxfPlan = Get-CleanupPlan -Label "DXF" -RawCSV $dxfRaw
    $allPlans += $dxfPlan
    $loserSize = ($dxfPlan.Losers | Measure-Object -Property SizeBytes -Sum).Sum
    if (-not $loserSize) { $loserSize = 0 }
    Write-Host "  Unique parts:  $($dxfPlan.PartCount)" -ForegroundColor Gray
    Write-Host "  Winners:       $($dxfPlan.Winners.Count) (keep)" -ForegroundColor Green
    Write-Host "  Losers:        $($dxfPlan.Losers.Count) (old revs / duplicates)" -ForegroundColor $(if($dxfPlan.Losers.Count -gt 0){"Yellow"}else{"Green"})
    if ($dxfPlan.Stale.Count -gt 0) {
        Write-Host "  Gone from disk: $($dxfPlan.Stale.Count)" -ForegroundColor DarkGray
    }
    if ($dxfPlan.Losers.Count -gt 0) {
        Write-Host "  Reclaimable:   $([math]::Round($loserSize / 1MB, 1)) MB" -ForegroundColor Cyan
    }
    $totalLosers += $dxfPlan.Losers.Count
    $totalStale += $dxfPlan.Stale.Count
    $totalSizeBytes += $loserSize
    Write-Host ""

    if ($dxfPlan.Losers.Count -gt 0) {
        $reasons = $dxfPlan.Losers | Group-Object Reason | Sort-Object Count -Descending
        foreach ($r in $reasons) {
            Write-Host "    $($r.Count.ToString().PadLeft(5)) - $($r.Name)" -ForegroundColor DarkYellow
        }
        Write-Host ""
    }
}

# --- Export report ---
if ($ExportReport) {
    $reportPath = Join-Path $indexFolder "cleanup_report_$(Get-Date -Format 'yyyy-MM-dd').csv"
    $reportRows = [System.Collections.ArrayList]::new()
    foreach ($p in $allPlans) {
        foreach ($l in $p.Losers) {
            [void]$reportRows.Add([PSCustomObject]@{
                Type       = $p.Label
                Action     = "REMOVE"
                BasePart   = $l.BasePart
                FileName   = $l.FileName
                FullPath   = $l.FullPath
                Rev        = $l.RevRaw
                Reason     = $l.Reason
                SizeKB     = [math]::Round($l.SizeBytes / 1024, 1)
                WinnerFile = $l.WinnerFile
                WinnerRev  = $l.WinnerRev
            })
        }
        foreach ($w in $p.Winners) {
            [void]$reportRows.Add([PSCustomObject]@{
                Type       = $p.Label
                Action     = "KEEP"
                BasePart   = $w.BasePart
                FileName   = $w.FileName
                FullPath   = $w.FullPath
                Rev        = $w.RevRaw
                Reason     = "Winner"
                SizeKB     = [math]::Round($w.SizeBytes / 1024, 1)
                WinnerFile = ""
                WinnerRev  = ""
            })
        }
        foreach ($s in $p.Stale) {
            [void]$reportRows.Add([PSCustomObject]@{
                Type       = $p.Label
                Action     = "STALE"
                BasePart   = $s.BasePart
                FileName   = $s.FileName
                FullPath   = $s.FullPath
                Rev        = ""
                Reason     = $s.Reason
                SizeKB     = 0
                WinnerFile = ""
                WinnerRev  = ""
            })
        }
    }
    $reportRows | Sort-Object Type, BasePart, Action | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8
    Write-Host "[+] Report saved: $reportPath ($($reportRows.Count) entries)" -ForegroundColor Green
    Write-Host ""
}

# --- Execute cleanup ---
if ($Clean -and $totalLosers -gt 0) {
    Write-Host "==========================================" -ForegroundColor Yellow
    if ($Delete) {
        Write-Host "  DELETING $totalLosers files ($([math]::Round($totalSizeBytes / 1MB, 1)) MB)" -ForegroundColor Red
    } else {
        Write-Host "  QUARANTINING $totalLosers files ($([math]::Round($totalSizeBytes / 1MB, 1)) MB)" -ForegroundColor Yellow
    }
    Write-Host "==========================================" -ForegroundColor Yellow
    Write-Host ""

    foreach ($p in $allPlans) {
        if ($p.Losers.Count -eq 0) { continue }
        Write-Host "  Cleaning $($p.Label) files..." -ForegroundColor Gray
        $result = Invoke-Cleanup -Plan $p -DeleteMode $Delete
        if ($Delete) {
            Write-Host "    Deleted: $($result.Deleted)  Failed: $($result.Failed)  Skipped: $($result.Skipped)" -ForegroundColor $(if($result.Failed -gt 0){"Yellow"}else{"Green"})
        } else {
            Write-Host "    Moved: $($result.Moved)  Failed: $($result.Failed)  Skipped: $($result.Skipped)" -ForegroundColor $(if($result.Failed -gt 0){"Yellow"}else{"Green"})
        }
    }
    Write-Host ""
    Write-Host "  [!] Run IndexRebuild.ps1 to update indexes after cleanup" -ForegroundColor Yellow
    Write-Host "  Log: $logFile" -ForegroundColor Gray
    Write-Host ""
}

# --- Summary ---
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  SUMMARY" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

if ($totalLosers -eq 0) {
    Write-Host "  Nothing to clean - file system looks good" -ForegroundColor Green
} else {
    Write-Host "  Files to clean: $totalLosers ($([math]::Round($totalSizeBytes / 1MB, 1)) MB)" -ForegroundColor Yellow
    if (-not $Clean) {
        Write-Host ""
        Write-Host "  Run with -Clean to quarantine these files" -ForegroundColor Gray
        Write-Host "  Run with -Clean -Delete to permanently remove" -ForegroundColor Gray
        Write-Host "  Run with -ExportReport to review list in Excel first" -ForegroundColor Gray
    }
}

if ($totalStale -gt 0) {
    Write-Host "  Stale index entries: $totalStale (files already gone from disk)" -ForegroundColor DarkGray
}

Write-Host "==========================================" -ForegroundColor Cyan
