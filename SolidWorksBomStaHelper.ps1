param(
    [Parameter(Mandatory = $true)]
    [string]$AssemblyPath,
    [string]$OutputFile = ""
)

$ErrorActionPreference = "Stop"

function Out-Log {
    param([string]$Message)
    $line = "LOG|" + $Message
    if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
        Add-Content -Path $OutputFile -Value $line -ErrorAction SilentlyContinue
    } else {
        Write-Output $line
    }
}

function Out-Part {
    param([string]$PartNumber)
    if (-not [string]::IsNullOrWhiteSpace($PartNumber)) {
        $line = "PART|" + $PartNumber.Trim().ToUpperInvariant()
        if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
            Add-Content -Path $OutputFile -Value $line -ErrorAction SilentlyContinue
        } else {
            Write-Output $line
        }
    }
}

function Get-FileBaseNameUpper {
    param([string]$PathOrName)
    if ([string]::IsNullOrWhiteSpace($PathOrName)) { return "" }
    try { return [System.IO.Path]::GetFileNameWithoutExtension($PathOrName).Trim().ToUpperInvariant() } catch { return "" }
}

function Get-CompModelPath {
    param([object]$Comp)
    if ($null -eq $Comp) { return "" }
    $p = ""
    try { $p = [string]$Comp.GetPathName() } catch { }
    if ([string]::IsNullOrWhiteSpace($p)) {
        try {
            $m = $Comp.GetModelDoc2()
            if ($m) { $p = [string]$m.GetPathName() }
        } catch { }
    }
    return $p
}

function Get-CompPartNumber {
    param([object]$Comp)
    if ($null -eq $Comp) { return "" }

    # Optimization: If the filename itself looks like a standard part number (e.g. 12345-A01),
    # use it directly and skip the expensive COM property probe.
    $mp = Get-CompModelPath -Comp $Comp
    $part = Get-FileBaseNameUpper -PathOrName $mp
    if ($part -match '^\d{4,6}-[A-Z][A-Z0-9]{1,8}(?:-[LR])?$') {
        return $part
    }

    try {
        $m = $Comp.GetModelDoc2()
        if ($m) {
            $mgr = $null
            try { $mgr = $m.Extension.CustomPropertyManager("") } catch { }
            if ($mgr) {
                foreach ($n in @("Part Number","PartNumber","PartNo","Part No","DrawingNo","Drawing No","Number")) {
                    $v = ""; $rv = ""
                    try { $null = $mgr.Get4($n, $false, [ref]$v, [ref]$rv) } catch { }
                    if (-not [string]::IsNullOrWhiteSpace($rv)) { $part = $rv; break }
                    if (-not [string]::IsNullOrWhiteSpace($v)) { $part = $v; break }
                }
            }
        }
    } catch { }

    return ([string]$part).Trim().ToUpperInvariant()
}

function Collect-PartsRecursive {
    param(
        [object]$Comp,
        [hashtable]$PartSet,
        [hashtable]$SeenModelCfg
    )
    if ($null -eq $Comp) { return }

    $suppressed = $false
    try { $suppressed = [bool]$Comp.IsSuppressed() } catch { }
    if ($suppressed) { return }

    $modelPath = Get-CompModelPath -Comp $Comp
    if (-not [string]::IsNullOrWhiteSpace($modelPath)) {
        $cfg = ""
        try { $cfg = [string]$Comp.ReferencedConfiguration } catch { }
        $k = ($modelPath + "|" + $cfg).ToLowerInvariant()
        if ($SeenModelCfg.ContainsKey($k)) { return }
        $SeenModelCfg[$k] = $true
    }

    $pn = Get-CompPartNumber -Comp $Comp
    if (-not [string]::IsNullOrWhiteSpace($pn)) {
        if ($pn -notmatch '(?i)LOAD[\s_]?CERT|SCOPE|MANUAL') {
            $PartSet[$pn] = $true
        }
    }

    # Resolve sub-assembly children (top-level ResolveAllLightWeightComponents should handle this,
    # but we check if Children() returns null to be safe).
    $children = $null
    try { $children = $Comp.GetChildren() } catch { }
    
    # Fallback: if no children but it's an assembly, try resolving just this one component.
    if ($null -eq $children) {
        try {
            $compDoc = $Comp.GetModelDoc2()
            if ($compDoc -and $compDoc.GetType() -eq 2) {
                $null = $compDoc.ResolveAllLightWeightComponents($true)
                $children = $Comp.GetChildren()
            }
        } catch { }
    }

    if ($null -eq $children) { return }
    foreach ($c in @($children)) {
        if ($null -ne $c) { Collect-PartsRecursive -Comp $c -PartSet $PartSet -SeenModelCfg $SeenModelCfg }
    }
}

function Try-OpenAssembly {
    param(
        [object]$SwApp,
        [string]$Path
    )

    $openErrs = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Path)) {
        $openErrs.Add("Path empty")
        return @{ Model = $null; Errors = $openErrs }
    }

    try {
        $existing = $SwApp.GetOpenDocumentByName($Path)
        if ($existing) {
            return @{ Model = $existing; Errors = @() }
        }
    } catch { }

    try { $SwApp.SetCurrentWorkingDirectory((Split-Path -Path $Path -Parent)) | Out-Null } catch { }

    foreach ($opt in @(3,1,0,2)) {
        $errs = 0
        $warns = 0
        try {
            $m = $SwApp.OpenDoc6($Path, 2, $opt, "", [ref]$errs, [ref]$warns)
            if ($m) { return @{ Model = $m; Errors = @() } }
            $openErrs.Add("OpenDoc6 failed (opt=$opt, err=$errs, warn=$warns)")
        } catch {
            $openErrs.Add("OpenDoc6 exception (opt=$opt): $($_.Exception.Message)")
        }
    }

    try {
        $m2 = $SwApp.OpenDoc($Path, 2)
        if ($m2) { return @{ Model = $m2; Errors = @() } }
        $openErrs.Add("OpenDoc failed")
    } catch {
        $openErrs.Add("OpenDoc exception: $($_.Exception.Message)")
    }

    return @{ Model = $null; Errors = $openErrs }
}

try {
    $fullPath = $AssemblyPath
    try { $fullPath = [System.IO.Path]::GetFullPath($AssemblyPath) } catch { }
    if (-not [string]::IsNullOrWhiteSpace($OutputFile)) {
        try { Remove-Item -Path $OutputFile -Force -ErrorAction SilentlyContinue } catch { }
    }
    Out-Log ("STA helper starting for: " + $fullPath)
    if ([string]::IsNullOrWhiteSpace($fullPath)) {
        throw "Assembly path is empty"
    }

    $sw = $null
    $created = $false
    try {
        try {
            $sw = [Runtime.InteropServices.Marshal]::GetActiveObject("SldWorks.Application")
        } catch { }

        if ($sw) {
            try { $sw.Visible = $true } catch { }
            try { $sw.UserControl = $true } catch { }
            Out-Log "Attached to existing SolidWorks instance in STA helper"
        } else {
            $sw = New-Object -ComObject SldWorks.Application
            $created = $true
            try { $sw.Visible = $true } catch { }
            try { $sw.UserControl = $true } catch { }
            Out-Log "Created SolidWorks instance in STA helper (visible)"
        }
    } catch {
        throw
    }

    $opened = Try-OpenAssembly -SwApp $sw -Path $fullPath
    $model = $opened.Model
    if ($null -eq $model) {
        foreach ($e in @($opened.Errors | Select-Object -First 12)) { Out-Log $e }
        throw "Failed to open assembly in STA helper"
    }

    try { $null = $model.ResolveAllLightWeightComponents($true) } catch { }

    $root = $null
    try {
        $cfg = $model.GetActiveConfiguration()
        if ($cfg) { $root = $cfg.GetRootComponent3($true) }
    } catch { }
    if ($null -eq $root) { throw "Root component unavailable" }

    $partSet = @{}
    $seen = @{}
    Collect-PartsRecursive -Comp $root -PartSet $partSet -SeenModelCfg $seen

    $asmBase = Get-FileBaseNameUpper -PathOrName $fullPath
    if (-not [string]::IsNullOrWhiteSpace($asmBase)) { $partSet[$asmBase] = $true }

    foreach ($pn in @($partSet.Keys | Sort-Object)) {
        Out-Part $pn
    }
    Out-Log ("STA helper finished. Parts: " + $partSet.Count)

    if ($created -and $sw) {
        try { $sw.ExitApp() } catch { }
    } elseif ($model) {
        try {
            $title = $model.GetTitle()
            if (-not [string]::IsNullOrWhiteSpace($title)) { $sw.CloseDoc($title) }
        } catch { }
    }

    if ($sw) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sw) } catch { }
    }
    exit 0
} catch {
    Out-Log ("ERROR: " + $_.Exception.Message)
    exit 2
}
