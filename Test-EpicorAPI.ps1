# ==============================================================================
#  Test-EpicorAPI.ps1 - Nordic Minesteel Technologies
#  Verifies Epicor REST API connectivity and part/revision lookup.
#  Reads credentials from config.json.
# ==============================================================================

param(
    [string]$PartNum    = "1206-P404",
    [string]$ConfigPath = "config.json"
)

$scriptDir = Split-Path $PSCommandPath -Parent
$cfgPath   = if ([System.IO.Path]::IsPathRooted($ConfigPath)) { $ConfigPath } else { Join-Path $scriptDir $ConfigPath }
$cfg       = Get-Content $cfgPath -Raw | ConvertFrom-Json
$ec        = $cfg.epicor

function Get-StoredEpicorCredential {
    $credPath = Join-Path (Join-Path $env:LOCALAPPDATA "EpicorOrderMonitor") "epicor-creds.xml"
    if (-not (Test-Path $credPath)) { return $null }
    try {
        $stored = Import-Clixml -Path $credPath
        if ($stored.PSObject.Properties.Name -contains "Username" -and $stored.PSObject.Properties.Name -contains "Password") {
            $userSecure = ConvertTo-SecureString -String $stored.Username
            $passSecure = ConvertTo-SecureString -String $stored.Password
            $username = [System.Net.NetworkCredential]::new("", $userSecure).Password
            $password = [System.Net.NetworkCredential]::new("", $passSecure).Password
            return [pscustomobject]@{
                UserName = $username
                Password = $password
            }
        }
        if ($stored -is [pscredential]) {
            return [pscustomobject]@{
                UserName = $stored.UserName
                Password = $stored.GetNetworkCredential().Password
            }
        }
        return $null
    } catch {
        return $null
    }
}

if (-not $ec) { Write-Host "No 'epicor' block in config.json" -ForegroundColor Red; exit 1 }
$username = $ec.username
$password = $ec.password
$storedCred = Get-StoredEpicorCredential
if ($storedCred) {
    if (($username -in @("YOUR_EPICOR_USERNAME", "YOUR_USERNAME_HERE")) -or [string]::IsNullOrWhiteSpace($username)) {
        $username = $storedCred.UserName
    }
    if (($password -in @("YOUR_EPICOR_PASSWORD", "YOUR_PASSWORD_HERE")) -or [string]::IsNullOrWhiteSpace($password)) {
        $password = $storedCred.Password
    }
}
if (($username -in @("YOUR_EPICOR_USERNAME", "YOUR_USERNAME_HERE")) -or [string]::IsNullOrWhiteSpace($username)) {
    Write-Host "Set your username in config.json or run .\Setup-EpicorCredentials.ps1 first." -ForegroundColor Red; exit 1
}
if (($password -in @("YOUR_EPICOR_PASSWORD", "YOUR_PASSWORD_HERE")) -or [string]::IsNullOrWhiteSpace($password)) {
    Write-Host "Set your password in config.json or run .\Setup-EpicorCredentials.ps1 first." -ForegroundColor Red; exit 1
}

$base = $ec.apiUrl.TrimEnd('/')
$b64  = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${username}:${password}"))
$h    = @{ Authorization = "Basic $b64"; Accept = "application/json"; "x-epicor-company" = $ec.company }

Write-Host "=== Epicor API Test ===" -ForegroundColor Cyan
Write-Host "URL     : $base"
Write-Host "Company : $($ec.company)"
Write-Host "User    : $username"
Write-Host "Part    : $PartNum"

# --- 1. Connectivity ping ---
Write-Host "`n[1] API connectivity ping..." -ForegroundColor Yellow
try {
    Invoke-RestMethod -Uri "$base/api/v1" -Headers $h -Method Get -TimeoutSec 10 -ErrorAction Stop | Out-Null
    Write-Host "    OK - v1 API reachable" -ForegroundColor Green
} catch {
    Write-Host "    FAIL [$($_.Exception.Response.StatusCode.value__)]: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# --- 2. Part existence ---
Write-Host "`n[2] Part master lookup for '$PartNum'..." -ForegroundColor Yellow
try {
    $pUrl     = "$base/api/v1/Erp.BO.PartSvc/Parts?`$filter=PartNum eq '$PartNum'&`$select=PartNum,PartDescription,TypeCode&`$top=1"
    $partResp = Invoke-RestMethod -Uri $pUrl -Headers $h -Method Get -TimeoutSec 15 -ErrorAction Stop
    # v1 OData returns value[] for collections; single result may be direct object
    $partRows = if ($partResp.value -and @($partResp.value).Count -gt 0) { @($partResp.value) }
                elseif ($partResp.PartNum) { @($partResp) }
                else { @() }
    if ($partRows.Count -gt 0) {
        Write-Host "    FOUND: $($partRows[0].PartNum) - $($partRows[0].PartDescription) (Type: $($partRows[0].TypeCode))" -ForegroundColor Green
    } else {
        Write-Host "    NOT FOUND via OData filter (GetByID will still work)" -ForegroundColor Yellow
    }
} catch {
    Write-Host "    FAIL: $($_.Exception.Message)" -ForegroundColor Red
}

# --- 3. Revision lookup via GetByID (returns full dataset including PartRev child rows) ---
Write-Host "`n[3] Revision lookup via GetByID for '$PartNum'..." -ForegroundColor Yellow
Write-Host "    Note: PartRevs OData entity is always empty - must use GetByID + returnObj.PartRev" -ForegroundColor DarkGray
try {
    $pnEnc   = [Uri]::EscapeDataString($PartNum)
    $gdResp  = Invoke-RestMethod -Uri "$base/api/v1/Erp.BO.PartSvc/GetByID?partNum=$pnEnc" `
        -Headers $h -Method Get -TimeoutSec 15 -ErrorAction Stop
    $revRows = @($gdResp.returnObj.PartRev)
    if ($revRows.Count -gt 0) {
        Write-Host "    Revisions found:" -ForegroundColor Green
        $revRows | ForEach-Object {
            $status = if ($_.Approved) { "Approved" } else { "Not Approved" }
            Write-Host "      Rev $($_.RevisionNum)  [$status]  Effective: $($_.EffectiveDate.ToString().Substring(0,10))  By: $($_.ApprovedBy)" -ForegroundColor Green
        }
        # Determine latest approved rev
        $approved = @($revRows | Where-Object { $_.Approved })
        $latest   = ($approved + $revRows) | Sort-Object { try { [int]($_.RevisionNum -replace '[^0-9]','') } catch { 0 } } -Descending | Select-Object -First 1
        Write-Host "    Latest approved rev: $($latest.RevisionNum)" -ForegroundColor Cyan
    } else {
        Write-Host "    No revisions found in Epicor for this part" -ForegroundColor Yellow
    }
} catch {
    Write-Host "    FAIL: $($_.Exception.Message)" -ForegroundColor Red
}

# --- 4. Quick rev match test ---
Write-Host "`n[4] Rev match simulation (order rev vs Epicor)..." -ForegroundColor Yellow
if ($revRows -and $revRows.Count -gt 0) {
    $latestRev = ($revRows | Where-Object { $_.Approved } | Sort-Object { try { [int]($_.RevisionNum -replace '[^0-9]','') } catch { 0 } } -Descending | Select-Object -First 1).RevisionNum
    Write-Host "    If order rev = $latestRev  → MATCH  (Status: Complete)" -ForegroundColor Green
    $otherRev = [string]([int]$latestRev + 1)
    Write-Host "    If order rev = $otherRev  → MISMATCH  (Status: Rev Mismatches Found)" -ForegroundColor Yellow
}

Write-Host "`n=== All tests complete ===" -ForegroundColor Cyan
