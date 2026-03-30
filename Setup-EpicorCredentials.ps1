# ==============================================================================
#  Setup-EpicorCredentials.ps1 - Nordic Minesteel Technologies
#  Prompts for Epicor credentials and stores them encrypted with Windows DPAPI.
#  The saved file can only be decrypted by the same Windows user on this machine.
# ==============================================================================

$credFolder = Join-Path $env:LOCALAPPDATA "EpicorOrderMonitor"
$credPath   = Join-Path $credFolder "epicor-creds.xml"

Write-Host ""
Write-Host "=== Epicor Credential Setup ===" -ForegroundColor Cyan
Write-Host "Credentials will be stored here:" -ForegroundColor Gray
Write-Host "  $credPath" -ForegroundColor Yellow
Write-Host "They are encrypted with Windows DPAPI and tied to this Windows account." -ForegroundColor Gray
Write-Host ""

$username = Read-Host "Epicor username"
if ([string]::IsNullOrWhiteSpace($username)) {
    Write-Host "No username entered. Nothing was saved." -ForegroundColor Red
    exit 1
}

$password = Read-Host "Epicor password" -AsSecureString
if (-not $password) {
    Write-Host "No password entered. Nothing was saved." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $credFolder)) {
    New-Item -ItemType Directory -Path $credFolder -Force | Out-Null
}

$userSecure = ConvertTo-SecureString -String $username -AsPlainText -Force
$payload = [pscustomobject]@{
    Username = ConvertFrom-SecureString -SecureString $userSecure
    Password = ConvertFrom-SecureString -SecureString $password
}
$payload | Export-Clixml -Path $credPath -Force

Write-Host ""
Write-Host "[OK] Credentials saved." -ForegroundColor Green
Write-Host "Username and password were both saved in protected form." -ForegroundColor Gray
Write-Host "The monitor and API test will load them automatically." -ForegroundColor Cyan
Write-Host ""
