@echo off
setlocal
cd /d "%~dp0"
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Setup-EpicorCredentials.ps1"
endlocal
