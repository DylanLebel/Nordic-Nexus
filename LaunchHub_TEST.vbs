Set objShell = CreateObject("WScript.Shell")
' Get the directory this VBS lives in
strPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
' Launch HubService in TEST MODE (hidden window, no console)
objShell.Run "powershell.exe -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File """ & strPath & "\HubService.ps1"" -TestMode", 0, False
