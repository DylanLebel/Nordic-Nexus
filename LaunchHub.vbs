Set objShell = CreateObject("WScript.Shell")
' Get the current directory of the VBS script
strPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
' Launch PowerShell hidden
objShell.Run "powershell.exe -ExecutionPolicy Bypass -Sta -WindowStyle Hidden -File """ & strPath & "\HubService.ps1""", 0, False
