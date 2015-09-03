' Execute PowerShell command silently.
command = "Powershell.exe -WindowStyle Hidden C:\scripts\PurgeArchiveFiles\PurgeArchiveScript.ps1 C:\scripts\PurgeArchiveFiles\PurgeArchiveConfigs.xml A"
set shell = CreateObject("WScript.Shell")
shell.Run command,0