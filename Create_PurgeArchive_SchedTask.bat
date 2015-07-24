schtasks /Create /TN "PurgeArchive Daily" /TR "Powershell.exe -WindowStyle:Hidden C:\scripts\PurgeArchiveScript.ps1 C:\scripts\PurgeArchiveConfigs.xml" /SC Daily /ST 12:00:00
pause