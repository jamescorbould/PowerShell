schtasks /Create /TN "Archive Files Daily" /TR "C:\scripts\PurgeArchiveFiles\ExecuteArchiveScript.vbs" /SC Daily /ST 02:00:00
schtasks /Create /TN "Purge Files Daily" /TR "C:\scripts\PurgeArchiveFiles\ExecutePurgeScript.vbs" /SC Daily /ST 22:00:00
pause