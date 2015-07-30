<#
.NOTES
	Author: 					James Corbould
	Company:				Datacom Systems NZ Ltd
	Purpose: 				Powershell script to archive files on Windows systems.
	Dependency List:	.NET 3.5 minimum needs to be installed on the running machine.
									Powershell needs to be installed on the running machine.
.CHANGE_HISTORY:
Version		Date					Who		Change Description
-----------		----------------		-------		----------------------------
1.0.0.0 		14/07/2015		JC			Created.
#>

#================================================
# PARAMETERS
#================================================
[CmdletBinding()]
Param(
	[Parameter(Mandatory=$True)]
	[string]$pFullPurgeArchiveConfigsXMLPath,  # Path to XML config file.
	[Parameter(Mandatory=$True)]
	[string]$pPurgeOrArchiveSwitch  # Switch indicating if we should purge or archive.
)

#================================================
# GLOBAL VARS
#================================================
$_LogName = "PurgeArchiveSchTask"
$_LogSourceName = "Purge Archive Script"
$_ComputerName = $env:COMPUTERNAME

#================================================
# FUNCTIONS
#================================================
Function WriteToEventLog ($log, $source, $computername, $type, $eventid, $message)
{
	try
	{
		$success = $true
		CreateLog $log $source $computername
		Write-EventLog -LogName $log -EntryType Error -EventId 1 -Message $message.toString() -Source $source.toString() -ComputerName $computername.toString()
	}
	catch [System.Exception]
	{
		$errMessage = [string]::Format("Function WriteToEventLog`n{0}", $error[0])
		Write-EventLog -LogName Application -EntryType Error -EventId 1 -Message $errMessage.toString() -Source $source.toString() -ComputerName $computername.toString()
		$success = $false
	}
	
	return $success
}

Function CreateLog ($log, $source, $computername)
{
	try
	{
		$success = $true
		
		if (-not [System.Diagnostics.EventLog]::SourceExists($source))  # Returns false if log source exists.
		{
			New-EventLog -LogName $log.toString() -Source $source.toString() -ComputerName $computername.toString()
			
			# Create source in Application event log too, in case of errors writing to the purge archive event log we can default to the Application log instead.
			New-EventLog -LogName Application -Source $source.toString() -ComputerName $computername.toString()
		}
		else
		{
			# Do nothing - the log source already exists.
		}
	}
	catch [System.Exception]
	{
		$errMessage = [string]::Format("Function CreateLog`n{0}", $error[0])
		WriteToEventLog "Application" $_LogSourceName $_ComputerName "Error" 1 $errMessage
		$success = $false
	}
	
	return $success
}

Function GetXMLConfigFile
{
	try
	{
		$xmlFile = [XML](Get-Content -ErrorAction:Stop -Path:$pFullPurgeArchiveConfigsXMLPath)
	}
	catch [System.Exception]
	{
		$errMessage = [string]::Format("Function ReturnXMLConfigFile`n{0}", $error[0])
		WriteToEventLog $_LogName $_LogSourceName $_ComputerName "Error" 1 $errMessage
		$xmlFile = $null
	}
	
	return $xmlFile
}

Function GetFilesOlderThanXDays ($DirectoryPath, $CreationDateLimit)
{
	$files = $null
	
	try
	{
		[Array]$files = Get-ChildItem -Path:$DirectoryPath -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -le $CreationDateLimit } 
	}
	catch [System.Exception]
	{
		$errMessage = [string]::Format("Function GetFilesOlderThanXDays`n{0}", $error[0])
		WriteToEventLog $_LogName $_LogSourceName $_ComputerName "Error" 1 $errMessage
		$files = $null
	}
	
	return $files
}

Function Create7ZipFile ($DirectoryPath, $ZipFileName, $ArchiveMaskArray, $CreationDateLimit)
{
	try
	{
		$success = $true
		$ZipArchiveDestination =  [string]::Format("`"{0}\{1}`"", $DirectoryPath, $ZipFileName)
		$SourceFiles = ""
		$filesToArchive = $null
		
		# Get all files in the specified directory older than the date limit for archival.
		[Array]$filesToArchive = GetFilesOlderThanXDays $DirectoryPath $CreationDateLimit
		
		if ($filesToArchive.Count -gt 1)
		{
			if ($ArchiveMaskArray.Count -le 1)
			{
				# No file masks specified - assume all files in the directory should be added to the compressed file.
				# Equivalent to the file mask *.*.
				
				foreach ($file in $filesToArchive)
				{
					$SourceFiles = $SourceFiles + [string]::Format("{0}\{1} ", $DirectoryPath, $file)
				}
			}
			else
			{
				# File mask(s) have been specified: only archive files that match the file mask.
				
				foreach ($mask in $ArchiveMaskArray)
				{
					foreach ($file in $filesToArchive)
					{
						$extension = [System.IO.Path]::GetExtension($file)
						
						if ([string]$extension.ToUpper() -match $mask.ToUpper())
						{
							$SourceFiles = $SourceFiles + [string]::Format("`"{0}\{1}`" ", $DirectoryPath, $file)
							
							# Delete this file from the file array to reduce processing time O(n^2)
							# TODO
						}
					}
				}
			}

			$result = & ".\7za.exe" a -tzip $ZipArchiveDestination $SourceFiles -r -mmt
			
			if (-not [string]$result.ToUpper() -match 'EVERYTHING IS OK')
			{
				$errMessage = [string]::Format("Function Create7ZipFile`nFailed to zip files using 7zip.")
				WriteToEventLog $_LogName $_LogSourceName $_ComputerName "Error" 1 $errMessage
				$success = $false
			}
		}
		else
		{
			# No files to archive.
			
		}
	}
	catch [System.Exception]
	{
		$errMessage = [string]::Format("Function Create7ZipFile`n{0}", $error[0])
		WriteToEventLog $_LogName $_LogSourceName $_ComputerName "Error" 1 $errMessage
		$success = $false
	}
	
	return $success
}

Function DoPurge ($PurgeArchiveConfigXML)
{
	$ProcessReport = [string]::Format("Purge Process Status Report for Computer {0}`n", $_ComputerName)
	$ProcessReport = $ProcessReport + "------------------------------------------------------------------------------`n`n"
	$ReadConfig = $true
	
	if ($PurgeArchiveConfigXML -ne $null)
	{
		# Loop through each project config and delete files matching the file mask(s).
		foreach($PC in $PurgeArchiveConfigXML.PurgeArchiveConfigs.PurgeConfig)
		{
			try
			{
				$ProjectName = $PC.ProjectName
				$ProjectActive = $PC.ProjectActive
				$DirectoryPath = $PC.DirectoryPath
				$KeepDays = $PC.KeepDays
				$DeleteMaskList = $PC.DeleteMasks
				$DeleteMaskArray = $DeleteMaskList -split ";"
				$CreationTimeLimit = (Get-Date).AddDays(-$DeleteDays)
				$FilesCount = 0
				$FilesDeletedCount = 0
			}
			catch [System.Exception]
			{
				$ProcessReport = $ProcessReport + "Failed - Purge of files older than '" + $KeepDays + " days' from directory '" + $DirectoryPath + "' could not be carried out.`n`n"
				WriteToEventLog $_LogName $_LogSourceName $_ComputerName "Error" 1 [string]::Format("Function DoPurge`n{0}`n{1}", $ProcessReport, $error[0])
				$ReadConfig = $false
			}
			
			$ProcessReport = $ProcessReport + [string]::Format("Project Name: {0}`nActive: {1}`nPurgeArchive Status: ", $ProjectName, $ProjectActive.ToUpper())
			
			if($ProjectActive.ToUpper() -eq "TRUE" -and $ReadConfig -ne $false)
			{
				$DirectoryPathExists = Test-Path $DirectoryPath -PathType:Container
				
				if($DirectoryPathExists -eq $true)
				{
					try
					{
						if($DeleteMaskArray.Count -eq 0)
						{
							$FilesCount = (Get-ChildItem -Path:$DirectoryPath -Recurse -Force).Count
							
							# No delete list specified, so delete **all** files in the specified directory path.
							Get-ChildItem -Path:$DirectoryPath -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -le $CreationTimeLimit } | Remove-Item -Force
							
							$FilesDeletedCount = $FilesCount - (Get-ChildItem -Path:$DirectoryPath -Recurse -Force).Count
						}
						elseif($DeleteMaskArray.Count -ne 0)
						{
							$FilesCount = (Get-ChildItem -Path:$DirectoryPath -Recurse -Force).Count
							
							# Delete list specified, so only delete those files that match the file mask.
							Get-ChildItem -Path:$DirectoryPath -Include:$DeleteMaskArray -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -le $CreationTimeLimit } | Remove-Item -Force
							
							$FilesDeletedCount = $FilesCount - (Get-ChildItem -Path:$DirectoryPath -Recurse -Force).Count
						}
						
						$ProcessReport = $ProcessReport + [string]::Format("Success - Purged {0} file(s) older than {1} days' from directory '{2}'.`n`n", $FilesDeletedCount, $KeepDays, $DirectoryPath)
					}
					catch [System.Exception]
					{
						$ProcessReport = $ProcessReport + [string]::Format("Failed - Purge of files older than {0} days' from directory '{1}' could not be carried out.`n`n", $KeepDays, $DirectoryPath)
					}
				}
				elseif($DirectoryPathExists -eq $FALSE)
				{
					$ProcessReport = $ProcessReport + [string]::Format("Failed - Purge of files older than {0} days' from directory '{1}' could not be carried out.  Directory path does not exist.`n`n", $KeepDays, $DirectoryPath)
				}
				
				$DeleteMaskList = ""
				$ExcludeMaskList = ""
			}
			elseif($ProjectActive.ToUpper() -eq "FALSE")
			{
				$ProcessReport = $ProcessReport + [string]::Format("N/A - Purge of files **not** carried out from directory '{0}' as the project '{1}' had been configured as inactive in the XML configuration file.`n`n", $DirectoryPath, $ProjectName)
			}
		}
	}
	else
	{
		$ProcessReport = $ProcessReport + "Unexpected fatal error: The 'PurgeArchive' configuration XML file cannot be read."
	}

	$ProcessReport = $ProcessReport + "------------------------------------------------------------------------------`n"
	$ProcessReport = $ProcessReport + "End of Report"

	WriteToEventLog $_LogName $_LogSourceName $_ComputerName "Information" 1 $ProcessReport
}

Function DoArchive ($PurgeArchiveConfigXML)
{
	$ProcessReport = [string]::Format("Archive Process Status Report for Computer {0}`n", $_ComputerName)
	$ProcessReport = $ProcessReport + "------------------------------------------------------------------------------`n`n"
	$ReadConfig = $true
	
	if ($PurgeArchiveConfigXML -ne $null)
	{
		# Loop through each project config and archive files in the specified directory.
		foreach($PC in $PurgeArchiveConfigXML.PurgeArchiveConfigs.ArchiveConfig)
		{
			try
			{
				$ProjectName = $PC.ProjectName
				$ProjectActive = $PC.ProjectActive
				$DirectoryPath = $PC.DirectoryPath
				$KeepDays = $PC.KeepDays
				$ArchiveMaskList = $PC.ArchiveMasks
				$ArchiveMaskArray = $ArchiveMaskList -split ";"
				$CreationTimeLimit = (Get-Date).AddDays(-$KeepDays)
				$FilesCount = 0
				$FilesArchivedCount = 0
				$DateTime = [string](Get-Date -Format "yyyy-MM-dd-hh_mm_ss")
				$ZipFileName = [string]::Format("{0}-{1}{2}", $PC.ProjectName, $DateTime, ".zip")
			}
			catch [System.Exception]
			{
				$ProcessReport = $ProcessReport + "Failed - Archive of files older than '" + $KeepDays + " days' from directory '" + $DirectoryPath + "' could not be carried out.`n`n"
				WriteToEventLog $_LogName $_LogSourceName $_ComputerName "Error" 1 [string]::Format("Function DoArchive`n{0}`n{1}", $ProcessReport, $error[0])
				$ReadConfig = $false
			}
			
			$ProcessReport = $ProcessReport + [string]::Format("Project Name: {0}`nActive: {1}`nFilename: {2}`nPurgeArchive Status: ", $ProjectName, $ProjectActive.ToUpper(), $ZipFileName)
			
			if($ProjectActive.ToUpper() -eq "TRUE" -and $ReadConfig -ne $false)
			{
				$DirectoryPathExists = Test-Path $DirectoryPath -PathType:Container
				
				if($DirectoryPathExists -eq $true)
				{
					try
					{
						$success = Create7ZipFile $DirectoryPath $ZipFileName $ArchiveMaskArray $CreationTimeLimit
						
						if (-not $success)
						{
							throw [System.Exception]
						}
						
						$ProcessReport = $ProcessReport + [string]::Format("Success - Archived file(s) older than {0} days' from directory '{1}'.`n`n", $KeepDays, $DirectoryPath)
					}
					catch [System.Exception]
					{
						$ProcessReport = $ProcessReport + [string]::Format("Failed - Archive of files older than {0} days' from directory '{1}' could not be carried out.`n`n", $KeepDays, $DirectoryPath)
					}
				}
				elseif($DirectoryPathExists -eq $FALSE)
				{
					$ProcessReport = $ProcessReport + [string]::Format("Failed - Archive of files older than {0} days' from directory '{1}' could not be carried out.  Directory path does not exist.`n`n", $KeepDays, $DirectoryPath)
				}
			}
			elseif($ProjectActive.ToUpper() -eq "FALSE")
			{
				$ProcessReport = $ProcessReport + [string]::Format("N/A - Purge of files **not** carried out from directory '{0}' as the project '{1}' had been configured as inactive in the XML configuration file.`n`n", $DirectoryPath, $ProjectName)
			}
		}
	}
	else
	{
		$ProcessReport = $ProcessReport + "Unexpected fatal error: The 'PurgeArchive' configuration XML file cannot be read."
	}

	$ProcessReport = $ProcessReport + "------------------------------------------------------------------------------`n"
	$ProcessReport = $ProcessReport + "End of Report"

	WriteToEventLog $_LogName $_LogSourceName $_ComputerName "Information" 1 $ProcessReport
}

#================================================
# MAIN EXECUTION HERE
#================================================

# Call local func to create event log.
CreateLog $_LogName $_LogSourceName $_ComputerName > $null

$xmlConfigFile = GetXMLConfigFile

if ($pPurgeOrArchiveSwitch.ToUpper() -eq "P")
{
	DoPurge $xmlConfigFile > $null
}
elseif ($pPurgeOrArchiveSwitch.ToUpper() -eq "A")
{
	DoArchive $xmlConfigFile > $null
}
else
{
	$errMessage = [string]::Format("Fatal - `"{0}`" switch is not recognised for parameter `"pPurgeOrArchiveSwitch`".", $pPurgeOrArchiveSwitch)
	WriteToEventLog $_LogName $_LogSourceName $_ComputerName "Error" 1 $errMessage > $null
}