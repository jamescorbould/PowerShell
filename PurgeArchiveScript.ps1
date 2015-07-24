<#
Author: 				James Corbould
Company:				Datacom Systems NZ Ltd
Purpose: 				Powershell script to archive files on Windows systems.
Dependency List:	.NET 3.5 minimum needs to be installed on the running machine.
							Powershell needs to be installed on the running machine.
Change History:
Version			Date					Who			Change Description
-----------		----------------		-------		----------------------------
1.0.0.0 		14/07/2015				JC			Created.
#>

#================================================
# PARAMETERS
#================================================
[CmdletBinding()]
Param(
	[Parameter(Mandatory=$True)]
	[string]$pFullPurgeArchiveConfigsXMLPath
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
		WriteToEventLog "Application" $_LogSourceName $_ComputerName "Error" 1 [string]::Format("Function CreateLog`n{0}", $error[0])
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
		WriteToEventLog $_LogName $_LogSourceName $_ComputerName "Error" 1 [string]::Format("Function ReturnXMLConfigFile`n{0}", $error[0])
		$xmlFile = $null
	}
	
	return $xmlFile
}

Function DoPurgeAndArchive ($PurgeArchiveConfigXML)
{
	$ProcessReport = "PurgeArchive Process Status Report`n"
	$ProcessReport = $ProcessReport + "------------------------------------------------`n`n"
	$ReadConfig = $true
	
	if ($PurgeArchiveConfigXML -ne $null)
	{
		Write-Host here
		Write-Host $PurgeArchiveConfigXML.toString()
		# Loop through each project config and delete files matching the file mask(s).
		foreach($PC in $PurgeArchiveConfigXML.PurgeArchiveConfigs.PurgeConfig)
		{
			try
			{
				Write-Host here 2
				$ProjectName = $PC.ProjectName
				$ProjectActive = $PC.ProjectActive
				$DirectoryPath = $PC.DirectoryPath
				$KeepDays = $PC.KeepDays
				$DeleteMaskList = $PC.DeleteMasks
				$CreationTimeLimit = (Get-Date).AddDays(-$DeleteDays)
				$FilesCount = 0
				$FilesDeletedCount = 0
				Write-Host projectname $ProjectName
			}
			catch [System.Exception]
			{
				$ProcessReport = $ProcessReport + "Failed - Purge of files older than '" + $KeepDays + " days' from directory '" + $DirectoryPath + "' could not be carried out.`n`n"
				WriteToEventLog $_LogName $_LogSourceName $_ComputerName "Error" 1 [string]::Format("Function DoPurgeAndArchive`n{0}`n{1}", $ProcessReport, $error[0])
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
						if($DeleteMaskList.Length -eq 0)
						{
							$FilesCount = (Get-ChildItem -Path:$DirectoryPath -Recurse -Force).Count
							
							# No delete list specified, so delete **all** files in the specified directory path.
							Get-ChildItem -Path:$DirectoryPath -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $CreationTimeLimit } | Remove-Item -Force
							
							$FilesDeletedCount = $FilesCount - (Get-ChildItem -Path:$DirectoryPath -Recurse -Force).Count
						}
						elseif($DeleteMaskList.Length -ne 0)
						{
							$FilesCount = (Get-ChildItem -Path:$DirectoryPath -Recurse -Force).Count
							
							# Delete list specified, so only delete those files that match the file mask.
							Get-ChildItem -Path:$DirectoryPath -Include:$DeleteMaskList -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $CreationTimeLimit } | Remove-Item -Force
							
							$FilesDeletedCount = $FilesCount - (Get-ChildItem -Path:$DirectoryPath -Recurse -Force).Count
						}
						
						$ProcessReport = $ProcessReport + [string]::Format("Success - Purged {0} files older than {1} days' from directory '{2}'.`n`n", $FilesDeletedCount, $KeepDays, $DirectoryPath)
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

	$ProcessReport = $ProcessReport + "`n`n------------------------------------------------`n"
	$ProcessReport = $ProcessReport + "End of Report"

	WriteToEventLog $_LogName $_LogSourceName $_ComputerName "Information" 1 $ProcessReport
}

#================================================
# MAIN EXECUTION HERE
#================================================

# Call local func to create event log.
CreateLog $_LogName $_LogSourceName $_ComputerName > $null

$xmlConfigFile = GetXMLConfigFile
DoPurgeAndArchive $xmlConfigFile
