#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------


#Sample function that provides the location of the script
function Get-ScriptDirectory
{
<#
	.SYNOPSIS
		Get-ScriptDirectory returns the proper location of the script.

	.OUTPUTS
		System.String
	
	.NOTES
		Returns the correct path within a packaged executable.
#>
	[OutputType([string])]
	param ()
	if ($null -ne $hostinvocation)
	{
		Split-Path $hostinvocation.MyCommand.path
	}
	else
	{
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}

Function Export-OutlookEmailsByCategory
{
	[CmdletBinding()]
	Param (
		[Parameter(Mandatory = $true)]
		[string]$CategoryName,
		[Parameter(Mandatory = $true)]
		[string]$ExportPath,
		[Parameter(Mandatory = $false)]
		[string]$PSTFolderPath,
		[Parameter(Mandatory = $false)]
		[string]$FolderName,
		[Parameter(Mandatory = $false)]
		[switch]$UsePST,
		[Parameter(Mandatory = $false)]
		[switch]$Force
	)
	
	try
	{
		# Check for running Outlook processes
		if ($Force)
		{
			$outlookProcesses = Get-Process Outlook -ErrorAction SilentlyContinue
			if ($outlookProcesses)
			{
				$response = Read-Host "Outlook is running. Do you want to close it? [Y/N]"
				if ($response -eq 'Y')
				{
					Write-Host "Closing Outlook..."
					$outlookProcesses | Stop-Process
					Start-Sleep -Seconds 3 # Wait for Outlook to close completely
				}
				else
				{
					Write-Host "Script execution stopped by user."
					return
				}
			}
		}
		
		$Outlook = New-Object -ComObject Outlook.Application
		Write-Host "Outlook Application started successfully."
		
		$Namespace = $Outlook.GetNamespace("MAPI")
		$Categories = $Outlook.Session.Categories
		
		# Determine which folder to use
		$Folder = $null
		if ($UsePST)
		{
			Write-Host "PST folder path provided: $PSTFolderPath"
			$PSTTopFolder = $Namespace.Folders.Item($PSTFolderPath)
			if ($null -eq $PSTTopFolder)
			{
				Write-Host "PST folder not found: $PSTFolderPath"
				return
			}
			if (-not $FolderName)
			{
				Write-Host "Folder name is required when using the -UsePST switch."
				return
			}
			$Folder = $PSTTopFolder.Folders.Item($FolderName)
			if ($null -eq $Folder)
			{
				Write-Host "Subfolder not found in PST: $FolderName"
				return
			}
		}
		else
		{
			Write-Host "Using default Inbox."
			$Folder = $Namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)
		}
		
		# Check if the category exists
		$CategoryExists = $false
		foreach ($Category in $Categories)
		{
			if ($Category.Name -eq $CategoryName)
			{
				$CategoryExists = $true
				break
			}
		}
		
		if (-not $CategoryExists)
		{
			Write-Host "Category '$CategoryName' does not exist."
			return
		}
		
		
		$Filter = "[Categories] = '$CategoryName'"
		$FilteredEmails = $Folder.Items.Restrict($Filter)
		
		if (-not (Test-Path -Path $ExportPath))
		{
			New-Item -ItemType Directory -Path $ExportPath -Force
			Write-Host "Directory created at $ExportPath"
		}
		
		if (-not $FilteredEmails)
		{
			Write-Host "No emails found in the specified category."
			return
		}
		
		foreach ($Email in $FilteredEmails)
		{
			
			if (-not $Email)
			{
				Write-Host "Encountered a null email object, skipping..."
				continue
			}
			
			if (-not $Email.Subject -or -not $Email.ReceivedTime)
			{
				Write-Host "Email subject or received time is null, skipping..."
				continue
			}
			$safeSubject = $Email.Subject -replace '[\\\/\:\*\?\"\<\>\|]+', '_'
			$receivedTime = $Email.ReceivedTime.ToString("yyyy-MM-dd_HH-mm-ss")
			$filename = "$safeSubject-$receivedTime.msg"
			$filePath = Join-Path $ExportPath $filename
			
			if (Test-Path -Path $filePath)
			{
				$overwrite = Read-Host "$filePath already exists. Do you want to overwrite it? [Y/N]"
				if ($overwrite -ne 'Y')
				{
					Write-Host "Skipping file: $filePath"
					continue
				}
			}
			
			$Email.SaveAs($filePath, [Microsoft.Office.Interop.Outlook.OlSaveAsType]::olMSG)
		}
		
		Write-Host "Export complete. Emails saved in $ExportPath"
	}
	catch
	{
		Write-Host "An error occurred. Error: $_"
	}
	finally
	{
		
		
		# Clean up Outlook COM object
		if ($null -ne $Folder)
		{
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Folder) | Out-Null
		}
		if ($null -ne $Namespace)
		{
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Namespace) | Out-Null
		}
		if ($null -ne $Outlook)
		{
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
		}
		[System.GC]::Collect()
		[System.GC]::WaitForPendingFinalizers()
		
		# Close the Outlook process if the script opened it with -Force
		if ($Force)
		{
			Get-Process Outlook -ErrorAction SilentlyContinue | Stop-Process
		}
	}
}

#Sample variable that provides the location of the script
[string]$ScriptDirectory = Get-ScriptDirectory
