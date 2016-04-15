<# 
.SYNOPSIS 
	Clears identified folders if target disk is below specified threshold.
.DESCRIPTION 
	Clears out folders identified in the "Disk Cleanup Wizard". In addition, will clean out Google Chrome cache and SMS Cache files. Can produce email report
.NOTES 
    File Name  : HardDriveCleaner.ps1
    Author     : Brenton keegan - brenton.keegan@gmail.com 
    Licenced under GPLv3  
.LINK 
	https://github.com/bkeegan/HardDriveCleaner
    License: http://www.gnu.org/copyleft/gpl.html
.EXAMPLE 
	HardDriveCleaner -t 2000 -sms -tmp -ie -utmp -chrome -owp -dpf -to "alerts@contoso.com" -from "alerts@contoso.com" -smtp "mail.contoso.com"

#> 

#============================================HELPER FUNCTION
Function CleanSystemFolder
{
	[cmdletbinding()]
	Param
	(	
		[parameter(Mandatory=$true)]
		[alias("t")]
		[string]$targetFolder,
		
		[parameter(Mandatory=$true)]
		[alias("obj")]
		$reportingObject,
		
		[parameter(Mandatory=$true)]
		[alias("n")]
		[string]$itemName	
	)
	
	
	$originalSize = (Get-ChildItem -r $targetFolder | Measure-Object -property length -sum)
	$originalSizeMB = "{0:N2}" -f ($originalSize.sum / 1MB) + " MB"
	$reportingObject.Add("$itemName - Original Size", $originalSizeMB)
	Remove-Item $targetFolder\* -recurse -force -ErrorAction SilentlyContinue

	$newSize = (Get-ChildItem -r $targetFolder | Measure-Object -property length -sum)
	$newSizeMB = "{0:N2}" -f ($newSize.sum / 1MB) + " MB"
	$reportingObject.Add("$itemName - New Size", $newSizeMB)

}
#============================================HELPER FUNCTION
Function CleanProfileFolder
{
	[cmdletbinding()]
	Param
	(
		[parameter(Mandatory=$true)]
		[alias("u")]
		[string]$usersFolder,
		
		[parameter(Mandatory=$true)]
		[alias("t")]
		[string]$targetFolder,
		
		[parameter(Mandatory=$true)]
		[alias("obj")]
		$reportingObject,
		
		[parameter(Mandatory=$true)]
		[alias("n")]
		[string]$itemName	
	)

	$userFolders = get-childitem $usersFolder
	foreach($userFolder in $userFolders)
	{
		$fullTargetPath = "$($userFolder.FullName)$targetFolder"
		if(Test-Path $fullTargetPath)
		{
			$targetItems = Get-ChildItem -r -force $fullTargetPath | Measure-Object -property length -sum
			$targetItemsSizeTotal += $targetItems.sum 
			Remove-Item $fullTargetPath\* -Recurse -Force -ErrorAction SilentlyContinue
			$targetItemsNew = Get-ChildItem -r -force $fullTargetPath | Measure-Object -property length -sum
			$targetItemsSizeTotalNew += $targetItemsNew.sum 
		}
	}
	$targetItemsSizeTotal  = "{0:N2}" -f ($targetItemsSizeTotal / 1MB) + " MB"
	$targetItemsSizeTotalNew  = "{0:N2}" -f ($targetItemsSizeTotalNew / 1MB) + " MB"
	$reportingObject.Add("$itemName - Original Size", $targetItemsSizeTotal)
	$reportingObject.Add("$itemName - New Size", $targetItemsSizeTotalNew)
}

#==================================================MAIN FUNCTION
Function HardDriveCleaner
{
	[cmdletbinding()]
	Param
	(
		[parameter(Mandatory=$true)]
		[alias("t")]
		[int]$lowDiskThreshold,
		
		[parameter(Mandatory=$false)]
		[alias("sms")]
		[switch]$cleanSMSCache,
		
		[parameter(Mandatory=$false)]
		[alias("tmp")]
		[switch]$cleanTemp,
		
		[parameter(Mandatory=$false)]
		[alias("owp")]
		[switch]$cleanOfflineWeb,

		[parameter(Mandatory=$false)]
		[alias("dpf")]
		[switch]$cleanDownloadProgFiles,
		
		[parameter(Mandatory=$false)]
		[alias("prof")]
		[switch]$cleanProfiles,
		
		[parameter(Mandatory=$false)]
		[alias("ie")]
		[switch]$ieCacheAllUsers,
		
		[parameter(Mandatory=$false)]
		[alias("chrome")]
		[switch]$googleCacheAllUsers,
		
		[parameter(Mandatory=$false)]
		[alias("utmp")]
		[switch]$allUsersTempFolders,
		
		[parameter(Mandatory=$false)]
		[alias("noemail")]
		[switch]$doNotSendEmail,
		
		[parameter(Mandatory=$true)]
		[alias("To")]
		[string]$emailRecipient,
		
		[parameter(Mandatory=$true)]
		[alias("From")]
		[string]$emailSender,
		
		[parameter(Mandatory=$true)]
		[alias("smtp")]
		[string]$emailServer,
		
		[parameter(Mandatory=$false)]
		[alias("d")]
		[string]$targetDrive="C:",
		
		[parameter(Mandatory=$false)]
		[alias("dp")]
		[string]$delProf2Location="C:\Scripts"
		
	)
	#=====================================================================VARIABLE INITS
	#FILE LOCATIONS - used in constructing paths
	$ieCache = "\AppData\Local\Microsoft\Windows\INetCache\IE" #reletive to user profile
	$googleChromeCache = "\AppData\Local\Google\Chrome\User Data\Default\Cache" #reletive to user profile
	$userTmp = "\AppData\Local\Temp" #reletive to user profile
	
	$usersFolder = "\Users" #reletive to drive target
	$systemTemp = "\WIndows\Temp" #reletive to drive target
	$smsCache = "\Windows\ccmcache" #reletive to drive target
	$offlineWebPages = "\Windows\Offline Web Pages"
	$downloadedProgramFiles = "\Windows\Downloaded Program Files"
	
	
	#FILE NAME VARIABLES - saves temporary html report to to attach to  email
	[string]$dateStamp = Get-Date -UFormat "%Y%m%d_%H%M%S" #timestamp for naming report
	$tempFolder = get-item env:temp #temp folder
	
	#INITIAL INFO
	$thresholdInBytes = 1MB * $lowDiskThreshold
	$disk = Get-WmiObject Win32_LogicalDisk -ComputerName . -Filter "DeviceID='$targetDrive'" | Select-Object Size,FreeSpace
	$freeSpaceOriginal = $disk.FreeSpace
	$freeSpaceOrigMB = "{0:N2}" -f ($freeSpaceOriginal / 1MB) + " MB"
	$computerName =$env:COMPUTERNAME
	
	#OBJECT TO STORE REPORT INFO
	$hdCleanUpResults = new-object 'system.collections.generic.dictionary[string,string]'	#dictionary object to store results
	#$notRun = "--NOT RUN:"
	#$completedWithErrors = "--CompletedWithErrors:"
	if ($freeSpaceOriginal -lt $thresholdInBytes)
	{
		$hdCleanUpResults.Add(".DISK: Original Freespace", $freeSpaceOrigMB)
		
		#==============================================================CLEAN SMS CACHE
		If($cleanSMSCache)
		{
			Try
			{	
				$smsCacheSize = (Get-ChildItem -r $targetDrive$smsCache | Measure-Object -property length -sum)
				$smsCacheSizeMB = "{0:N2}" -f ($smsCacheSize.sum / 1MB) + " MB"
				$hdCleanUpResults.Add("SMS Cache: Original Size", $smsCacheSizeMB)
				$UIResourceMgr = New-Object -ComObject UIResource.UIResourceMgr
				$Cache = $UIResourceMgr.GetCacheInfo()
				$CacheElements = $Cache.GetCacheElements()
				foreach ($Element in $CacheElements)
				{
					$Cache.DeleteCacheElement($Element.CacheElementID)
				}
				$members = $cacheElements | get-Member -ErrorAction SilentlyContinue
				if(!($members))
				{
					Remove-Item $targetDrive$smsCache\* -recurse -force -ErrorAction Stop
				}
				
			}
			Catch
			{
				$completedWithErrors += "SMS Cache, "
			}
		
			$smsCacheSizeNew = (Get-ChildItem -r $targetDrive$smsCache | Measure-Object -property length -sum)
			$smsCacheSizeNewMB = "{0:N2}" -f ($smsCacheSizeNew.sum / 1MB) + " MB"
			$hdCleanUpResults.Add("SMS Cache: New Size", $smsCacheSizeNewMB)
	
		}
		Else
		{
			$notRun += "SMS Cache, "
		}

		#==============================================================CLEAN SYSTEM TEMP
		if($cleanTemp)	
		{
			CleanSystemFolder -t $targetDrive$systemTemp -obj $hdCleanUpResults -n "System Temp"
		}
		Else
		{
			$notRun += "System Temp, "
		}
		
		#==============================================================OFFLINE WEB PAGES
		if($cleanOfflineWeb)	
		{
			CleanSystemFolder -t $targetDrive$offlineWebPages -obj $hdCleanUpResults -n "Offline Web Pages"
		}
		Else
		{
			$notRun += "Offline Web Pages, "
		}
		
		#==============================================================DOWNLOADED PROGRAM FILES
		if($cleanDownloadProgFiles)	
		{
			CleanSystemFolder -t $targetDrive$downloadedProgramFiles -obj $hdCleanUpResults -n "Downloaded Program Files"
		}
		Else
		{
			$notRun += "Downloaded Program Files, "
		}
		
		#==============================================================CLEAN ENTIRE PROFILES (Using DELPROF2.EXE)
		if($cleanProfiles)
		{
			Try
			{
				$profSize = (Get-ChildItem -r $targetDrive$usersFolder | Measure-Object -property length -sum)
				$profSizeMB = "{0:N2}" -f ($profSize.sum / 1MB) + " MB"
				$hdCleanUpResults.Add("Profiles: Original Size", $profSizeMB)
				& "$delProf2Location\delprof2.exe" "/q"
			}
			Catch
			{
				$completedWithErrors += "Profiles(delprof2.exe), "
			}
		
			$profSizeNew = (Get-ChildItem -r $targetDrive$usersFolder | Measure-Object -property length -sum)
			$profSizeNewMB = "{0:N2}" -f ($profSize.sum / 1MB) + " MB"
			$hdCleanUpResults.Add("Profiles: New Size", $profSizeNewMB)
			
		}
		Else
		{
			$notRun += "Profiles(delprof2.exe), "
		}
		
		#==============================================================CLEAN IE CACHE (ALL USERS)
		if($ieCacheAllUsers)
		{
			CleanProfileFolder -u $targetDrive$usersFolder -t $ieCache -obj $hdCleanUpResults -n "IE Cache"
		}
		Else
		{
			$notRun += "IE Cache, "
		} 
		
		#==============================================================CLEAN GOOGLE CHROME CACHE (ALL USERS)
		if($ieCacheAllUsers)
		{
			CleanProfileFolder -u $targetDrive$usersFolder -t $googleChromeCache -obj $hdCleanUpResults -n "Google Chrome Cache"
		}
		Else
		{
			$notRun += "Google Chrome Cache, "
		} 
		
		
		#==============================================================CLEAN TEMP FOLDER (ALL USERS)
		if($ieCacheAllUsers)
		{
			CleanProfileFolder -u $targetDrive$usersFolder -t $userTmp -obj $hdCleanUpResults -n "User Temp"
		}
		Else
		{
			$notRun += "User Temp, "
		} 
		
		
		#==============================================================POST INFORMATION
		$hdCleanUpResults.Add("..NOT RUN",$notRun)
		$hdCleanUpResults.Add("..COMPLETED WITH ERRORS",$completedWithErrors)
		
		$disk = Get-WmiObject Win32_LogicalDisk -ComputerName . -Filter "DeviceID='$targetDrive'" | Select-Object Size,FreeSpace
		$freeSpaceNew = $disk.FreeSpace
		$freeSpaceNewMB = "{0:N2}" -f ($freeSpaceNew / 1MB) + " MB"
		$hdCleanUpResults.Add(".DISK: New Freespace", $freeSpaceNewMB)
		
		#==============================================================EMAIL INFORMATION
		#generate HTML Report
		$hdCleanUpResults.GetEnumerator() | Sort-Object -property Key | ConvertTo-HTML | Out-File "$($tempFolder.value)\$dateStamp-HDCleanUpReport.html"
		If(!($doNotSendEmail))
		{
			$emailSubject="HD Cleaning Report - $computerName"
			$emailBody = "Disk cleaning operations were performed on $computerName. The currenly low disk threshold is set to $lowDiskThreshold MB. See attached report for details."
			#send email to specified recipient and attach HTML report
			Send-MailMessage -To $emailRecipient -Subject $emailSubject -smtpServer $emailServer -From $emailSender -body $emailBody  -Attachments "$($tempFolder.value)\$dateStamp-HDCleanUpReport.html"
		}
	}
}

HardDriveCleaner -t 2000 -sms -tmp -ie -utmp -chrome -owp -dpf -noemail -to "bogus@limcollege.edu" -from "bogus@limcollege.edu" -smtp "bogus.limcollege.edu"
