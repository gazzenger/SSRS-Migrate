<#
https://github.com/microsoft/ReportingServicesTools	
#>

Import-Module ReportingServicesTools

$cred = Get-Credential -Message "Please provide the report service account details for putting in all the SQL data sources. In the format DOMAIN\USERNAME"

$ServiceName = $cred.GetNetworkCredential().UserName
$ServicePwd = $cred.GetNetworkCredential().Password

if ( $ServiceName -eq '') {
	exit
}
if ( $ServicePwd -eq '') {
	exit
}

$ReportServerUrl = 'http://SERVER/Reports'
$ReportServerUri = 'http://SERVER/ReportServer'
$RootPath = 'C:\Temp\SearchAndReplace\Download'
$ScriptPath = 'C:\Temp\SearchAndReplace\'
$ReplacePath = 'C:\Temp\SearchAndReplace\Replace'
$ReplaceConnectFile = 'C:\Temp\SearchAndReplace\replacements-ConnectString.txt'
$SubsciptionPath = 'C:\Temp\SearchAndReplace\Subscriptions'

$ssrs = New-RsWebServiceProxy -ReportServerUri $ReportServerUri
$proxyNamespace = $ssrs.GetType().Namespace


# This function searches the replacements.txt file and replaces all occurences in a string before returning it
function Search-Replacements {
    param(
        [string] $ReplaceFile,
		[string] $InputString
    )
	#Write-Host $InputString
	foreach($line in Get-Content $ReplaceFile) {
		$LineStr = $line.Split(',')
		
		$InputString = $InputString -ireplace [regex]::Escape($LineStr[0]), $LineStr[1]
		#$InputString = $InputString.Replace($LineStr[0],$LineStr[1])
	}
	#Write-Host $InputString
	return $InputString
}

Echo ""
Echo "Setup Download and Subscriptions folder"
Echo ""
Set-Location -Path $ScriptPath
New-Item -ItemType Directory -Force -Path $RootPath
New-Item -ItemType Directory -Force -Path $SubsciptionPath

Echo ""
Echo "Emptying the Download and Subscription folders"
Echo ""
#Emtpy the Download and Subscriptions directories
Remove-Item -path ($RootPath + '\*') -Recurse
Remove-Item -path ($SubsciptionPath + '\*') -Recurse


Echo "Ensure the old SSRS database is configured in the Reporting Services Configuration Manager"
pause

Echo ""
Echo "Downloading all files from SSRS server"
Echo ""
#Download all files from SSRS
#The following command uses the a different api to access the reports, etc, however it misses downloading the report parts
#Out-RsFolderContent -ReportServerUri $ReportServerUri -RsFolder / -Destination $RootPath -Recurse
#The following command uses RestAPI to download all the reports, etc, IT does download the Report parts (with a modification to the C:\Program Files\WindowsPowerShell\Modules\ReportingServicesTools\0.0.5.1\Functions\Common\Get-FileExtension.ps1 file)
Out-RsRestFolderContent -ReportPortalUri $ReportServerUrl -RsFolder / -Destination $RootPath -RestApiVersion 'v1.0' -Recurse

Echo ""
Echo "Downloading all subscriptions from SSRS server"
Echo ""

#Copy ONLY directory structure to subscriptions folder
Get-ChildItem -Path $RootPath | ? {$_.PSIsContainer} | Copy-Item -Destination $SubsciptionPath -Recurse -Exclude '*.*'

#List all rdl files
$Files = Get-ChildItem -Recurse $RootPath -Include *.rdl | Where { ! $_.PSIsContainer }

#Loop through all these RDL files and download the subscription files
foreach ($File in $Files) {
	$RptName = $File.Name.Split('.')[0]
	$TempPath = Split-Path -Path $File.FullName
	$RptPath = $TempPath.Replace($RootPath,'').Replace('\','/')
	$NewPath = $TempPath.Replace($RootPath,$SubsciptionPath)
	Get-RsSubscription -ReportServerUri $ReportServerUri -Path ($RptPath + '/' + $RptName) | Export-RsSubscriptionXml ($NewPath + '/' + $RptName + '.xml' )
}

Echo "Done"
pause

#Rename Report Parts with RDL - to allow search and replacing
Get-ChildItem -Recurse $RootPath -Include 'Report Parts' | Where { $_.PSIsContainer } | foreach {Get-ChildItem $_.FullName | Rename-Item -NewName { $_.Name + ".rdl" } }

#Search and Replace all hard coded references in the files
Echo ""
Echo "The Search and Replace Script"
Echo "Provide there path as:"
Echo $RootPath
Echo ""
Echo "Select utf+8 encoding"
Echo ""
Echo "Then for file type select rdl files (in lowercase)"
Echo ""
& ($ReplacePath + '/script.exe')

#Rename Report Parts to RSC - ready for uploading
Get-ChildItem -Recurse $RootPath -Include 'Report Parts' | Where { $_.PSIsContainer } | foreach {Get-ChildItem $_.FullName | Rename-Item -NewName {$_.name -replace ".rdl",".rsc"} }

#pause and wait for user to change the db for SSRS
Echo "Please now configure SSRS to point to the new database in the Reporting Services Configuration Manager"
pause
Echo "Make Sure to return the service account back to a public ONLY account after the database is created."
pause
Echo "Are you sure?"
pause

#Add site roles
Echo ""
Echo "Adding AD Site Security Settings"
Echo ""
Grant-AccessToRs -ReportServerUri $ReportServerUri -UserOrGroupName "DOMAIN\ADMIN-USER" -RoleName "System Administrator"
Grant-AccessToRs -ReportServerUri $ReportServerUri -UserOrGroupName "DOMAIN\ADMIN-VIEWER" -RoleName "System User"

#Remove old reports
Echo ""
Echo "Emptying SSRS Folders"
Echo ""
Get-RsFolderContent -ReportServerUri $ReportServerUri -RsFolder / | Remove-RsCatalogItem -ReportServerUri $ReportServerUri -Confirm:$false

#Grant folder roles to parent level folder (this will filter through child folders)
Echo ""
Echo "Granting AD Folder Access"
Echo ""
Grant-RsCatalogItemRole -ReportServerUri $ReportServerUri -Identity 'DOMAIN\ADMIN-USER' -RoleName 'Browser' -Path '/'
Grant-RsCatalogItemRole -ReportServerUri $ReportServerUri -Identity 'DOMAIN\ADMIN-USER' -RoleName 'Content Manager' -Path '/'
Grant-RsCatalogItemRole -ReportServerUri $ReportServerUri -Identity 'DOMAIN\ADMIN-USER' -RoleName 'My Reports' -Path '/'
Grant-RsCatalogItemRole -ReportServerUri $ReportServerUri -Identity 'DOMAIN\ADMIN-USER' -RoleName 'Publisher' -Path '/'
Grant-RsCatalogItemRole -ReportServerUri $ReportServerUri -Identity 'DOMAIN\ADMIN-USER' -RoleName 'Report Builder' -Path '/'
Grant-RsCatalogItemRole -ReportServerUri $ReportServerUri -Identity 'DOMAIN\ADMIN-VIEWER' -RoleName 'Browser' -Path '/'

#Create new tree structure
Echo ""
Echo "Building Tree Structure in SSRS"
Echo ""
$Directories = Get-ChildItem -Path $RootPath -Recurse | ?{ $_.PSIsContainer } | Select-Object FullName | Sort FullName
#Loop through each directory
foreach ($Directory in $Directories) {
	$UploadPath = $Directory.FullName.Replace($RootPath,'').Replace('\','/')
	$UploadPathSplit = $UploadPath.Split('/')
	$Folder = $UploadPathSplit[$UploadPathSplit.Length-1]
	if ($UploadPathSplit.Length -eq 2) {
		$UploadPath = '/'
	}
	else {
		$UploadPath = $UploadPath.Replace(('/'+$Folder),'')
	}
	New-RsFolder -ReportServerUri $ReportServerUri -RsFolder $UploadPath -FolderName $Folder
}

#Upload all RSDS files
Echo ""
Echo "Uploading Shared Data Sources"
Echo ""
Write-RsFolderContent -ReportServerUri $ReportServerUri -Path ($RootPath + "\Data Sources") -Destination "/Data Sources" -Recurse

Echo ""
Echo "Updating Shared Data Source Connection Strings"
Echo ""
#List all RSDS files
$RSDSFiles = Get-ChildItem -Recurse $RootPath -Include *.rsds | Where { ! $_.PSIsContainer }
#Loop through all these RDL files
foreach ($Rsds in $RSDSFiles) {
	$RsdsName = $Rsds.Name.Split('.')[0]
	$TempPath = Split-Path -Path $Rsds.FullName
	$RsdsPath = $TempPath.Replace($RootPath,'').Replace('\','/')
	#Update the shared data sources
	$DataSource = Get-RsDataSource -ReportServerUri $ReportServerUri -Path ($RsdsPath + '/' + $RsdsName)
	echo ($RsdsPath + '/' + $RsdsName)
	$DataSource.ConnectString = Search-Replacements -ReplaceFile $ReplaceConnectFile -InputString $DataSource.ConnectString
	$DataSource.WindowsCredentials = $true
	$DataSource.Extension = "SQL"
	$DataSource.UserName = $ServiceName
	$DataSource.Password = $ServicePwd
	$DataSource.ImpersonateUserSpecified = $true
	$DataSource.ImpersonateUser = $false
	$DataSource.EnabledSpecified = $true
	$DataSource.Enabled = $true
	$DataSource.CredentialRetrieval = 'Store'
	#Update the data source
	Set-RsDataSource -ReportServerUri $ReportServerUri -RsItem ($RsdsPath + '/' + $RsdsName) -DataSourceDefinition $DataSource
}

pause

Echo ""
Echo "Uploading and Updating RDL Files"
Echo ""
#List all RDL files
$RDLFiles = Get-ChildItem -Recurse $RootPath -Include *.rdl, *.rsd | Where { ! $_.PSIsContainer }

#Loop through all these RDL files
foreach ($Rdl in $RDLFiles) {
	$RptName = $Rdl.Name.Split('.')[0]
	$TempPath = Split-Path -Path $Rdl.FullName
	$RptPath = $TempPath.Replace($RootPath,'').Replace('\','/')
	#Update the RDL file on the server
	Write-RsCatalogItem -ReportServerUri $ReportServerUri -Path $Rdl.FullName -RsFolder $RptPath -Overwrite
	#Update the embedded data sources
	$DataSources = Get-RsItemDataSource -ReportServerUri $ReportServerUri -RsItem ($RptPath + '/' + $RptName)
	foreach ($DataSource in $DataSources) {
		#if ( $DataSource.Item.Extension -eq 'SQL' )
		if ( $DataSource.Item.ConnectString )
		{
			echo ($RptPath + '/' + $RptName)
			$DataSource.Item.ConnectString = Search-Replacements -ReplaceFile $ReplaceConnectFile -InputString $DataSource.Item.ConnectString
			Write-Host $DataSource.Item.ConnectString
			$DataSource.Item.WindowsCredentials = $true
			$DataSource.Item.UserName = $ServiceName
			$DataSource.Item.Password = $ServicePwd
			$DataSource.Item.ImpersonateUserSpecified = $true
			$DataSource.Item.ImpersonateUser = $false
			$DataSource.Item.EnabledSpecified = $true
			$DataSource.Item.Enabled = $true
			$DataSource.Item.CredentialRetrieval = 'Store'
			Set-RsItemDataSource -ReportServerUri $ReportServerUri -RsItem ($RptPath + '/' + $RptName) -DataSource $DataSource
		}
		else {
			#If a data source IS a reference, but the reference DOES NOT exist
			if (!$DataSource.Item.Reference) {
				$string = Get-Content $Rdl
				$pattern = ('<DataSource Name="' + $DataSource.Name + '(.*?)</DataSource>')
				$source = [regex]::match($string, $pattern).Value
				if (!$source) { $source = $string }
				$pattern2 = "<DataSourceReference>(.*?)</DataSourceReference>"
				$sourceref = [regex]::match($source, $pattern2).Value
				if ($sourceref) {
					$sourceref = $sourceref.Replace('<DataSourceReference>','').Replace('</DataSourceReference>','')
					#check for shared data source existence on SSRS
					$error.clear()
					try { 
						Get-RsDataSource -ReportServerUri $ReportServerUri -Path ('/Data Sources/' + $sourceref)
					}
					catch { 
						#Shared Data Source DOES NOT Exist
						echo ("Shared Data Source: " + $sourceref + " does NOT exist")
						echo "Please rectify"
						echo ""
					}
					if (!$error) {
						#Shared Data Source DOES Exist
						$myDataSource = New-Object ("$proxyNamespace.DataSource") 
						$myDataSource[0].Name = $DataSource.Name
						$myDataSource[0].Item = New-Object ("$proxyNamespace.DataSourceReference")
						$myDataSource[0].Item.Reference = ('/Data Sources/' + $sourceref)
						$ssrs.SetItemDataSources(($RptPath + '/' + $RptName), @($myDataSource))
					}
				}
			}
		}
	}
	#Add the subscriptions
	$RptSubscribePath = $TempPath.Replace($RootPath,$SubsciptionPath)
	if(Test-Path ($RptSubscribePath + '/' + $RptName + '.xml'))
	{
		Import-RsSubscriptionXml ($RptSubscribePath + '/' + $RptName + '.xml') -ReportServerUri $ReportServerUri | Copy-RsSubscription -ReportServerUri $ReportServerUri -Path ($RptPath + '/' + $RptName)
		#check if the report in an archived folder, then disable these subscriptions
		if ($RptPath.ToUpper() -like '*_Archived*') {
			$subscriptions = $ssrs.ListSubscriptions(($RptPath + '/' + $RptName));
			ForEach ($subscription in $subscriptions)  
			{  
				$ssrs.DisableSubscription($subscription.SubscriptionID);  
			}  
		}
	}
}

pause

Echo ""
Echo "Please manually upload each of the Report Parts, these will now load in Internet Explorer"
Echo ""
pause

$ReportPartFolders = Get-ChildItem -Recurse $RootPath -Include 'Report Parts' | Where { $_.PSIsContainer }
foreach ($ReportPartFolder in $ReportPartFolders) {
	$ReportPartPath = $ReportPartFolder.FullName.Replace($RootPath,'').Replace('\','/')
	start  ($ReportServerUrl + '/browse' + $ReportPartPath )
	Echo ("Open a browser to " + ($ReportServerUrl + '/browse' + $ReportPartPath ))
	Echo ("And upload all the files from the folder " + $ReportPartFolder)
	pause
}

Echo ""
Echo "Done"
Echo ""
