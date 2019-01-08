
<#
.SYNOPSIS
Get-O365GroupCheck.ps1 examines each O365 Group (Unified Group) in a given tenant and determines if properties in the group meet Lilly requirements.

.DESCRIPTION
Get-O365GroupCheck.ps1 examine each O365 (unified) group in a given tenant and determines if the group meets the following criteria.
No users in any of the extended properties on O365 groups

If any or all of the checks listed above are not met, the group along with additional information, is written to a SharePoint Online list (location of the List is noted in the json config file).
Once the list is complete, a REST call is made to start a FLOW designed by John Carter to handle notifications of issues to the Group Owners.

The script reads all parameters from a JSON file that is fed to the script via the pipeline.  Check the Example for a demo of this.

Date Written: 10/18/2018
CR/SR/INC: N/A
Author: Bob Dillon

.EXAMPLE
(Get-Content -Raw "E:\ScheduledTasks\CAAccountData.json") | ConvertFrom-Json | E:\ScheduledTasks\Get-O365GroupCheck.ps1

.LINK
	http://ix1tfsprod01.rf.lilly.com:8080/tfs/web/UI/Pages/Scc/Explorer.aspx?pguid=cf26c64d-6a89-42b6-a8f9-785e6cd6e57d#path=%24%2FLillyNet%2FScripts%2FPowerShell
#>

[CmdletBinding()]
param (
	[parameter(ValueFromPipelineByPropertyName=$true)][string]$CAInfoFile,
    [parameter(ValueFromPipelineByPropertyName=$true)][string]$AppIdFile,
	[parameter(ValueFromPipelineByPropertyName=$true)][string]$listSite,
	[parameter(ValueFromPipelineByPropertyName=$true)][string]$listName,
	[parameter(ValueFromPipelineByPropertyName=$true)][string]$ListURLForQuickLaunch,
	[parameter(ValueFromPipelineByPropertyName=$true)][string]$TenantID
)

function Invoke-Disposal {
	Get-Variable -exclude Runspace | Where-Object {$_.Value -is [System.IDisposable]} | 
		Foreach-Object {$_.Value.Dispose()}
		}

# Load required Modules
If(!(Get-Module -Name "o365.Utility")){Import-Module -Name "o365.Utility"}
If(!(Get-Module -Name "SharePointPnPPowerShellOnline")){Import-Module "SharePointPnPPowerShellOnline"}

# Set up webproxy access
$Wcl = new-object System.Net.WebClient
$Wcl.Headers.Add("user-agent", "PowerShell Script")
$Wcl.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

#Create Some Creds for access
$GraphCred = (ConvertFrom-Json -InputObject (Get-Content $CAInfoFile -ReadCount 0 | Out-String) | ForEach-Object{New-Object -TypeName System.Management.Automation.PSCredential ($_.UserName, (ConvertTo-SecureString $_.PasswdEncrStr))})
$SPOCred = (ConvertFrom-Json -InputObject (Get-Content $AppIdFile -ReadCount 0 | Out-String) | ForEach-Object{New-Object -TypeName System.Management.Automation.PSCredential ($_.UserName, (ConvertTo-SecureString $_.PasswdEncrStr))})

# And now connect to Graph and SharePoint Online
Connect-MSGraphAPI -Credential $GraphCred -TenantId $TenantID | out-null
Connect-PnPOnline -Url $listSite -AppID $SPOCred.UserName -AppSecret ($SPOCred.GetNetworkCredential().Password)

# Check the see if the List exists.  If $true, then capture all the list items and delete the list
If(Get-PnPList -Identity $listName){
	$lastRunList = Get-PnPListItem -List $listName
	Remove-PnPList -Identity $listName -Force
	Start-Sleep 10
}

# Create the List using a template from the Template Gallery on the site collection
$ListTemplateInternalName = "$listName.stp"
$Context = Get-PnPContext
$Web = $Context.Site.RootWeb
$ListTemplates = $Context.Site.GetCustomListTemplates($Web)
$Context.Load($Web)
$Context.Load($ListTemplates)
Invoke-PnPQuery
$ListTemplate = $ListTemplates | Where-Object{$_.InternalName -eq $ListTemplateInternalName}
$ListCreation = New-Object Microsoft.SharePoint.Client.ListCreationInformation -Property @{Title=$ListName;ListTemplate=$ListTemplate;QuickLaunchOption="On"}
$Web.Lists.Add($ListCreation)
Invoke-PnPQuery | out-null

# Find all groups that do not meet Lilly requirements
$groupsThatDoNotMeetCriteria = Get-o365UnifiedGroup -Property DisplayName,createdDateTime,extbgcqkws7_objectOwner,Classification | Where-Object{$_.extbgcqkws7_objectOwner.primaryOwner, $_.extbgcqkws7_objectOwner.primaryOwnerId, $_.extbgcqkws7_objectOwner.secondaryOwner, $_.extbgcqkws7_objectOwner.secondaryOwnerId, $_.Classification -contains $null}
$groupsThatDoNotMeetCriteria = {$groupsThatDoNotMeetCriteria}.Invoke()

ForEach($group in $groupsThatDoNotMeetCriteria){
	Connect-MSGraphAPI -Credential $GraphCred -TenantId $TenantID | out-null
	$isinLastRun = $null
	$isinLastRun = $lastRunList | Where-Object{$_.FieldValues.Title -eq $group.id}
	$newcount,$deleteDate = If($isinLastRun){
		($isinLastRun)['DaysNonCompliant'] + 1
		($isinLastRun)['GroupDeletionDate']
	}
	Else{
		1,(get-Date).AddDays(30)
	}
	$additionalProperty = $null
	$additionalProperty = Get-o365UnifiedGroup -Id $group.id | Select-Object -expand Owners | Select-Object -Last 1 | Select-Object -expand id
	Add-PnPListItem -List $listName -Values @{"Title" = "$($group.id)";"O365GroupDisplayName" = "$($group.displayName)";"CreatedBy" = "$additionalProperty";"PrimaryOwner" = "$($group.extbgcqkws7_objectOwner.primaryOwner)";"PrimaryOwnerID" = "$($group.extbgcqkws7_objectOwner.primaryOwnerId)";"SecondaryOwner" = "$($group.extbgcqkws7_objectOwner.secondaryOwner)";"SecondaryOwnerID" = "$($group.extbgcqkws7_objectOwner.secondaryOwnerId)";"GroupCreationDate" = "$($group.createdDateTime)";"Classification" = "$($group.classification)";"DaysNonCompliant" = "$newCount";"GroupDeletionDate" = "$deleteDate"} | out-null
	}

#Finish Up, close the connections
Write-Verbose "Adding Navigation to QuickLaunch"
Add-PnPNavigationNode -Title "GroupGovernance" -Url $ListURLForQuickLaunch -Location QuickLaunch | out-null

Disconnect-MSGraphAPI
Disconnect-PnPOnline
