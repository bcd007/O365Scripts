<#
.SYNOPSIS
This script will retrieve the list of current O365 Service Alerts from GraphAPI, and update a SPO List with the alerts flagged as 'serviceDegradation'

.DESCRIPTION
This script will retrieve the list of current O365 Service Alerts with a status of 'serviceDegradation' from GraphAPI, and update a SPO/Teams List with the alert data
This script will also remove any items in the SPO list where the Status of the Alert has changed from 'serviceDegradation' to any other status
Note:   This script uses only GraphAPI, no other modules are required.
        This script requires a json file with the AppID Key and encrypted secret.  The script can be modified to support a certificate or Azure keystore

The SPO list is a standard SharePoint Online list that has the following columns:
title,alertID,impactDescription,classification,featureGroup,service,status,startDateTime

This script requires an Azure AD ApplicationID that has the following GraphAPI Application Permissions added and consented:
     ServiceMessage.Read.All
     ServiceHealth.Read.All
     Sites.Manage.All
     or
     Sites.Selected (If you can get this working)

Date Written: 01/12/2022
CR/SR/INC: N/A
Author: Bob Dillon
Version:  1.0

.EXAMPLE
(Get-Content -Raw "C:\Temp\O365IssuesScriptData.json") | ConvertFrom-Json | C:\Temp\Get-O365AlertsUsingGraph.ps1

.EXAMPLE
Get-O365AlertsUsingGraph.ps1 -tenantID (tenantID) -clientFile (Path and/or Filename of CredFile.json) -siteID (GUID of SharePoint Site> -listId <GUID of SPO List to write to)
#>

[CmdletBinding()]
param (
	[parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true)][GUID]$tenantID,
	[parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true)][string]$clientFile,
	[parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true)][GUID]$siteID,
	[parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true)][GUID]$listID
)

# Create the credential object from the encrypted AppID file ($clientFile)
$applicationReg = (ConvertFrom-Json -InputObject (Get-Content $clientFile -ReadCount 0 | Out-String)) | ForEach-Object{New-Object -TypeName System.Management.Automation.PSCredential ($_.UserName, (ConvertTo-SecureString $_.PasswdEncrStr))}

# Create the GraphAPI connection and token to use throughout the script
$ReqTokenBody = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    client_Id     = $applicationReg.UserName
    Client_Secret = $applicationReg.GetNetworkCredential().password
} 
$TokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token" -Method POST -Body $ReqTokenBody
$Header = @{
    Authorization = "$($TokenResponse.token_type) $($TokenResponse.access_token)"
    ConsistencyLevel = "eventual"
}

# Retrieves the current list of service impacts to O365 and creates an object with just the .value property of the Graph REST call
# Note:  This retrieves the last 100 incidents per GraphAPI paging rules
$serviceHealthresults = @()
$serviceStatusReturnURI = "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/issues"
$SvcHealthAlerts = Invoke-RestMethod -Headers $Header -Uri $serviceStatusReturnURI -Method Get -ContentType "application/json"
$serviceHealthresults = $SvcHealthAlerts.value

# Get another 900 because Micrsoft hasn't figured out how to sort their alerts by date....yet.
$count = 0
Do{$count++ ; $serviceHealthresults += (Invoke-RestMethod -Uri $SvcHealthAlerts.'@odata.nextLink' -Headers $Header -Method Get -ContentType "application/json").Value}
Until($count -eq 9)

# Microsoft puts duplicate alerts in this report, lets get rid of them because 'dumb'
$serviceHealthresults = $serviceHealthresults | Sort-Object -Property id -Unique

# Get the List and List Items via Graph, and set up the Add List Item URI
$listItemsURI = "https://graph.microsoft.com/v1.0/sites/$siteID/lists/$listID/items?expand=fields"
$listItemsToExamine = Invoke-RestMethod -Headers $Header -Uri $listItemsURI -Method Get -ContentType "application/json"
$currentListItems = $listItemsToExamine.value
$listItemstoAddURI = "https://graph.microsoft.com/v1.0/sites/$siteID/lists/$listID/items"

# Parse through each service alert and determine if the individual alert needs to be added to or removed from the List.
ForEach($svcAlertToAnalyze in $serviceHealthresults){
    If($currentListItems.fields.AlertID -notcontains $svcAlertToAnalyze.Id -And $svcAlertToAnalyze.Status -eq "serviceDegradation"){
        # Create a nested hashtable to match the JSON format/case needed to add a list item via GraphAPI
        $bodyToWrite = @{
            fields = 
                @{Title = $svcAlertToAnalyze.title
                AlertID = $svcAlertToAnalyze.id
                impactDescription = $svcAlertToAnalyze.impactDescription
                classification = $svcAlertToAnalyze.classification
                featureGroup = $svcAlertToAnalyze.featureGroup
                service = $svcAlertToAnalyze.service
                status = $svcAlertToAnalyze.status
                startDateTime = $svcAlertToAnalyze.startDateTime
            }
        }
        # Add any new alerts with an status of 'serviceDegradation' that are not in the list already
        $addnewlistItemResults = Invoke-RestMethod -Headers $Header -Uri $listItemstoAddURI -Method POST -ContentType "application/json" -Body ($bodyToWrite | ConvertTo-Json)
    }
    # Remove any list items where the status has changed from 'serviceDegradation' to any other status
    If($currentListItems.fields.AlertID -contains $svcAlertToAnalyze.Id -And $svcAlertToAnalyze.Status -eq "serviceRestored"){
        $currentlistitems | Where-Object{$_.fields.AlertID -eq $svcAlertToAnalyze.Id} | ForEach-Object{
            $deleteListItemURI = "https://graph.microsoft.com/v1.0/sites/$siteID/lists/$listID/items/$($_.id)"
            $listItemsTodelete = Invoke-RestMethod -Headers $Header -Uri $deleteListItemURI  -Method DELETE -ContentType "application/json"
        }
    }
}
