# PowerShell Scripts To Help Manage O365

## Scripts

**Get-O365AlertsUsingGraph.ps1**:
Script to retrieve O365 service alerts and write the alert data to a specified SharePoint/Teams list using only GraphAPI

This script has the following requirements:
    A Teams/SPO List with the following columns and types:

---

  - title (Single line of text)
  - alertID (Single line of text)
  - impactDescription (Single line of text)
  - classification (Single line of text)
  - featureGroup (Single line of text)
  - service (Single line of text)
  - status (Single line of text)
  - startDateTime (Date/Time)

---

- An Azure Application ID that has FullControl on the Team/SPO site where the List to write to is located and ServiceHealth.Read.All as Application Permissions
- The GUID of the SPO/Team and List
- This iteration uses a JSON file that has the AppID and encrypted AppSecret using "Username" as the AppID property name
    *(Note:  You can always change this code to use a certificate or Credential Store)*

Get-Help .\Get-O365AlertsUsingGraph.ps1 -full
