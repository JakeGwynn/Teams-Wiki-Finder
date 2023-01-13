# Teams-Wiki-Finder
This solution checks all Teams for the precense of a Wiki page and exports metadata about those Wikis to a CSV file. 
This can be used to evaluate how many Wiki pages need to moved to OneNote.

## Architecture
The PowerShell script uses the PnP.PowerShell and ExchangeOnlineManagement modules to connect to Exchange Online and SharePoint Online. An Exchange Online Administrator's account will be required to authenticate when the script runs, and an App Registration will be used to authenticate to SharePoint Online with PnP.PowerShell. 

## Prerequisites
This solution requires an Azure AD App Registration with Sites.Read.All permissions. Before the Teams Wiki Finder script can be run, an App Registration must be created with "Sites.Read.All" permissions. The below command (requires PnP.PowerShell module) will create an app registration in Azure AD, add Sites.Read.All API permissions, consent to the API permission, generate a certificate to be used for authentication in the script, and finally uploads the certificate to the App Registration. 

**Please run the following command with a Global Administrator account:**
  ```powershell
Register-PnPAzureADApp -ApplicationName "Teams Wiki Finder" -Interactive `
-Tenant <contoso>.onmicrosoft.com -Store CurrentUser -Username "GlobalAdmin@contoso.com" `
-Password (Read-Host -AsSecureString -Prompt "Enter Password") -SharePointApplicationPermissions "Sites.Read.All"
  ```
**Note the "AzureAppId/ClientId" and "Certificate Thumbprint". These values will be used as parameters when running the Get-AllTeamsWiki.ps1 script**

## Running the Script
1. Download the **Get-AllTeamsWikis.ps1** script to the machine where Register-PnPAzureADApp was run.
2. In a PowerShell windows, navigate to the folder where the Get-AllTeamsWikis.ps1 script is located.
3. Run the following command within the PowerShell windows, which will run the script. Replace the AppId and CertThumbprint parameters with the values output from the Register-PnPAzureADApp command. Replace the TenantName parameter with your tenant's xxx.onmicrosoft.com domain name. Finally, replace the CsvExportPath parameter with the location and file name of the CSV that will contain all of the Wiki metadata. If the folder does not already exist, it will be created.

  ```powershell
.\Get-AllTeamsWikis.ps1 -AppId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -CertThumbprint "161C9622BAEFE47C50EFB305FD6805A95D37579E" `
-TenantName "contoso.onmicrosoft.com" -CsvExportPath "C:\Temp\WikiFilesInTeams.csv"
  ```
