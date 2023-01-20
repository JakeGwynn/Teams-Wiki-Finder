<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

<#
.SYNOPSIS
Checks all Teams for the precense of a Wiki page and exports metadata about those Wikis to a CSV file. 
This can be used to evaluate how many Wiki pages need to moved to OneNote.
.EXAMPLE
.\Get-AllTeamsWikis.ps1 -AppId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -CertThumbprint "161C9622BAEFE47C50EFB305FD6805A95D37579E" `
-TenantName "contoso.onmicrosoft.com" -CsvExportPath "C:\Temp\WikiFilesInTeams.csv"
#>

Param(
   [Parameter(Mandatory=$true)]
   [string]$AppId,

   [Parameter(Mandatory=$true)]
   [string]$CertThumbprint,

   [Parameter(Mandatory=$true)]
   [string]$TenantName,

   [Parameter(Mandatory=$false)]
   [string]$CsvExportPath = "C:\Temp\WikiFilesInTeams.csv"
)

function DisplayInBytes($num) 
{
    $suffix = "B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"
    $index = 0
    while ($num -gt 1kb) 
    {
        $num = $num / 1kb
        $index++
    } 

    "{0:N1} {1}" -f $num, $suffix[$index]
}

function InstallModules ($modules) {
    if ((Get-PSRepository).InstallationPolicy -eq "Untrusted") {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        $psTrustDisabled = $true
    }

    foreach ($module in $modules) {
        $instModule = Get-InstalledModule -Name $module -ErrorAction:SilentlyContinue
        if (!$instModule) {
            if ($module -eq "PnP.PowerShell") {
                $spModule = Get-InstalledModule -Name "SharePointPnPPowerShellOnline" -ErrorAction:SilentlyContinue
                if ($spModule) {
                    throw('Please remove the older "SharePointPnPPowerShellOnline" module before the deployment can install the new cross-platform module "PnP.PowerShell"')                    
                }
                else {
                    Install-Module -Name $module -Scope CurrentUser -AllowClobber -Confirm:$false -MaximumVersion 1.9.0
                }
            }
            else {
                try {
                    Write-Host('Installing required PowerShell Module {0}' -f $module) -ForegroundColor Yellow
                    Install-Module -Name $module -Scope CurrentUser -AllowClobber -Confirm:$false
                }
                catch {
                    throw('Failed to install PowerShell module {0}: {1}' -f $module, $_.Exception.Message)
                } 
            }

        }
           
    }
    
    if ($psTrustDisabled) {
        Set-PSRepository -Name PSGallery -InstallationPolicy Untrusted
    }
}

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

$preReqModules =  "PnP.PowerShell", "ExchangeOnlineManagement"

# Install required PS Modules
Write-Host "Installing required PowerShell Modules..." -ForegroundColor Yellow
InstallModules -Modules $preReqModules
foreach ($module in $preReqModules) {
    $instModule = Get-InstalledModule -Name $module -ErrorAction:SilentlyContinue
    Import-Module $module
    if (!$instModule) {
        throw('Failed to install module {0}' -f $module)
    }
}
Write-Host "Installed modules" -ForegroundColor Green

Write-Host "Authenticate to Exchange Online with an account that has an Exchange Online Administrator role `r`n" -ForegroundColor Yellow
Connect-ExchangeOnline -ShowBanner:$False

Write-Host "Retrieving M365 Groups attached to a MS Team `r`n" -ForegroundColor Yellow
$AllGroups = Get-UnifiedGroup -Filter {ResourceProvisioningOptions -eq "Team"} -ResultSize Unlimited
[System.Collections.Generic.List[object]]$AllWikiFiles = @()

Write-Host "Getting Wiki metadata for all MS Teams `r`n" -ForegroundColor Yellow
foreach ($Group in $AllGroups) {
    $WikiModifiedSinceCreation = $false
    $DocLibItems = $null
    Write-Host "Team Name: $($Group.DisplayName)"
    try {
        if ($null -ne $Group.SharePointSiteUrl) {
            Connect-PnPOnline -ClientId $AppId -Thumbprint $CertThumbprint -Url $Group.SharePointSiteUrl -Tenant $TenantName -WarningAction Silently
            $WikiDocLib = Get-PnPList -Identity "Teams Wiki Data"
            if ($WikiDocLib) {
                $DocLibItems = (Get-PnPListItem -List "Teams Wiki Data"  -PageSize 5000).FieldValues
                foreach ($File in $DocLibItems) {
                    if ($File.File_x0020_Type -eq "mht"){
                        if ($File.Created -le ($File.Modified).AddSeconds(-30)) {
                            $WikiModifiedSinceCreation = $true
                        }
                        $AllWikiFiles.Add([pscustomobject]@{
                            WikiFileName = $File.FileLeafRef
                            SharePointSiteUrl = $Group.SharePointSiteUrl
                            FileSize = (DisplayInBytes $File.SMTotalFileStreamSize)
                            CreatedDate = $File.Created
                            ModifiedDate = $File.Modified
                            Channel = ($File.FileDirRef -split "Teams Wiki Data")[1].Trim("/")
                            ModifiedSinceCreation = $WikiModifiedSinceCreation
                        })
                    }
                }
            }
            Write-Host "`r`n"
        }
    }
    catch {
        Write-Host "Error getting metadata for this group `r`n" -ForegroundColor Red
    }
}
$AllWikiFiles | Export-Csv -Path (New-Item -Path $CsvExportPath -Force) -NoTypeInformation

Disconnect-ExchangeOnline -Confirm:$false
