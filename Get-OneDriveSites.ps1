<#
.SYNOPSIS
    Get all OneDrive sites in the tenant and list them in a table with the Name, Owner, and NameUnderscore properties.
.DESCRIPTION
    This script gets all OneDrive sites in the tenant and lists them in a table with the Name, Owner, and NameUnderscore properties.
    The script uses the PnP PowerShell module to connect to the SharePoint Admin site and get all OneDrive sites.
    The script then creates a custom object with the Name, Owner, and NameUnderscore properties and adds it to an array.
    The script outputs the array to the console.
    The script also clears the console and writes a message to the console.
.NOTES
    File Name      : Get-OneDriveSites.ps1
    Author         : Per-Torben Sørensen - agderinthe.cloud
    Tested on      : PowerShell 7.4.2 and pnp.powershell 2.4.0
#>
<#
.DISCLAIMER
    This script is provided AS IS without warranty of any kind. The author further disclaims all implied warranties including, 
    without limitation, any implied warranties of merchantability or of fitness for a particular purpose.

    The entire risk arising out of the use or performance of this script remains with you. 
    
    In no event shall the author be held liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, 
    or other pecuniary loss) arising out of the use of or inability to use this script, even if the author has been advised of the possibility of such damages.
    
    The use of this script carries no support from the author, unless otherwise specified. By using this script, you agree to these terms.
#>

#region Set these parameters to the correct values
$InitialDomain = "M365x32391902"
$connectionsParams = @{
    ClientId = 'db69ee7c-a859-4000-8d99-111111111111'
    Thumbprint = 'DB1E8A01D12628F6FB2704D6BC12111111111111'
    Tenant = '31e7505b-5cfa-47e8-a66f-1111111111111'
}
#endregion

# Connect to the SharePoint Admin site
$AdminSiteURL = "https://$($InitialDomain)-admin.sharepoint.com/"
Connect-PnPOnline -Url $adminSiteurl @connectionsParams

# Get All OneDrive Sites output to variable: $export
$AllOneDrives = Get-PnPTenantSite -IncludeOneDriveSites -Filter "Url -like '-my.sharepoint.com/personal/'"
$export = @()
$AllOneDrives | foreach {
    $output = [PSCustomObject]@{
        Name = $_.Title
        Owner = $_.Owner
        NameUnderscore = $_.Url.Split('/')[-1]
    }
    $export += $output
}
Clear-Host
Write-Host
Write-host -ForegroundColor Green 'Type $export to see the list'
Write-Host
