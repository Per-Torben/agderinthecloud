<#
.SYNOPSIS
    This script copies files and folders from one OneDrive to another within the same tenant.
.DESCRIPTION
    This script copies files and folders from one OneDrive to another within the same tenant. 
    It uses the PnP PowerShell module to connect to the source and target OneDrive sites. 
    It then collects the content from the source and target sites and copies the files and folders from the source to the target. 
    The script also creates the folder structure in the target site before copying the files. 
    The script also handles errors and logs them to a CSV file.
.EXAMPLE
    .\Copy-OneDriveFiles.ps1 -SourceUserUnderscore "adelev_m365x32391902_onmicrosoft_com" -DestinationUserUnderscore "miriamg_m365x32391902_onmicrosoft_com"
.NOTES
    File Name      : Copy-OneDriveFiles.ps1
    Author         : Per-Torben Sørensen - agderinthe.cloud
    Tested on      : PowerShell 7.4.2 and pnp.powershell 2.4.0 
#>
<#
.DISCLAIMER
    This script is provided AS IS without warranty of any kind. The author further disclaims all implied warranties including, without limitation, any implied warranties of merchantability or of fitness for a particular purpose. 
    The entire risk arising out of the use or performance of this script remains with you. 
    In no event shall the author be held liable for any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss of business information, 
    or other pecuniary loss) arising out of the use of or inability to use this script, even if the author has been advised of the possibility of such damages. 
    The use of this script carries no support from the author, unless otherwise specified. By using this script, you agree to these terms.
#>
param(
[Parameter(Mandatory=$true)]
[string]$SourceUserUnderscore,

[Parameter(Mandatory=$true)]
[string]$DestinationUserUnderscore
)

#region Set these parameters to the correct values

    # Connection parameters to connect to the source and target OneDrive sites using certificate authentication
    $connectionsParams = @{
        ClientId = 'db69ee7c-a859-4000-8d99-111111111111'
        Thumbprint = 'DB1E8A01D12628F6FB2704D6BC12111111111111'
        Tenant = '31e7505b-5cfa-47e8-a66f-111111111111'
    }

    # The Tenant inital domain
    $InitialDomain = "M365x32391902"
    
    # The name of the folder where the files will be copied in the target OneDrive
    $targetfoldername = "Files from another OneDrive"
    
    # The path to save the error files
    $csvpath = "C:\scripts"

#endregion

#region Do not change anything in this region
    # Build the URLs for the source and target OneDrive sites
    $TenantURL = "$($InitialDomain)-my.sharepoint.com"
    $urlsource = "https://$($TenantURL)/personal/$($sourceuserunderscore)/"
    $urltarget = "https://$($TenantURL)/personal/$($destinationUserUnderscore)/"
    $destinationOneDriveSiteRelativePath = "Documents/$targetfoldername"
    $destinationOneDrivePath = "/personal/$destinationUserUnderscore/Documents/$targetfoldername"
    $departingOneDrivePath = "/personal/$($sourceuserunderscore)/Documents"
#endregion

#region Collect content
    # Verify errorlog folder is valid
    try {
        Get-ChildItem $csvpath -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "Caught an error: $_" -ForegroundColor Red
        # Stop the script
        break
    }
    # Collecting content from source and target
    Connect-PnPOnline -Url $urlsource @connectionsParams
    $Content = Get-PnPListItem -List Documents -PageSize 1000

    Connect-PnPOnline -Url $urltarget @connectionsParams
    $Content2 = Get-PnPListItem -List Documents -PageSize 1000
#endregion

#region folders
    # Etablish folder structure in target
    $folders = $Content| Where-Object {$_.FileSystemObjectType -contains "Folder"}
    Write-Host "`nCreating Folder Structure in Target" -ForegroundColor Blue

    $foldererrors = @()
    $count = $folders.Count
    $i=0
    foreach ($folder in $folders) {
        # Write progress bar on screen
        $i++
        $percentprogress = [math]::Round((($i/$count) *100),2)
        Write-Progress -Activity "Creating $count folders" -Status "Folder $($i) of $($count), $($percentprogress)%" -PercentComplete $percentprogress
            
        $path = ('{0}{1}' -f $destinationOneDriveSiteRelativePath, $folder.fieldvalues.FileRef).Replace($departingOneDrivePath, '')
        Try {
            $newfolder = Resolve-PnPFolder -SiteRelativePath $path -ErrorAction Stop
        }
        catch {
            $fileerrors += [PSCustomObject]@{
                ErrorMessage = $_.Exception.Message
                Folder = $path
            }
        }
    }

    if ($foldererrors.count -gt 0) {
        $now = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"
        $foldererrors | Export-Csv -Path "$($csvpath)\$($now)_foldererrors.csv" -NoTypeInformation
        Write-Host "Folder errors detected, see $($csvpath)\$($now)_foldererrors.csv" -ForegroundColor Red
    }
    else {
        Write-Host "No folder errors detected" -ForegroundColor Green
    }
#endregion

#region files
# Copy files from source to target
$files = $Content | Where-Object {$_.FileSystemObjectType -contains "File"}
Write-Host "`nCopying Files" -ForegroundColor Blue

$fileerrors = @()
$count = $files.Count
$i=0
foreach ($file in $files) {
    # Write progress bar on screen
    $i++
    $percentprogress = [math]::Round((($i/$count) *100),2)
    Write-Progress -Activity "Copying $count files" -Status "File $($i) of $($count), $($percentprogress)%" -PercentComplete $percentprogress
      
    $destpath = ("$destinationOneDrivePath$($file.fieldvalues.FileDirRef)").Replace($departingOneDrivePath, "")
    if ($Content2.fieldvalues.FileLeafRef -notcontains $file.fieldvalues.FileLeafRef) {
        try {
            $newfile = Copy-PnPFile -SourceUrl $file.fieldvalues.FileRef -TargetUrl $destpath -Force -ErrorVariable errors -ErrorAction Stop
        }
        catch {
            $fileerrors += [PSCustomObject]@{
                ErrorMessage = $_.Exception.Message
                SourceFile = $file.fieldvalues.FileRef
                TargetFile = $destpath
            }
        }
    }
}

if ($fileerrors.count -gt 0) {
    $now = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"
    $fileerrors | Export-Csv -Path "$($csvpath)\$($now)_fileerrors.csv" -NoTypeInformation
    Write-Host "File errors detected, see $($csvpath)\$($now)_fileerrors.csv" -ForegroundColor Red
}
else {
    Write-Host "No file errors detected" -ForegroundColor Green
}
#endregion
