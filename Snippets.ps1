# PowerShell snippets for 2021 M365 Collab conference

# Create certificate for login
# stolen from https://www.powershellgallery.com/packages/ExchangeOnlineManagement/1.0.1/Content/Create-SelfSignedCertificate.ps1 and other places
# Google "Create-SelfSignedCertificate.ps1" for more places

.\Create-SelfSignedCertificate.ps1 -CommonName "M365Collab" -StartDate 2021-11-07 -EndDate 2022-01-20 -Password (ConvertTo-SecureString "pass@word1" -Force -AsPlainText)

# Run in Windows PowerShell 5.x in an Admin console
Install-Module PnP.PowerShell -Force –Scope AllUsers
Install-Module Microsoft.Online.SharePoint.PowerShell -Force –Scope AllUsers
Install-Module Microsoft.PowerApps.Administration.PowerShell
Install-Module Microsoft.PowerApps.PowerShell –AllowClobber

# Only gets the modules installed in the current PowerShell host
Get-InstalledModule

# Gets all of the modules installed on the machine
Get-Module -ListAvailable

# Make sure a module is installed where both PowerShells and all users can use it
Get-InstalledModule -Name Microsoft.Online.SharePoint.PowerShell | Select-Object InstalledLocation

Update-Module Microsoft.Online.SharePoint.PowerShell
  
# Use stored credentials
Add-PnPStoredCredential -Name "https://tenant.sharepoint.com" -Username yourname@tenant.onmicrosoft.com
Connect-PnPOnline -Url "https://tenant.sharepoint.com”
Connect-PnPOnline -Url "https://tenant.sharepoint.com/sites/hr”

$Credentials = Get-PnPStoredCredential -Name "https://tenant.sharepoint.com" 
Connect-SPOService -Url https://tenant-admin.sharepoint.com -Credentials $ Credentials

# Upload a file
$web = https://tenant.sharepoint.com/sites/hr
$folder = "Shared Documents"
Connect-PnPOnline -Url $web
Add-PnPFile -Path '.\Boot fairs with Graphic design.docx' -Folder $folder

# Add a folder
Add-PnPFolder -Name "Folder 1" -Folder $folder
Add-PnPFile -Path '.\Building materials licences to budget for Storytelling.docx'  -Folder "$folder\Folder 1“

# Get Internal Shared Files example
# Doesn't actually work
$doclibs = Get-PnPList -Includes DefaultViewUrl,IsSystemList | Where-Object -Property IsSystemList -EQ -Value $false | Where-Object -Property BaseType -EQ -Value "DocumentLibrary"
    Foreach ($doclib in $doclibs) 
        {
        $docs = Get-PnPListItem -List $DocLib
        foreach ($doc in $docs) {
            if (($doc.FieldValues).SharedWithUsers -ne $null) {
                foreach ($user in (($doc.FieldValues).SharedWithUsers))  {
                    Write-Output "$(($doc.FieldValues).FileRef) - $($user.email)"
                    }
                }
             }
         } 

# Get Extended File Info example
# Doesn't work either
$doclibs = Get-PnPList -Includes DefaultViewUrl,IsSystemList | Where-Object -Property IsSystemList -EQ -Value $false | Where-Object -Property BaseType -EQ -Value "DocumentLibrary“

    Foreach ($doclib in $doclibs) 
        {
	 $doclibTitle = $doclib.Title
        $docs = Get-PnPListItem -List $DocLib
	 $docs | ForEach-Object {  Get-PnPProperty -ClientObject $_ -Property File, ContentType, ComplianceInfo}
        foreach ($doc in $docs) {
		 [pscustomobject]@{Library= $doclibTitle;Filename = ($doc.File).Name;ContentType = ($doc.ContentType).Name;Label = ($doc.ComplianceInfo).ComplianceTag}  

} 

# Bulk Undelete Files example
# Actually does work
Connect-PnPOnline -Url https://sadtenant.sharepoint.com/ -Credentials SadTenantAdmin
$bin = Get-PnPRecycleBinItem | Where-Object -Property Leafname -Like -Value "*.jpg"  | Where-Object -Property Dirname -Like -Value “Important Photos/Shared Documents/*"  | Where-Object -Property DeletedByEmail -EQ -Value baduser@sadtenant.phooey
$bin.count
$bin | ForEach-Object  -begin { $a = 0} -Process  {Write-Host "$a - $($_.LeafName)" ; $_ | Restore-PnPRecycleBinItem -Force ; $a++ } -End { Get-Date }
($bin[20001..30000]) | ForEach-Object  -begin { $a = 0} -Process  {Write-Host "$a - $($_.LeafName)" ; $_ | Restore-PnPRecycleBinItem -Force ; $a++ } -End { Get-Date }

# https://www.toddklindt.com/PoshRestoreSPOFiles

# Create sites examples
New-SPOSite 
<# No Group, No Team, No Bueno
Can be Groupified later #>
New-PnPSite -Type TeamSite -Title “Modern Team Site" -Alias ModernTeamSite –IsPublic
<# Group, No Team
Can be Teamified later#>

New-Team -DisplayName “Fancy Group" -Description “Fancy Group made by PowerShell?" -Alias FancyGroup -AccessType Public
# There is no later!

# Group Membership example
# Set some values 
# use Get-PnPUnifiedGroup to get Unified Group names 
# Name of Unified Group whose owners and membership we want to copy 
$source = "Regulations"
# Name of Unified Group whose owners and membership we want to populate 
$destination = "Empty"
# Whether to overwrite Destination membership or merge them 
$mergeusers = $false
# Check to see if PnP Module is loaded 
$pnploaded = Get-Module PnP.PowerShell
if ($pnploaded -eq $false) { 
    	Write-Host "Please load the PnP PowerShell and run again" 
    	Write-Host "install-module PnP.PowerShell" 
    	break 
    } 

# PnP Module is loaded
# Check to see if user is connected to Microsoft Graph 
try 
{ 
    $owners = Get-PnPMicrosoft365GroupOwners -Identity $source 
} 
catch [System.InvalidOperationException] 
{ 
    Write-Host "No connection to Microsoft Graph found"  -BackgroundColor Black -ForegroundColor Red 
    Write-Host "No Azure AD connection, please connect first with Connect-PnPOnline -Graph" -BackgroundColor Black -ForegroundColor Red 
break 
} 
catch [System.ArgumentNullException] 
{ 
        Write-Host "Group not found"  -BackgroundColor Black -ForegroundColor Red 
        Write-Host "Verify connection to Azure AD with Connect-PnPOnline -Graph" -BackgroundColor Black -ForegroundColor Red 
        Write-Host "Use Get-PnPUnifiedGroup to get Unified Group names"  -BackgroundColor Black -ForegroundColor Red 
        break 
} 
catch 
{ 
    Write-Host "Some other error"   -BackgroundColor Black -ForegroundColor Red 
break 
}

$members = Get-PnPMicrosoft365GroupMembers -Identity $source
if ($mergeusers -eq $true) { 
     # Get existing owners and members of Destination so that we can combine them 
    $ownersDest = Get-PnPMicrosoft365GroupOwners -Identity $destination 
    $membersDest = Get-PnPMicrosoft365GroupMembers -Identity $destination
    # Add the two lists together so we don't overwrite any existing owners or members in Destination 
    $owners = $owners + $ownersDest 
    $members = $members + $membersDest 
    }
# Set the owners and members of Destination 
$owners | ForEach-Object -begin  {$ownerlist  = @() } -process {$ownerlist += $($_.UserPrincipalName) } 
$members | ForEach-Object -begin  {$memberlist  = @() } -process {$memberlist += $($_.UserPrincipalName) }
Set-PnPMicrosoftGroup -Identity $destination -Members $memberlist -Owners $ownerlist
# https://www.toddklindt.com/PoshCopyO365GroupMembers

# Get all the Flow
Add-PowerAppsAccount

Get-AdminFlow | ForEach-Object { $ownername = (Get-MsolUser -ObjectId $_.CreatedBy.userId).DisplayName ; $owneremail = (Get-MsolUser -ObjectId $_.CreatedBy.userId).UserPrincipalName ; Write-Host $_.DisplayName, $ownername, $owneremail }



