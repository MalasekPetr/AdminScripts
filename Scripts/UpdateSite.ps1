# Set environment variables
$path = "C:\Repos\AdminScripts"

# Prepare Log File Name
$dt = Get-Date -Format "yy-MM-dd-HH-mm"
$fname = "{0}\Logs\log-{1}.txt" -f $path, $dt 

# Start Logging
Start-Transcript -Path $fname -Append -NoClobber

# Greet Log Readers
Write-Information "Hi there ;)"

# Check Installed Module version
Get-Module -Name SharePointPnPPowerShellOnline -ListAvailable | Select-Object Name,Version | Out-Default

#########################################################################
################################ Comment ################################
#########################################################################

# Get Admin Credentials from params.json
$params = Get-Content $path\params.json -ErrorAction Stop | Out-String | ConvertFrom-Json
$orgName = $params.orgName
$password = $params.adminPSW | ConvertTo-SecureString -asPlainText -Force 
$userCredential = New-Object System.Management.Automation.PSCredential($params.adminUPN,$password)

# Connect to SharePoint Administration
$url = "https://{0}-admin.sharepoint.com" -f $orgName
Connect-PnPOnline -Url $url -Credential $userCredential -ErrorAction Stop

# Return Connection Information
Get-PnPConnection -ErrorAction Stop | Out-Default
Write-Information "Hi 01"

# Get Script configuration from .json
$config = Get-Content $path\Scripts\addsite.json -ErrorAction Stop | Out-String | ConvertFrom-Json

# Get SiteDesign configuration from .json files
$featuresjson = Get-Content $path\Designs\Scripts\Features.json -Raw
$features = Add-PnPSiteScript -Title "AVX-DMS Features" -Content $featuresjson

$sitecolumnsjson = Get-Content $path\Designs\Scripts\SiteColumns.json -Raw
$sitecolumns = Add-PnPSiteScript -Title "AVX-DMS Site Columns" -Content $sitecolumnsjson

$sitecolumnsxmljson = Get-Content $path\Designs\Scripts\SiteColumnsXML.json -Raw
$sitecolumnsxml = Add-PnPSiteScript -Title "AVX-DMS Site Columns (XML)" -Content $sitecolumnsxmljson

$contenttypesjson = Get-Content $path\Designs\Scripts\ContentTypes.json -Raw
$contenttypes = Add-PnPSiteScript -Title "AVX-DMS Content Types" -Content $contenttypesjson

$baselistsjson = Get-Content $path\Designs\Scripts\BaseLists.json -Raw
$baselists = Add-PnPSiteScript -Title "AVX-DMS Base Lists" -Content $baselistsjson

$baselistsupdatejson = Get-Content $path\Designs\Scripts\BaseListsUpdate.json -Raw
$baselistsupdate = Add-PnPSiteScript -Title "AVX-DMS Base Lists Update " -Content $baselistsupdatejson

$documentlibrariesjson = Get-Content $path\Designs\Scripts\DocumentLibraries.json -Raw
$documentlibraries = Add-PnPSiteScript -Title "AVX-DMS Document Libraries" -Content $documentlibrariesjson

$navigationjson = Get-Content $path\Designs\Scripts\Navigation.json -Raw
$navigation = Add-PnPSiteScript -Title "AVX-DMS Navigation" -Content $navigationjson

$logojson = Get-Content $path\Designs\Scripts\Logo.json -Raw
$logo = Add-PnPSiteScript -Title "AVX-DMS Logo" -Content $logojson

$themejson = Get-Content $path\Designs\Scripts\Theme.json -Raw
$theme = Add-PnPSiteScript -Title "AVX-DMS Theme" -Content $themejson

# Create new SiteDesign
$scriptname = "AVX-DMS Team Site ({0})" -f $dt
$sitedesign = Add-PnPSiteDesign -Title $scriptname `
 -WebTemplate "64" `
 -SiteScriptIds $features.Id, $baselists.Id, $sitecolumns.Id, $sitecolumnsxml.Id, $contenttypes.Id, $baselistsupdate.Id, $documentlibraries.Id, $navigation.Id, $logo.Id, $theme.Id `
 -Description "AVX-DMS Team Site"

# Grant Permissions for new SiteDesign
Grant-PnPSiteDesignRights -Identity $sitedesign.Id -Principals $config.dmsAdmins -Rights View

Write-Information "Hi 02"

# Create new Modern site with no connection to Office 365 Group
New-PnPTenantSite `
  -Title $config.dmsName `
  -Url $config.dmsUrl `
  -Owner $params.adminUPN `
  -Lcid $config.lcid `
  -Template "STS#3" `
  -TimeZone $config.timeZone `
  -StorageQuota $config.storageQuota `
  -Wait | Out-Default

Write-Information "Hi 03"

Disconnect-PnPOnline  

# Reconnect to site scope
Connect-PnPOnline -Url $config.dmsUrl -Credential $userCredential

# Return Connection Information
Get-PnPConnection | Out-Default

Write-Information "Hi 04"

<# # Enable hidden feture neccessary for Matadata Fields functionality
Enable-PnPFeature -Identity 73ef14b1-13a9-416b-a9b5-ececa2b0604c -Scope Site | Out-Default #>

Write-Information "Hi 05"
Add-PnPFile -Path $path\Assets\avx.jpg -Folder "SiteAssets"
Write-Information "Hi 06"
# Set SiteDesign
#Invoke-SPOSiteDesign -Identity $sitedesign.Id -WebUrl $config.dmsUrl | Out-Default
# Delete SiteDesign definitions
#Remove-SPOSiteDesign -Identity $sitedesign.Id
#Remove-SPOSiteScript -Identity $sitescript.Id

#Read more: https://www.sharepointdiary.com/2016/05/sharepoint-online-enable-versioning-on-all-lists-powershell.html#ixzz6GvmuY4bH

# Set SiteDesign
<# Invoke-PnPSiteDesign -Identity $sitedesign.Id | Out-Default #>

Write-Information "Hi 07"
<# Remove-PnPSiteDesign -Identity $sitedesign.Id -Force
Remove-PnPSiteScript -Identity $navigation.Id -Force
Remove-PnPSiteScript -Identity $baselists.Id -Force
Remove-PnPSiteScript -Identity $documentlibraries.Id -Force
Remove-PnPSiteScript -Identity $contenttypes.Id -Force
Remove-PnPSiteScript -Identity $sitecolumnsxml.Id -Force
Remove-PnPSiteScript -Identity $sitecolumns.Id -Force #>

Write-Information "Hi 08"
<# Set-PnPList -Identity $config.dmsLibraries[0].lib `
-EnableVersioning $True `
-MajorVersions 50000 `
-EnableMinorVersions $True `
-MinorVersions 50000 `
-EnableFolderCreation $False `
-EnableAttachments $False `
-ListExperience NewExperience #>

Write-Information "Hi 09"

#########################################################################
################################## END ##################################
#########################################################################

# Disconnect Tenant
#Disconnect-SPOService

# Disconnect Site
Disconnect-PnPOnline

# Stop Logging
Stop-Transcript