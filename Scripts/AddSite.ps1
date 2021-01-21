# Set environment variables
$path = "C:\Repos\AdminScripts"

# Prepare Log File Name
$dt = Get-Date -Format "yy-MM-dd-HH-mm"
$fname = "{0}\Logs\log-{1}.txt" -f $path, $dt 

# Start Logging
Start-Transcript -Path $fname -Append -NoClobber

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

$documentcontenttypesjson = Get-Content $path\Designs\Scripts\DocumentContentType.json -Raw
$documentcontenttypes = Add-PnPSiteScript -Title "AVX-DMS Base Lists Update " -Content $documentcontenttypesjson

$documentlibrariesjson = Get-Content $path\Designs\Scripts\DocumentLibraries.json -Raw
$documentlibraries = Add-PnPSiteScript -Title "AVX-DMS Document Libraries" -Content $documentlibrariesjson

$navigationjson = Get-Content $path\Designs\Scripts\Navigation.json -Raw
$navigation = Add-PnPSiteScript -Title "AVX-DMS Navigation" -Content $navigationjson

$themejson = Get-Content $path\Designs\Scripts\Theme.json -Raw
$theme = Add-PnPSiteScript -Title "AVX-DMS Theme" -Content $themejson

# Create new SiteDesign
$scriptname = "AVX-DMS Team Site ({0})" -f $dt
$sitedesign = Add-PnPSiteDesign -Title $scriptname `
 -WebTemplate "64" `
 -SiteScriptIds $features.Id, $sitecolumns.Id, $sitecolumnsxml.Id, $contenttypes.Id, $baselists.Id, $documentcontenttypes.Id, $documentlibraries.Id, $navigation.Id, $theme.Id `
 -Description "AVX-DMS Team Site"

# Grant Permissions for new SiteDesign
Grant-PnPSiteDesignRights -Identity $sitedesign.Id -Principals $config.dmsAdmins -Rights View

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

Add-PnPFile -Path $path\Assets\avx.jpg -Folder "SiteAssets"

# Disconnect Site
Disconnect-PnPOnline

# Stop Logging
Stop-Transcript