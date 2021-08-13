# Clear Console
Clear-Host

# Prepare Log File Name
$dt = Get-Date -Format "yy-MM-dd-HH"
$fname = "log-{0}.txt" -f $dt

# Start Logging
Start-Transcript -Path ..\Logs\$fname -Append -NoClobber

# Get Admin Credentials from params.json
$params = Get-Content ..\params.json | Out-String | ConvertFrom-Json
$orgName = $params.orgName
$password = $params.adminPSW | ConvertTo-SecureString -asPlainText -Force 
$userCredential = New-Object System.Management.Automation.PSCredential($params.adminUPN,$password)

# Connect to SharePoint Administration
$url = "https://{0}-admin.sharepoint.com" -f $orgName
Connect-PnPOnline -Url $url -Credential $userCredential -ErrorAction Stop | Out-Default

$baselistsjson = Get-Content ..\Designs\BaseLists.json -Raw
$baselists = Add-PnPSiteScript -Title "CLOUDEDU Base Lists" -Content $baselistsjson

$navigationjson = Get-Content ..\Designs\Navigation.json -Raw
$navigation = Add-PnPSiteScript -Title "CLOUDEDU Navigation" -Content $navigationjson

$themejson = Get-Content ..\Designs\Theme.json -Raw
$theme = Add-PnPSiteScript -Title "CLOUDEDU Theme" -Content $themejson

# Create new SiteDesign
$scriptname = "CLOUDEDU Team Site ({0})" -f $dt
Add-PnPSiteDesign -Title $scriptname `
                -WebTemplate "64" `
                -SiteScriptIds $baselists.Id, $navigation.Id, $theme.Id `
                -Description "CLOUDEDU Team Site"

# Disconnect Site
Disconnect-PnPOnline

# Stop Logging
Stop-Transcript