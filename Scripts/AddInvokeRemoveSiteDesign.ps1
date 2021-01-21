# Clear Console
Clear-Host

# Prepare Log File Name
$dt = Get-Date -Format "yy-MM-dd-HH"
$fname = "log-{0}.txt" -f $dt

# Start Logging
Start-Transcript -Path Logs/$fname -Append -NoClobber

# Greet Log Readers
Write-Information "Hi there ;)"

# Get Admin Credentials from params.json
$params = Get-Content .\params.json -ErrorAction Stop | Out-String | ConvertFrom-Json
$orgName = $params.orgName
$password = $params.adminPSW | ConvertTo-SecureString -asPlainText -Force 
$userCredential = New-Object System.Management.Automation.PSCredential($params.adminUPN,$password)

# Connect to SharePoint Administration
$url = "https://{0}-admin.sharepoint.com" -f $orgName
Connect-SPOService -Url $url -Credential $userCredential -ErrorAction Stop

# Return Tenant Information
Get-SPOTenant  | Out-Default

#########################################################################
################################ Comment ################################
#########################################################################

$sitescript = Get-Content .\Designs\Scripts\SiteScript.json -Raw | Add-SPOSiteScript -Title "AVX-DMS Site Columns"
$sitedesign = Add-SPOSiteDesign -Title "AVX-DMS" -WebTemplate "64" -SiteScripts $sitescript.Id -Description "AVX-DMS Base Setup"
Grant-SPOSiteDesignRights -Identity $sitedesign.Id -Principals $params.dmsGroup -Rights View
Invoke-SPOSiteDesign -Identity $sitedesign.Id -WebUrl $params.dmsSite | Out-Default
Remove-SPOSiteDesign -Identity $sitedesign.Id

#########################################################################
################################## END ##################################
#########################################################################

# Disconnect Tenant
Disconnect-SPOService

# Stop Logging
Stop-Transcript