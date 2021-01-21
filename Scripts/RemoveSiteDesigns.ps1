# Set environment variables
$path = "C:\Repos\AdminScripts"

# Check Installed Module versionS
Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable | Select-Object Name,Version

# Get Admin Credentials from params.json
$params = Get-Content $path\params.json -ErrorAction Stop | Out-String | ConvertFrom-Json
$orgName = $params.orgName
$password = $params.adminPSW | ConvertTo-SecureString -asPlainText -Force 
$userCredential = New-Object System.Management.Automation.PSCredential($params.adminUPN,$password)

# Connect to SharePoint Administration
$url = "https://{0}-admin.sharepoint.com" -f $orgName
Connect-PnPOnline -Url $url -Credential $userCredential -ErrorAction Stop | Out-Default

Get-PnPSiteDesign | ForEach-Object {Remove-PnPSiteDesign -Identity $_.Id -Force}
Get-PnPSiteScript | ForEach-Object {Remove-PnPSiteScript -Identity $_.Id -Force}

# Return Connection Information
Get-PnPConnection

# Disconnect
Disconnect-PnPOnline