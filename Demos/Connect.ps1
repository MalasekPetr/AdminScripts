# Set environment variables
$path = "C:\Repos\AdminScripts"

# Check Installed Module version
Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable | Select-Object Name,Version | Out-Default
Get-Module -Name SharePointPnPPowerShellOnline -ListAvailable | Select-Object Name,Version | Out-Default

# Possible Module Installation
#Install-Module -Name Microsoft.Online.SharePoint.PowerShell
#Install-Module -Name SharePointPnPPowerShellOnline

# Possible Module Update
#Update-Module -Name Microsoft.Online.SharePoint.PowerShell
#Update-Module -Name SharePointPnPPowerShellOnline

# Get Admin Credentials
#$adminUPN="malach@onlinecourse.cz"
#$userCredential = Get-Credential -UserName $adminUPN -Message "Type the password:"
#$orgName="onlinecourse"

# Get Admin Credentials from params.json
$params = Get-Content $path\params.json -ErrorAction Stop | Out-String | ConvertFrom-Json
$orgName = $params.orgName
$password = $params.adminPSW | ConvertTo-SecureString -asPlainText -Force 
$userCredential = New-Object System.Management.Automation.PSCredential($params.adminUPN,$password)

# Connect to SharePoint Administration
$url = "https://{0}-admin.sharepoint.com" -f $orgName
#Connect-SPOService -Url $url -Credential $userCredential -ErrorAction Stop
# Connect to SharePoint Site
Connect-PnPOnline -Url $url -Credential $userCredential -ErrorAction Stop | Out-Default
Connect-PnPOnline -Url https://onlinecourse.sharepoint.com/teams/avx-dms-003 -Credential $userCredential -ErrorAction Stop | Out-Default

# Return Tenant Information
#Get-SPOTenant | Out-Default

# Return Connection Information
Get-PnPConnection | Out-Default

# Disconnect Tenant
#Disconnect-SPOService

# Disconnect Site
#Remove-PnPSiteDesign 0da17435-01c0-4b4d-bc9c-168c2bb94d69 -Force
#Get-PnPSiteScript | ForEach-Object {Remove-PnPSiteScript -Identity $_.Id -Force} 
#Disconnect-PnPOnline
