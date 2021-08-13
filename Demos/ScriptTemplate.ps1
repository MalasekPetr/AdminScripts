# Clear Console
Clear-Host

# Prepare Log File Name
$dt = Get-Date -Format "yy-MM-dd-HH"
$fname = "log-{0}.txt" -f $dt

# Start Logging
Start-Transcript -Path ..\Logs\$fname -Append -NoClobber

# Greet Log Readers
Write-Information "Hi there ;)"

# Check Installed Module version
Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable | Select-Object Name,Version,Path | Out-Default
Get-Module -Name PnP.PowerShell -ListAvailable | Select-Object Name,Version,Path | Out-Default

# Possible Module Installation
#Install-Module -Name Microsoft.Online.SharePoint.PowerShell
#Install-Module -Name PnP.PowerShell

# Possible Module Update
#Update-Module -Name Microsoft.Online.SharePoint.PowerShell
#Update-Module -Name PnP.PowerShell

# Get Admin Credentials from params.json
$params = Get-Content ..\params.json | Out-String | ConvertFrom-Json
$orgName = $params.orgName
$password = $params.adminPSW | ConvertTo-SecureString -asPlainText -Force 
$userCredential = New-Object System.Management.Automation.PSCredential($params.adminUPN,$password)

# Connect to SharePoint Administration
$url = "https://{0}-admin.sharepoint.com" -f $orgName
Connect-SPOService -Url $url -Credential $userCredential
Connect-PnPOnline -Url $url -Credential $userCredential

# Return Tenant Information
Get-SPOTenant | Out-Default

# Return Connection Information
Get-PnPConnection | Out-Default


# Script here


# Disconnect Tenant
Disconnect-SPOService

# Disconnect Site
Disconnect-PnPOnline

# Stop Logging
Stop-Transcript