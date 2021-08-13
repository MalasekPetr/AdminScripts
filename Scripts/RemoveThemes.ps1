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

Get-PnPTenantTheme | ForEach-Object {Remove-PnPTenantTheme -Identity $_.Name}

# Disconnect
Disconnect-PnPOnline

# Stop Logging
Stop-Transcript