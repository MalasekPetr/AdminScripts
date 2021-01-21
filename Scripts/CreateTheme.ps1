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

# Get Admin Credentials from params.json
$params = Get-Content $path\params.json -ErrorAction Stop | Out-String | ConvertFrom-Json
$orgName = $params.orgName
$password = $params.adminPSW | ConvertTo-SecureString -asPlainText -Force 
$userCredential = New-Object System.Management.Automation.PSCredential($params.adminUPN,$password)

# Connect to SharePoint Administration
$url = "https://{0}-admin.sharepoint.com" -f $orgName
Connect-SPOService -Url $url -Credential $userCredential -ErrorAction Stop

$themepalette = @{
  "themePrimary" = "#1eb4eb";
  "themeLighterAlt" = "#f5fcfe";
  "themeLighter" = "#d9f2fc";
  "themeLight" = "#b8e8f9";
  "themeTertiary" = "#74d1f3";
  "themeSecondary" = "#38bded";
  "themeDarkAlt" = "#1ba2d3";
  "themeDark" = "#1789b2";
  "themeDarker" = "#116583";
  "neutralLighterAlt" = "#faf9f8";
  "neutralLighter" = "#f3f2f1";
  "neutralLight" = "#edebe9";
  "neutralQuaternaryAlt" = "#e1dfdd";
  "neutralQuaternary" = "#d0d0d0";
  "neutralTertiaryAlt" = "#c8c6c4";
  "neutralTertiary" = "#c6c3c3";
  "neutralSecondary" = "#8e8a8a";
  "neutralPrimaryAlt" = "#595555";
  "neutralPrimary" = "#423f3f";
  "neutralDark" = "#323030";
  "black" = "#252323";
  "white" = "#ffffff";
  }

Add-SPOTheme -Identity "AVX-DMS" -Palette $themepalette -IsInverted $false -Overwrite

# Disconnect Tenant
Disconnect-SPOService

# Stop Logging
Stop-Transcript