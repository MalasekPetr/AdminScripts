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
  
$tname = "CLOUDEDU-{0}" -f $dt
Add-SPOTheme -Identity $tname -Palette $themepalette -IsInverted $false -Overwrite

Write-Host "Theme name: $tname"

# Disconnect Tenant
Disconnect-SPOService

# Stop Logging
Stop-Transcript