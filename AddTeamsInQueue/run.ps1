# Input bindings are passed in via param block.
param($Timer)

# Variables
$tenantPrefix = $env:TENANT_PREFIX
$tenant = "$tenantPrefix.onmicrosoft.com"
$spoAdminCenter = "https://$tenantPrefix-admin.sharepoint.com/"

$env:PNPPOWERSHELL_UPDATECHECK = "false"

#region Import modules
Import-Module Microsoft.Graph.Authentication -RequiredVersion "1.3.1"
Import-Module Microsoft.Graph.Groups -RequiredVersion "1.3.1"
Import-Module PnP.PowerShell -RequiredVersion "1.3.0"
#endregion Import modules

#region Authentication
try {
    Write-Information "Connecting to PnP Online."
    Connect-PnPOnline -ClientId $env:CLIENT_ID -Url $spoAdminCenter -Thumbprint $env:CERT_THUMBPRINT -tenant $tenant -ErrorAction Stop
    Write-Information "Retrieving the Access Token."
    $accessToken = Get-PnPGraphAccessToken -ErrorAction Stop
    Write-Information "Connecting to Microsoft Graph."
    Connect-MgGraph -AccessToken $accessToken -ErrorAction Stop
}
catch {
    $exceptionMessage = $_.Exception.Message
    Write-Error "Couldn't authenticate successfully. $exceptionMessage"
    return
}
#endregion Authentication

Write-Information "Switching Microsoft Graph profile to beta."
Select-MgProfile -Name "beta" # necessary to get all the teams using the filter below...

# Retrieving all the teams
try {
    Write-Information "Retrieving all the Microsoft Teams."    
    $teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All -Property Id, DisplayName -ErrorAction Stop
}
catch {
    $exceptionMessage = $_.Exception.Message
    Write-Error "Couldn't retrieve all the teams. $exceptionMessage"
    return
}

Write-Information "$($teams.Count) Teams have been retrieved."

# Creating an array of string in the form "<TeamID>,<TeamDisplayName>"
$teamsArray = @()
foreach ($team in $teams) {
    Write-Information "$($team.Id) - $($team.DisplayName)"
    $teamsArray += "$($team.Id),$($team.DisplayName)"
}

# Pushing the teams in the queue
Write-Information "Pushing out results to the queue."
Push-OutputBinding -Name Queue -Value $teamsArray