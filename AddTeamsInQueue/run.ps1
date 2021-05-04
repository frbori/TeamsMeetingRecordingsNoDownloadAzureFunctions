# Input bindings are passed in via param block.
param($Timer)

#region Variables
$tenantPrefix = $env:TENANT_PREFIX
$tenant = "$tenantPrefix.onmicrosoft.com"
$spoAdminCenter = "https://$tenantPrefix-admin.sharepoint.com/"
$teamsCreationDateStartText = $env:TEAMS_CREATION_DATE_START
$teamsCreationDateStart = $null
$teamsCreationDateEndText = $env:TEAMS_CREATION_DATE_END
$teamsCreationDateEnd = $null
#endregion Variables

$env:PNPPOWERSHELL_UPDATECHECK = "false"

#region Import modules
Import-Module Microsoft.Graph.Authentication -RequiredVersion "1.3.1"
Import-Module Microsoft.Graph.Groups -RequiredVersion "1.3.1"
Import-Module PnP.PowerShell -RequiredVersion "1.3.0"
#endregion Import modules

#region Authentication
try {
    Write-Debug "Connecting to PnP Online."
    Connect-PnPOnline -ClientId $env:CLIENT_ID -Url $spoAdminCenter -Thumbprint $env:CERT_THUMBPRINT -tenant $tenant -ErrorAction Stop
    Write-Debug "Retrieving the Access Token."
    $accessToken = Get-PnPGraphAccessToken -ErrorAction Stop
    Write-Debug "Connecting to Microsoft Graph."
    Connect-MgGraph -AccessToken $accessToken -ErrorAction Stop | Out-Null
}
catch {
    Write-Error "Couldn't authenticate successfully."
    throw $_.Exception
}
#endregion Authentication

Write-Debug "Switching Microsoft Graph profile to beta."
Select-MgProfile -Name "beta" # necessary to get all the teams using the filter below...

# Retrieving all the teams
try {
    Write-Debug "Retrieving all the Microsoft Teams."
    $teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All -Property Id, DisplayName, CreatedDateTime -ErrorAction Stop
}
catch {
    Write-Error "Couldn't retrieve all the teams."
    throw $_.Exception
}

Write-Information "$($teams.Count) Teams have been retrieved."

#region Handling filters on teams creation date
if (![string]::IsNullOrEmpty($teamsCreationDateStartText) -and ![string]::IsNullOrEmpty($teamsCreationDateEndText)) {
    try {
        $teamsCreationDateStart = [Datetime]::ParseExact($teamsCreationDateStartText, 'dd/MM/yyyy', $null)
        $teamsCreationDateEnd = [Datetime]::ParseExact($teamsCreationDateEndText, 'dd/MM/yyyy', $null)
        Write-Information "Applying filter on teams creation date: $teamsCreationDateStartText <= creation date <= $teamsCreationDateEndText"
        $teams = $teams | Where-Object { $_.CreatedDateTime.Date -ge $teamsCreationDateStart -and $_.CreatedDateTime.Date -le $teamsCreationDateEnd }
    }
    catch {
        Write-Error "Couldn't convert start or end team creation date."
        throw $_.Exception
    }
}
elseif (![string]::IsNullOrEmpty($teamsCreationDateStartText)) {
    try {
        $teamsCreationDateStart = [Datetime]::ParseExact($teamsCreationDateStartText, 'dd/MM/yyyy', $null)
        Write-Information "Applying filter on teams creation date: $teamsCreationDateStartText <= creation date"
        $teams = $teams | Where-Object { $_.CreatedDateTime.Date -ge $teamsCreationDateStart }
    }
    catch {
        Write-Error "Couldn't convert start creation date."
        throw $_.Exception
    }
}
elseif (![string]::IsNullOrEmpty($teamsCreationDateEndText)) {
    try {
        $teamsCreationDateEnd = [Datetime]::ParseExact($teamsCreationDateEndText, 'dd/MM/yyyy', $null)
        Write-Information "Applying filter on teams creation date: creation date <= $teamsCreationDateEndText"
        $teams = $teams | Where-Object { $_.CreatedDateTime.Date -le $teamsCreationDateEnd }
    }
    catch {
        Write-Error "Couldn't convert end team creation date."
        throw $_.Exception
    }
}
else {
    Write-Information "No filter condition on teams creation date will be applied."
}
#endregion Handling filters on teams creation date

Write-Information "$($teams.Count) Teams will be processed."

# Creating an array of string in the form "<TeamID>,<TeamDisplayName>"
$teamsArray = @()
foreach ($team in $teams) {
    Write-Information "Team id: $($team.Id) | Created on: $($team.CreatedDateTime.ToShortDateString()) | Display name: $($team.DisplayName)"
    $teamsArray += "$($team.Id),$($team.DisplayName)"
}

# Pushing the teams in the queue
Write-Debug "Pushing out results to the queue."
Push-OutputBinding -Name Queue -Value $teamsArray