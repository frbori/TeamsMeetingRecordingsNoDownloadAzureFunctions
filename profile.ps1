# Azure Functions profile.ps1
#
# This profile.ps1 will get executed every "cold start" of your Function App.
# "cold start" occurs when:
#
# * A Function App starts up for the very first time
# * A Function App starts up after being de-allocated due to inactivity
#
# You can define helper functions, run commands, or specify environment variables
# NOTE: any variables defined that are not environment variables will get reset after the first execution

# Authenticate with Azure PowerShell using MSI.
# Remove this if you are not planning on using MSI or Azure PowerShell.
#if ($env:MSI_SECRET) {
#    Disable-AzContextAutosave -Scope Process | Out-Null
#    Connect-AzAccount -Identity
#}

# Uncomment the next line to enable legacy AzureRm alias in Azure PowerShell.
# Enable-AzureRmAlias

# You can also define functions or aliases that can be referenced in any of your PowerShell functions.

<#function LogOutcomeToTableStorage {
    [CmdletBinding()]
    param (
        [Parameter()][string]$TableBindingName,
        [Parameter()][string]$PartitionKey,
        [Parameter()][string]$TeamId,
        [Parameter()][datetime]$StartTime,
        [Parameter()][string]$Outcome,
        [Parameter()][string]$Details,
        [Parameter()][string]$Exceptions,
        [Parameter()][string]$TeamDisplayName,
        [Parameter()][string]$ChannelId
    )
    $endTime = (Get-Date).ToUniversalTime()
    Push-OutputBinding -Name $TableBindingName -Value @{
        PartitionKey    = $PartitionKey
        RowKey          = $TeamId
        StartTime       = $StartTime
        EndTime         = $endTime
        Duration        = ($endTime - $StartTime).Seconds
        Outcome         = $Outcome
        Details         = $Details
        Exceptions      = $Exceptions
        TeamDisplayName = $TeamDisplayName
        ChannelId       = $ChannelId
    }
}
#>
function GetTeamWebsiteUrl() {
    param
    (
        [string][Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$TeamID,
        [string][Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$AccessToken
    )
    $graphUrl = "https://graph.microsoft.com/v1.0/groups/$TeamID/sites/root?`$select=webUrl"
    $headers = @{
        "Content-Type"  = "application/json"
        "Authorization" = "Bearer $AccessToken"
    }
    $response = Invoke-RestMethod -Uri $graphUrl -Headers $headers -Method Get -ContentType "application/json"
    return $response.webUrl
}