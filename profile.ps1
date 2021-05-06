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

function GetTeamWebsiteUrl() {
    param
    (
        [string][Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$TeamID,
        [string][Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$AccessToken
    )
    Write-Debug "Retrieving Team '$teamDisplayName' SharePoint site url."
    $graphUrl = "https://graph.microsoft.com/v1.0/groups/$TeamID/sites/root?`$select=webUrl"
    $headers = @{
        "Content-Type"  = "application/json"
        "Authorization" = "Bearer $AccessToken"
    }
    try {
        $response = Invoke-RestMethod -Uri $graphUrl -Headers $headers -Method Get -ContentType "application/json"
    }
    catch {
        $outcome = "Team error."
        $details = "An error while retrieving the SharePoint Team Site url occurred."
        Write-Error "$outcome $details"
        throw $_.Exception
    }
    return $response.webUrl
}
function GetRestrictedView() {
    param
    (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$PnPConnection,
        [string][Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$SPSiteUrl,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$SPDocumentLibrary,
        [switch]$PrivateChannel
    )
    $restrictedView = $null
    try {
        Write-Debug "Retrieving Restricted View SharePoint permission level."
        $roleDefs = Get-PnPRoleDefinition -Connection $PnPConnection
        $restrictedView = $roleDefs | Where-Object { $_.RoleTypeKind -eq "RestrictedReader" }
        if ($null -eq $restrictedView) {
            try {
                $spDocumentsListId = $SPDocumentLibrary.Id
                $uri = "$SPSiteUrl/_api/web/Lists(@a1)/GetItemById(@a2)/GetSharingInformation?@a1=%27%7B$spDocumentsListId%7D%27&@a2=%271%27&`$Expand=sharingLinkTemplates"
                Invoke-PnPSPRestMethod -Method Post -Url $uri -ContentType "application/json" -Content @{} | Out-Null
    
                $roleDefs = Get-PnPRoleDefinition -Connection $PnPConnection
                $restrictedView = $roleDefs | Where-Object { $_.RoleTypeKind -eq "RestrictedReader" }
                $restrictedView.GetType() | Out-Null
            }
            catch {
                if ($null -eq $PrivateChannel) {
                    $outcome = "Team error."
                    $details = "An error while triggering Restricted View permission level occurred."
                    Write-Error "$outcome $details"
                    throw $_.Exception
                }
                else {
                    $outcome = "Channel error."
                    $details = "An error while retrieving SharePoint Role Definitions for private channel occurred. "
                    $exceptions = $_.Exception.Message
                    Write-Warning "$outcome $details $exceptions"
                }
            }
        }
    }
    catch {
        $outcome = "Team error."
        $details = "An error while handling the custom permission level occurred."
        Write-Error "$outcome $details"
        throw $_.Exception
    }   
    return $restrictedView
}
function GetSPDocumentsLibrary() {
    param
    (
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()]$PnPConnection,
        [switch]$PrivateChannel
    )
    $spLibrary = $null
    try {
        $documentsListName = "Documents"
        Write-Debug "Retrieving SharePoint '$documentsListName' document library."
        $spLibrary = Get-PnPList -Identity $documentsListName -Connection $PnPConnection
        if ($null -eq $spLibrary) {
            $documentsListName = "Documenti"
            Write-Debug "Retrieving SharePoint '$documentsListName' document library."
            $spLibrary = Get-PnPList -Identity $documentsListName -ErrorAction Stop -Connection $PnPConnection
        }
    }
    catch {
        if ($null -eq $PrivateChannel) {
            $outcome = "Team error."
            $details = "An error while retrieving the 'Documents' document library occurred."
            Write-Error "$outcome $details"
            throw $_.Exception
        }
        else {
            $outcome = "Channel error."
            $details = "An error while retrieving the 'Documents' document library for private channel occurred."
            $exceptions = $_.Exception.Message
            Write-Error "$outcome $details $exceptions"
        }
    }
    return $spLibrary
}