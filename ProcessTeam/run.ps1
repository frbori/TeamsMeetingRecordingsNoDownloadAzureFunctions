using namespace System.Net

# Input bindings are passed in via param block.
param([string] $QueueItem, $TriggerMetadata)

#region Variables
$teamId = $QueueItem.Split(",")[0]
$teamDisplayName = $QueueItem.Split(",")[1]
$tenantPrefix = $env:TENANT_PREFIX
$spoRootUrl = "https://$env:TENANT_PREFIX.sharepoint.com"
$tenant = "$tenantPrefix.onmicrosoft.com"
$spoAdminCenter = "https://$tenantPrefix-admin.sharepoint.com/"
$createRecordingsFolder = "true" -eq $env:CREATE_RECORDINGS_FOLDER
$teamChannels = $null    
$web = $null
$outcome = "Completed with no error."
$details = ""
$exceptions = ""
$restrictedView = $null
$restrictedViewTeamSite = $null
$spDocumentsListId = $null
$channelFolderUrlObj = $null
$channelFolderUrl = $null
$channelFolder = $null
$channelRecFolder = $null
$privateChannelSiteUrl = $null
$teamSiteConn = $null
$connectionToUse = $null
$members = $null
$visitors = $null
$owners = $null
$ownersRole = $null
$channelErrorOccurred = $false
#endregion Variables

$env:PNPPOWERSHELL_UPDATECHECK = "false"

#region Import modules
Import-Module Microsoft.Graph.Authentication -RequiredVersion "1.3.1"
Import-Module Microsoft.Graph.Groups -RequiredVersion "1.3.1"
Import-Module Microsoft.Graph.Teams -RequiredVersion "1.3.1"
Import-Module PnP.PowerShell -RequiredVersion "1.3.0"
#endregion Import modules

Write-Information "Processing team '$teamDisplayName' - $teamId."

#region Authentication
try {
    Write-Debug "Connecting to PnP Online - $spoAdminCenter"
    Connect-PnPOnline -ClientId  $env:CLIENT_ID -Url $spoAdminCenter -Thumbprint $env:CERT_THUMBPRINT -tenant $tenant -ErrorAction Stop
    
    Write-Debug "Retrieving the Access Token."
    $accessToken = Get-PnPGraphAccessToken -ErrorAction Stop
    
    Write-Debug "Connecting to Microsoft Graph."
    Connect-MgGraph -AccessToken $accessToken -ErrorAction Stop | Out-Null

    Write-Debug "Disconnecting from PnP Online - $spoAdminCenter"
    Disconnect-PnPOnline
}
catch {
    $outcome = "Team error."
    $details = "An authentication issue occurred."
    Write-Error "$outcome $details"
    throw $_.Exception    
}
#endregion Authentication

# Switching Microsoft Graph profile to beta endpoing
Write-Debug "Switching Microsoft Graph profile to beta."
Select-MgProfile "beta"

#region Getting the SharePoint site url associated to the team/o365 group
try {
    Write-Debug "Retrieving Team '$teamDisplayName' SharePoint site url."
    $spSiteUrl = GetTeamWebsiteUrl -TeamID $teamId -AccessToken $accessToken
}
catch {
    $outcome = "Team error."
    $details = "An error while retrieving the SharePoint Team Site url occurred."
    Write-Error "$outcome $details"
    throw $_.Exception
}
#endregion Getting the SharePoint site url associated to the team/o365 group

#region Connecting to the SharePoint site associated to the team/o365 group
try {
    Write-Debug "Connecting to PnP Online - $spSiteUrl"
    $teamSiteConn = Connect-PnPOnline -ClientId $env:CLIENT_ID -Url $spSiteUrl -Thumbprint $env:CERT_THUMBPRINT -tenant $tenant -ErrorAction Stop -ReturnConnection
}
catch {
    $outcome = "Team error."
    $details = "An error while connecting to the SharePoint Team Site."
    Write-Error "$outcome $details"
    throw $_.Exception
}
#endregion Connecting to the SharePoint site associated to the team/o365 group
    
#region Getting the current web site and associated SharePoint groups
try {
    Write-Debug "Retrieving SharePoint Team Web Site (SPWeb object) and associated default SharePoint groups."
    $web = Get-PnPWeb -Includes AssociatedMemberGroup, AssociatedVisitorGroup, AssociatedOwnerGroup -ErrorAction Stop -Connection $teamSiteConn
    $teamSiteMembers = $web.AssociatedMemberGroup
    $teamSitevisitors = $web.AssociatedVisitorGroup
    $teamSiteOwners = $web.AssociatedOwnerGroup
    
    Write-Debug "Retrieving default SharePoint Owner group associated permission."
    $teamSiteOwnersRole = Get-PnPGroupPermissions -Identity $teamSiteOwners -ErrorAction Stop -Connection $teamSiteConn | Where-Object { $_.Hidden -eq $false }
}
catch {
    $outcome = "Team error."
    $details = "An error while retrieving the SharePoint web site and/or the associated SharePoint groups occurred."
    Write-Error "$outcome $details"
    throw $_.Exception    
}   
#endregion Getting the current web site and associated SharePoint groups
        
#region Getting Documents document library
try {
    $documentsListName = "Documents"
    Write-Debug "Retrieving SharePoint '$documentsListName' document library."
    $spLibrary = Get-PnPList -Identity $documentsListName -Connection $teamSiteConn
    if ($null -eq $spLibrary) {
        $documentsListName = "Documenti"
        Write-Debug "Retrieving SharePoint '$documentsListName' document library."
        $spLibrary = Get-PnPList -Identity $documentsListName -ErrorAction Stop -Connection $teamSiteConn
    }
}
catch {
    $outcome = "Team error."
    $details = "An error while retrieving the 'Documents' document library occurred."
    Write-Error "$outcome $details"
    throw $_.Exception
}
#endregion Getting Documents document library

#region Handling Restricted View permission level
try {
    Write-Debug "Retrieving Restricted View SharePoint permission level."
    $roleDefs = Get-PnPRoleDefinition -Connection $teamSiteConn
    $restrictedViewTeamSite = $roleDefs | Where-Object { $_.RoleTypeKind -eq "RestrictedReader" }
    if ($null -eq $restrictedViewTeamSite) {
        try {
            $spDocumentsListId = $spLibrary.Id
            $uri = "$spSiteUrl/_api/web/Lists(@a1)/GetItemById(@a2)/GetSharingInformation?@a1=%27%7B$spDocumentsListId%7D%27&@a2=%271%27&`$Expand=sharingLinkTemplates"
            Invoke-PnPSPRestMethod -Method Post -Url $uri -ContentType "application/json" -Content @{} | Out-Null

            $roleDefs = Get-PnPRoleDefinition -Connection $teamSiteConn
            $restrictedViewTeamSite = $roleDefs | Where-Object { $_.RoleTypeKind -eq "RestrictedReader" }
        }
        catch {
            $outcome = "Team error."
            $details = "An error while triggering Restricted View permission level occurred."
            Write-Error "$outcome $details"
            throw $_.Exception
        }
    }
}
catch {
    $outcome = "Team error."
    $details = "An error while handling the custom permission level occurred."
    Write-Error "$outcome $details"
    throw $_.Exception
}   
#endregion Handling custom permission level
        
#region Retrieving all the channels
try {
    Write-Debug "Retrieving all the team '$teamDisplayName' channels."
    $teamChannels = Get-MgTeamChannel -TeamId $teamId -ErrorAction Stop
}
catch {
    $outcome = "Team error."
    $details = "An error while retrieving the public team channels occurred."
    Write-Error "$outcome $details"
    throw $_.Exception
}
#endregion Retrieving all the public channels

#region Processing each channel
foreach ($channel in $teamChannels) {
    Write-Information "Processing channel '$($channel.DisplayName)' - $($channel.Id)."
    #region Getting channel folder url information
    try {
        $channelFolderUrlObj = Get-MgTeamChannelFileFolder -TeamId $teamId -ChannelId $channel.Id -ErrorAction Stop
    }
    catch {
        $channelErrorOccurred = $true
        $outcome = "Channel error."
        $details = "An error while retrieving channel folder url for channel '$($channel.DisplayName)' occurred. "
        $exceptions = $_.Exception.Message
        Write-Warning "$outcome $details $exceptions"
        continue
    }
    #endregion Getting channel folder url information
    if ($channel.MembershipType -eq "private") {
        #region Handling Private Channel specific objects - SPWeb, AssociatedGroups, Roles
        # Connecting to the Private Channel SharePoint site
        try {
            $privateChannelSiteUrl = $channelFolderUrlObj.WebUrl.Substring(0, $channelFolderUrlObj.WebUrl.LastIndexOf("/", $channelFolderUrlObj.WebUrl.LastIndexOf("/") - 1))
            $connectionToUse = Connect-PnPOnline -ClientId $env:CLIENT_ID -Url $privateChannelSiteUrl -Thumbprint $env:CERT_THUMBPRINT -tenant $tenant -ErrorAction Stop -ReturnConnection
        }
        catch {
            $channelErrorOccurred = $true
            $outcome = "Channel error."
            $details = "An error while connecting to private channel site for channel '$($channel.DisplayName)' occurred. "
            $exceptions = $_.Exception.Message
            Write-Warning "$outcome $details $exceptions"
            continue
        }
        # Retrieving Private Channel SPWeb, Associated Groups and associated permissions
        try {
            $web = Get-PnPWeb -Includes AssociatedMemberGroup, AssociatedVisitorGroup, AssociatedOwnerGroup -ErrorAction Stop -Connection $connectionToUse 
            $members = $web.AssociatedMemberGroup
            $visitors = $web.AssociatedVisitorGroup
            $owners = $web.AssociatedOwnerGroup
            $ownersRole = Get-PnPGroupPermissions -Identity $owners -ErrorAction Stop -Connection $connectionToUse | Where-Object { $_.Hidden -eq $false }
        }
        catch {
            $channelErrorOccurred = $true
            $outcome = "Channel error."
            $details = "An error while retrieving the SharePoint web site and/or the associated SharePoint groups for private channel '$($channel.DisplayName)' occurred. "
            $exceptions = $_.Exception.Message
            Write-Warning "$outcome $details $exceptions"
            continue
        }
        # Retrieving Private Channel document library
        try {
            $documentsListName = "Documents"
            Write-Debug "Retrieving SharePoint '$documentsListName' document library for private channel '$($channel.DisplayName)'."
            $spLibrary = Get-PnPList -Identity $documentsListName -Connection  $connectionToUse
            if ($null -eq $spLibrary) {
                $documentsListName = "Documenti"
                Write-Debug "Retrieving SharePoint '$documentsListName' document library for private channel '$($channel.DisplayName)'."
                $spLibrary = Get-PnPList -Identity $documentsListName -ErrorAction Stop -Connection  $connectionToUse
            }
        }
        catch {
            $channelErrorOccurred = $true
            $outcome = "Channel error."
            $details = "An error while retrieving the 'Documents' document library for private channel '$($channel.DisplayName)' occurred."
            $exceptions = $_.Exception.Message
            Write-Error "$outcome $details $exceptions"
            continue
        }
        # Retrieving Private Channel permission levels
        try {
            Write-Debug "Retrieving Restricted View SharePoint permission level for private channel '$($channel.DisplayName)'."
            $roleDefs = Get-PnPRoleDefinition -Connection $connectionToUse
            $restrictedView = $roleDefs | Where-Object { $_.RoleTypeKind -eq "RestrictedReader" }
        }
        catch {
            $channelErrorOccurred = $true
            $outcome = "Channel error."
            $details = "An error while retrieving SharePoint Role Definitions for private channel '$($channel.DisplayName)' occurred. "
            $exceptions = $_.Exception.Message
            Write-Warning "$outcome $details $exceptions"
            continue
        }
        # Handling the case when the Private Channel Restricted View permission level is not already there
        if ($null -eq $restrictedView) {
            try {
                $spDocumentsListId = $spLibrary.Id
                $uri = "$privateChannelSiteUrl/_api/web/Lists(@a1)/GetItemById(@a2)/GetSharingInformation?@a1=%27%7B$spDocumentsListId%7D%27&@a2=%271%27&`$Expand=sharingLinkTemplates"
                Invoke-PnPSPRestMethod -Method Post -Url $uri -ContentType "application/json" -Content @{} -Connection $connectionToUse | Out-Null

                $roleDefs = Get-PnPRoleDefinition -Connection $connectionToUse
                $restrictedView = $roleDefs | Where-Object { $_.RoleTypeKind -eq "RestrictedReader" }
            }
            catch {
                $channelErrorOccurred = $true
                $outcome = "Channel error."
                $details = "An error while triggering Restricted View permission level for private channel '$($channel.DisplayName)' occurred."
                $exceptions = $_.Exception.Message
                Write-Error "$outcome $details $exceptions"
                continue
            }
        }
    }
    #endregion Handling Private Channel specific objects
    else {
        # it's a standard channel (not private), use the Team Site related objects...
        $connectionToUse = $teamSiteConn
        $owners = $teamSiteOwners
        $ownersRole = $teamSiteOwnersRole
        $members = $teamSiteMembers
        $visitors = $teamSitevisitors
        $restrictedView = $restrictedViewTeamSite
    }
    #region Getting channel folder object
    try {
        $channelFolderUrl = [System.Web.HttpUtility]::UrlDecode($channelFolderUrlObj.WebUrl).Replace("$spoRootUrl", "")
        $channelFolder = Get-PnPFolder -Url $channelFolderUrl -ErrorAction Stop -Connection $connectionToUse
    }
    catch {
        $channelErrorOccurred = $true
        $outcome = "Channel error."
        $details = "An error while retrieving channel folder for channel '$($channel.DisplayName)' occurred. "
        $exceptions = $_.Exception.Message
        Write-Warning "$outcome $details $exceptions"
        continue
    }
    #endregion Getting channel folder object
    #region Retrieving channel Recordings folder
    try {
        $channelRecFolder = Get-PnPFolder -Url "$channelFolderUrl/Recordings" -ErrorAction Stop -Connection $connectionToUse
    }
    catch {
        $channelRecFolder = $null
    }
    #endregion Retrieving channel Recordings folder
    #region Creating channel Recordings folder if needed
    if (($null -eq $channelRecFolder) -and ($true -eq $createRecordingsFolder)) {
        try {
            $channelFolder.AddSubFolder("Recordings", $null)
            $channelFolder.Context.ExecuteQuery()
        }
        catch {
            $channelErrorOccurred = $true
            $outcome = "Channel error."
            $details = "An error while creating 'Recordings' folder for channel '$($channel.DisplayName)' occurred. "
            $exceptions = $_.Exception.Message
            Write-Warning "$outcome $details $exceptions"
            continue
        }
    }
    elseif ($null -eq $channelRecFolder) {
        $outcome = "Channel error."
        $details = "An error while retrieving 'Recordings' folder for channel '$($channel.DisplayName)' occurred. "
        $exceptions = $_.Exception.Message
        Write-Warning "$outcome $details $exceptions"
        continue
    }
    #endregion Creating channel Recordings folder if needed
    #region Setting custom permissions on channel Recordings folder
    try {
        Set-PnPFolderPermission -List $documentsListName -Identity "$channelFolderUrl/Recordings" -Group $owners -AddRole $ownersRole.Name -Connection $connectionToUse -ClearExisting
        Set-PnPFolderPermission -List $documentsListName -Identity "$channelFolderUrl/Recordings" -Group $visitors -AddRole $restrictedView.Name -Connection $connectionToUse
        Set-PnPFolderPermission -List $documentsListName -Identity "$channelFolderUrl/Recordings" -Group $members -AddRole $restrictedView.Name -Connection $connectionToUse
    }
    catch {
        $channelErrorOccurred = $true
        $outcome = "Channel error."
        $details = "An error while changing permissions on 'Recordings' folder for channel '$($channel.DisplayName)' occurred. "
        $exceptions = $_.Exception.Message
        Write-Warning "$outcome $details $exceptions"
        continue
    }
    #endregion Setting custom permissions on channel Recordings folder
    
    $outcome = "Channel processed successfully."
    $details = ""
    $exceptions = ""
    Write-Debug "Channel '$($channel.DisplayName)' processed successfully."
}    
#endregion Processing each channel

if ($false -eq $channelErrorOccurred) {
    $outcome = "Team processed successfully."
    $details = "All team channels have been successfully processed."
    Write-Information "$outcome $details"
}
else {
    $outcome = "Team processed with channels errors."
    $details = "Some channels have been processed with errors."
    Write-Warning "$outcome $details"
}