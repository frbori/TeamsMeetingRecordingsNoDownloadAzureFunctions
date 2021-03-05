# TeamsMeetingRecordingsNoDownload
This solution aims to prevent that Teams channel meeting recordings stored into SharePoint can be downloaded by team members.

By default, Teams channel meeting recordings are saved into the SharePoint site associated to the team and team members are added to the default SharePoint members group; this gives them Edit permission on all the SharePoint contents, including the possibility of downloading files.

This solution basically changes the permissions assigned to the default SharePoint members group on the folders containing the recordings files (it breaks the permissions inheritance and assigns the desired permissions).

## Solution components
The solution is mainly composed by the follwing two components:
- an Azure AD App Registration
- an Azure Function App (built on PowerShell)

### Azure AD App Registration
The Azure AD App Registration is required to allow the two Azure functions to get authenticated and authorized. The Access Token is retrieved by specifying *client id* and *certificate thumbprint*.
The required (application) permissions to assign to the app are:
- Group.Read.All
- TeamSettings.Read.All
- Sites.FullControl.All

### Azure Function App
The Azure Function App contains two Azure functions:
- AddTeamsInQueue
- ProcessTeam

The PowerShell modules used by the solution are:
- PnP.PowerShell (loaded as  managed dependency)
- Microsoft.Graph.Authentication (loaded as  managed dependency)
- Microsoft.Graph.Teams (loaded as  managed dependency)
- Microsoft.Graph.Groups (loaded as  managed dependency)
- Microsoft.Graph.Files (explicitely inclueded in the solution due to issues when trying to load it as managed dependency)

#### AddTeamsInQueue
This is a scheduled function (time triggered) that lists all the teams in the tenant and, for each of them, adds a message into an Azure Queue called *teamsqueue*.
Each message contains the team id and the team display name separated by a comma (eg.: *332cfb44-c4b5-4513-8404-72f3ed82e6d1,HR*).

#### ProcessTeam
This is a queue triggered function (it triggers when new messages get into the *teamsqueue*) that processes the specific team.
The team processing entails:
- retrieving all the team channels (both standard and private)
- retrieving the channels folders
- creating the "Recordings" folder inside the channels folders (if not already created)
- changing the permissions on the **Recordings** folders so that team members won't be able to download files stored into those folders

This function logs processing outcome into an Azure storage table called *Log*.

## How to deploy
Deploying the solution on your tenant comprises 3 main steps:
1. Registering an App in Azure AD
2. Creating the required Azure resources (Resource Group, Storage Account, Function App)
3. Deploy the zip package to the Function App

You can complete the major part of those steps programmatically by using the sample script below (it requires [*PnP.PowerShell*](https://pnp.github.io/powershell/) and [*AZ*](https://docs.microsoft.com/en-us/powershell/azure/new-azureps-module-az) PowerShell modules):
```powershell
#region VARIABLES
$tenantPrefix = "*<tenantPrefix>*"    # this is the part just before .onmicrosoft.com
$appRegistrationName = "*<appName>*"
$certsOutputPath = "*<folderFullPath>*"    # this folder shoud be already existing
$resourceGroupName = "*<resourceGroupName>*"
$storageAccountName = "*<storageAccountName>*"
$location = "West Europe"
$functionAppName = "*<functionAppName>*"
$createRecordingsFolder = "true"
$zipPackage = "*<zipPackageFullPath>*"
$subscriptionName = ""    # leave blank if you have juts one subscription, otherwise specify which subscription you want to use
#endregion VARIABLES

#region AZURE APP REGISTRATION
Write-Host "Registering app '$appRegistrationName' in Azure AD"
$certPassword = Read-Host -Prompt "Enter certificate password" -AsSecureString
$appRegistration = Register-PnPAzureADApp -ApplicationName $appRegistrationName -Tenant "$tenantPrefix.onmicrosoft.com" -Store CurrentUser `
    -Scopes "MSGraph.Group.Read.All", "SPO.Sites.FullControl.All" `
    -DeviceLogin -OutPath $certsOutputPath -CertificatePassword $certPassword
$clientId = $appRegistration.'AzureAppId/ClientId'
$certThumbprint = $appRegistration.'Certificate Thumbprint'
Write-Host "Remember to add also scope 'MSGraph.TeamSettings.Read.All' to app '$appRegistrationName' and grant admin consent for those permissions" -ForegroundColor Yellow
#endregion AZURE APP REGISTRATION

Connect-AzAccount
#region RETRIEVING/CREATING AZURE RESOURCE GROUP, STORAGE ACCOUNT, FUNCTION APP
If (![string]::IsNullOrEmpty($subscriptionName))
{
    Write-Host "Setting Azure context to '$subscriptionName' subscription"
    Set-AzContext -Subscription $subscriptionName
}

Write-Host "Retrieving Resource Group '$resourceGroupName'"
$resourceGroup = Get-AzResourceGroup -Name $resourceGroupName -ErrorAction SilentlyContinue
if ($null -eq $resourceGroup)
{
    Write-Host "Resource Group '$resourceGroupName' is not present, creating it"
    New-AzResourceGroup -Name $resourceGroupName -Location $location
}

Write-Host "Retrieving Storage Account '$storageAccountName'"
$storageAccount = Get-AzStorageAccount -ResourceGroupName $resourceGroupName | ? {$_.StorageAccountName -eq $storageAccountName}
if ($null -eq $storageAccount)
{
    Write-Host "Storage Account '$storageAccountName' is not present, creating it"
    New-AzStorageAccount -ResourceGroupName $resourceGroupName -Name $storageAccountName -Location $location -SkuName Standard_LRS -Kind Storage
}

$appSettings = @{
        WEBSITE_RUN_FROM_PACKAGE = "1"
        CLIENT_ID = $clientId
        CERT_THUMBPRINT = $certThumbprint
        WEBSITE_LOAD_CERTIFICATES = $certThumbprint
        CREATE_RECORDINGS_FOLDER = $createRecordingsFolder
        TENANT_PREFIX = $tenantPrefix 
}

Write-Host "Retrieving Function App '$functionAppName'"
$functionApp = Get-AzFunctionApp -Name $functionAppName -ResourceGroupName $resourceGroupName -ErrorAction SilentlyContinue
if ($null -eq $functionApp)
{
    Write-Host "Function App '$functionAppName' is not present, creating it"
    New-AzFunctionApp -ResourceGroupName $resourceGroupName -Name $functionAppName -Location $location -Runtime PowerShell -OSType Windows -RuntimeVersion 7.0 -FunctionsVersion 3 -StorageAccountName $storageAccountName -AppSetting $appSettings
}
Write-Host "Remember to upload $appRegistrationName.pfx certificate to the Function App '$functionAppName'" -ForegroundColor Yellow
#endregion RETRIEVING/CREATING AZURE RESOURCE GROUP, STORAGE ACCOUNT, FUNCTION APP

# PUBLISHING THE ZIP PACKAGE
Write-Host "Publishing the zip package at '$zipPackage' to Function App '$functionAppName'"
Publish-AzWebapp -ResourceGroupName $resourceGroupName -Name $functionAppName -ArchivePath $zipPackage

Disconnect-AzAccount
```
The remaining manual steps are (as highlighted by the script itself):
- add *MSGraph.TeamSettings.Read.All* permission to the Azure AD app and grant admin consent for all the assigned permissions
- upload the private key certificate (pfx) in the Function App:
    - locate the certificate generated during the app registration (*$certsOutputPath* parameter)
    - navigate to the Function App
    - select **TLS/SSL settings**
    - select **Private Key Certificates (.pfx)** tab
    - click on **+ Upload Certificate**
    - select the certificate, enter the password (chosen during app registration) and click **Upload**
