# Teams Meeting Channel Recordings No Download
This solution aims to prevent that Teams channel meeting recordings stored into SharePoint can be downloaded by team members.

By default, Teams channel meeting recordings are now saved into the SharePoint site associated to the team and team members are added to the default SharePoint members group; this gives them Edit permission on all the SharePoint contents, including the possibility of downloading files.

This solution basically changes the permissions assigned to the default SharePoint members group on the folders containing the recordings files (it breaks the permissions inheritance and assigns the desired permissions).

[Here](https://github.com/JieYuan23/TeamsMeetingRecordingsNoDownload) you can find the same solution implemented as a single PowerShell script you might want to use for one-shot executions.

## Solution components
The solution is mainly composed by the follwing two components:
- an Azure AD App Registration
- an Azure Function App (built on PowerShell)

### Azure AD App Registration
The Azure AD App Registration is required to allow the two Azure functions to get authenticated and authorized. The Access Token is retrieved by specifying *client id* and *certificate thumbprint*.
The required (application) permissions to assign to the app are:
- Group.Read.All
- Sites.FullControl.All

### Azure Function App
The Azure Function App contains two Azure functions:
- [AddTeamsInQueue](/AddTeamsInQueue)
- [ProcessTeam](/ProcessTeam)

The PowerShell modules used by the solution are:
- PnP.PowerShell (loaded as  managed dependency)
- Microsoft.Graph.Authentication (loaded as  managed dependency)
- Microsoft.Graph.Teams (loaded as  managed dependency)
- Microsoft.Graph.Groups (loaded as  managed dependency)
- Microsoft.Graph.Files (explicitely inclueded in the solution due to issues when trying to load it as managed dependency)

The explicitely added Application Settings used by the Function App are:
- WEBSITE_RUN_FROM_PACKAGE (set to "1")
- CLIENT_ID (set as the Azure AD App Registration Id)
- CERT_THUMBPRINT (set as the certificate thumbripint)
- WEBSITE_LOAD_CERTIFICATES (set as the certificate thumbripint)
- CREATE_RECORDINGS_FOLDER (if "true" the solution creates the "Recordings" folders if not already there, otherwise it changes the permissions only on the already created "Recordings" folders. Set by default to "true")
- TENANT_PREFIX (set as the tenant prefix - the part of the tenant name just before ".onmicrosoft.com")
- SCHEDULE (defines the schedule of the **AddTeamsInQueue** function as [NCRONTAB expression](https://docs.microsoft.com/en-us/azure/azure-functions/functions-bindings-timer?tabs=csharp#ncrontab-expressions). Set by default at "0 0 6 * * *", that means each day at 6:00 AM UTC)

#### AddTeamsInQueue
This is a scheduled function (time triggered) that lists all the teams in the tenant and, for each of them, adds a message into an Azure Queue called **teamsqueue**.
Each message contains the team id and the team display name separated by a comma (eg.: *332cfb44-c4b5-4513-8404-72f3ed82e6d1,HR*).

If you want to manually add a message to the queue (e.g.: *332cfb44-c4b5-4513-8404-72f3ed82e6d1,HR*) in order to start the processing of a specific team, you can use the handy tool [Azure Storage Explorer](https://azure.microsoft.com/en-us/features/storage-explorer/).

The already defined schedule is each day at 6:00 AM UTC, you can change it by modifying the value of SCHEDULE application setting.

#### ProcessTeam
This is a queue triggered function (it triggers when new messages get into the **teamsqueue**) that processes the specific team.
The team processing entails:
- retrieving all the team channels (both standard and private)
- retrieving the channels folders
- creating the "Recordings" folder inside the channels folders (if not already created and CREATE_RECORDINGS_FOLDER application setting set to "true")
- changing the permissions on the **Recordings** folders so that team members won't be able to download files stored into those folders

This function logs processing outcome into an Azure storage table called **Log**. It logs at both teams and channels level.
The **Log** storage table contains the following columns:
- PartitionKey (set as the current year - eg.: *2021*)
- RowKey (set as team id or channels id + the current date and time)
- Timestamp
- TeamId
- TeamDisplayName
- ChannelId (null means the entity refers to the entire team, otherwise to the specific channel)
- ChannelDisplayName (null means the entity refers to the entire team, otherwise to the specific channel)
- StartTime
- EndTime
- Duration
- Outcome (tells if the team or the channel has been processed successfully or not)
- Details (tells in which step the error occurred)
- Exceptions (dumps the exceptions)
- CorrelationId (to correlate entities belonging to the same function execution instance)

All the dates and times are logged in UTC.

You can use the handy tool [Azure Storage Explorer](https://azure.microsoft.com/en-us/features/storage-explorer/) to see the entities that have been already added to the table.

## How to deploy
Deploying the solution on your tenant comprises 3 main steps:
1. Registering an App in Azure AD
2. Creating or retrieving the required Azure resources (Resource Group, Storage Account, Function App)
3. Deploy the zip package to the Function App (if you want to create the zip file by downloading this repository, keep in mind the zip file shouldn't contain a root folder but directly the contents; once downloaded assure you extract and re-zip the contents properly. [Zip deployment for Azure Functions](https://docs.microsoft.com/en-us/azure/azure-functions/deployment-zip-push)).

You can complete the major part of those steps programmatically by using the sample script below (it requires [PnP.PowerShell](https://pnp.github.io/powershell/) and [AZ](https://docs.microsoft.com/en-us/powershell/azure/new-azureps-module-az) PowerShell modules):
```powershell
#region VARIABLES
$tenantPrefix = "<tenantPrefix>"             # the part just before .onmicrosoft.com, e.g.: contoso
$appRegistrationName = "<appName>"           # the name of the Azure AD app registration
$certsOutputPath = "<folderFullPath>"        # the folder shoud be already existing, e.g.: c:\cert
$resourceGroupName = "<resourceGroupName>"   # the name of the Resource Group in which the resources will be created, if it doesn't match an existing Resource Group, a new one will be created with this name
$storageAccountName = "<storageAccountName>" # the name of the Storage Account in which the queue and the table will be created, if it doesn't match an existing Storage Account, it will be crated
$location = "West Europe"                    # the geographical location used for creating the resources
$functionAppName = "<functionAppName>"       # the name of the Function App
$createRecordingsFolder = "true"             # set this to "true" to have the script pre-create the Recordings folders if not already there
$zipPackage = "<zipPackageFullPath>"         # the full path to the zip file, e.g.: c:\package\file.zip
$subscriptionName = ""                       # leave blank if you have juts one subscription, otherwise specify which subscription you want to use
#endregion VARIABLES

#region AZURE APP REGISTRATION
Write-Host "Registering app '$appRegistrationName' in Azure AD"
$certPassword = Read-Host -Prompt "Enter certificate password" -AsSecureString
$appRegistration = Register-PnPAzureADApp -ApplicationName $appRegistrationName -Tenant "$tenantPrefix.onmicrosoft.com" -Store CurrentUser `
    -Scopes "MSGraph.Group.Read.All", "SPO.Sites.FullControl.All" `
    -DeviceLogin -OutPath $certsOutputPath -CertificatePassword $certPassword
$clientId = $appRegistration.'AzureAppId/ClientId'
$certThumbprint = $appRegistration.'Certificate Thumbprint'
Write-Host "Remember to grant admin consent for those permissions" -ForegroundColor Yellow
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
        SCHEDULE = "0 0 6 * * *"
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
The remaining manual steps are (as highlighted by the script as well):
- grant admin consent for all the permissions assigned to the Azure AD App Registration
- upload the private key certificate (pfx) in the Function App:
    - locate the certificate generated during the app registration (*$certsOutputPath* parameter)
    - navigate to the Function App
    - select **TLS/SSL settings**
    - select **Private Key Certificates (.pfx)** tab
    - click on **+ Upload Certificate**
    - select the certificate, enter the password (chosen during app registration) and click **Upload**
