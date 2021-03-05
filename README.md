# TeamsMeetingRecordingsNoDownload
This solution aims to prevent that Teams channel meeting recordings stored into SharePoint can be downloaded by team members.

By default, Teams channel meeting recordings are saved into the SharePoint site associated to the team and team members are added to the default SharePoint members group, this gives them Edit permission on all the SharePoint contents, including the possibility of downloading files.
This solution basically changes the permissions assigned to the default SharePoint members group on the folders containing the recordings files (it breaks the permission inheritance and assigns the desired permission).

## Solution components
The solution is mainly composed by two Azure Functions (built on PowerShell):
- AddTeamsInQueue
- ProcessTeam

The PowerShell modules used by the solution are:
- PnP.PowerShell
- Microsoft.Graph.Authentication
- Microsoft.Graph.Teams
- Microsoft.Graph.Groups
- Microsoft.Graph.Files
### AddTeamsInQueue
This is a scheduled function (time triggered) that lists all the teams in the tenant and, for each of them, insert a message into an Azure Queue called *teamsqueue*.
Each message contains the team id and the team display name separated by a comma (eg.: 332cfb44-c4b5-4513-8404-72f3ed82e6d1,HR).
### ProcessTeam
This is a queue triggered function (it triggers when new messages get into the *teamsqueue*) that processes the specific team.
The team processing entails:
- retrieving all the team channels (both standard and private)
- retrieving the channels folders
- creating the "Recordings" folder inside the channels folders (if not already created)
- changing the permissions on the **Recordings** folders so that team members won't be able to download files stored into those folders
