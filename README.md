# TeamsMeetingRecordingsNoDownload
This solution aims to prevent that Teams channel meeting recordings stored into SharePoint can be downloaded by team members.
By default, Teams channel meeting recordings are saved into the SharePoint site associated to the team and team members are added to the default SharePoint members group, this gives them Edit permission on all the SharePoint contents, including the possibility of downloading files.
This solution basically changes the permissions assigned to the default SharePoint members group on the folders containing the recordings files (it breaks the permission inheritance and assigns the desired permission).
