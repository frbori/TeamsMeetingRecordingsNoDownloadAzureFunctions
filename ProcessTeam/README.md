# ProcessTeam
This is a queue triggered function (it triggers when new messages get into the **teamsqueue**) that processes the specific team.
The team processing entails:
- retrieving all the team channels (both standard and private)
- retrieving the channels folders
- creating the "Recordings" folder inside the channels folders (if not already created)
- changing the permissions on the **Recordings** folders so that team members won't be able to download files stored into those folders