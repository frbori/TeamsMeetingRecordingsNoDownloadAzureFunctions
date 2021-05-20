# AddTeamsInQueue
This is a scheduled function (time triggered) that lists all the teams in the tenant and, for each of them, adds a message into an Azure Queue called **teamsqueue**.
It's possible to restrict the set of teams the function will add to the Azure Queue **teamsqueue** by specifing the following two application settings:
- TEAMS_CREATION_DATE_START (dd/MM/yyyy)
- TEAMS_CREATION_DATE_END (dd/MM/yyyy)
If none of the two settings is defined or set, all the Teams in the tenant will be added to **teamsqueue**.
If only the TEAMS_CREATION_DATE_START is set, the Teams added to the TeamsQueue will be the ones with *TeamsCreationDate >= TEAMS_CREATION_DATE_START*.
If only the TEAMS_CREATION_DATE_END is set, the Teams added to the TeamsQueue will be the ones with *TeamsCreationDate <= TEAMS_CREATION_DATE_END*.
If both are set, the Teams added to the TeamsQueue will be the ones with *TEAMS_CREATION_DATE_START <= TeamsCreationDate <= TEAMS_CREATION_DATE_END*.

Each message added to the Azure Queue contains the team id and the team display name separated by a comma (eg.: *332cfb44-c4b5-4513-8404-72f3ed82e6d1,HR*).

If you want to manually add a message to the queue (e.g.: *332cfb44-c4b5-4513-8404-72f3ed82e6d1,HR*) in order to start the processing of a specific team, you can use the handy tool [Azure Storage Explorer](https://azure.microsoft.com/en-us/features/storage-explorer/).

The already defined schedule is each day at 6:00 AM UTC, you can change it by modifying the value of SCHEDULE application setting.