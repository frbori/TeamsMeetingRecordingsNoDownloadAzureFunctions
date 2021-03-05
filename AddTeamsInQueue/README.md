# AddTeamsInQueue
This is a scheduled function (time triggered) that lists all the teams in the tenant and, for each of them, adds a message into an Azure Queue called **teamsqueue**.
Each message contains the team id and the team display name separated by a comma (eg.: *332cfb44-c4b5-4513-8404-72f3ed82e6d1,HR*).

The already defined schedule is each day at 6:00 AM UTC (file [function.json](./function.json)).
