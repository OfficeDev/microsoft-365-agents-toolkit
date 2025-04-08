# Welcome to Microsoft 365 Agents Toolkit!

## Quick Start

1. In the debug dropdown menu, select Dev Tunnels > Create A Tunnel (set authentication type to Public) or select an existing public dev tunnel
2. Right-click the 'TeamsApp' project in Solution Explorer and select **Microsoft 365 Agents Toolkit > Select Microsoft 365 Account**
3. Sign in to Microsoft 365 Agents Toolkit with a **Microsoft 365 work or school account**
4. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
5. In the opened web browser, select Add button to install the app in Teams
6. Open Windows PowerShell and post a HTTP request to trigger the notification:

   Invoke-WebRequest -Uri "http://localhost:5130/api/notification" -Method Post

> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).


## Learn more

New to Teams app development or Microsoft 365 Agents Toolkit? Learn more about 
Teams app manifests, deploying to the cloud, and more in the documentation 
at https://aka.ms/teams-toolkit-vs-docs

Learn more advanced topic like how to customize your notification bot code in 
tutorials at https://aka.ms/notification-bot-tutorial

## Report an issue

Select Visual Studio > Help > Send Feedback > Report a Problem. 
Or, you can create an issue directly in our GitHub repository: 
https://github.com/OfficeDev/TeamsFx/issues
