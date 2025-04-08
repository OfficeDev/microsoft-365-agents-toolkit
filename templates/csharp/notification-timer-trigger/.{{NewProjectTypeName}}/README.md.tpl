# Welcome to Teams Toolkit!

## Quick Start

1. In the debug dropdown menu, select Dev Tunnels > Create A Tunnel (set authentication type to Public) or select an existing public dev tunnel
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/create-devtunnel-button.png)
2. Right-click the '{{NewProjectTypeName}}' project in Solution Explorer and select **Teams Toolkit > Select Microsoft 365 Account**
3. Sign in to Teams Toolkit with a **Microsoft 365 work or school account**
4. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
5. In the opened web browser, select Add button to install the app in Teams
6. [If you selected http trigger] Open Windows PowerShell and post a HTTP request to trigger 
the notification(replace \<endpoint\> with real endpoint, for example localhost:5130):

   Invoke-WebRequest -Uri "http://\<endpoint\>/api/notification" -Method Post

## Run the app on other platforms

The Teams app can run in other platforms like Outlook and Microsoft 365 app. See https://aka.ms/vs-ttk-debug-multi-profiles for more details.

## Get more info

New to Teams app development or Teams Toolkit? Explore Teams app manifests, cloud deployment, and much more in the https://aka.ms/teams-toolkit-vs-docs.

Get more info on advanced topic like how to customize your notification bot code in tutorials at https://aka.ms/notification-bot-tutorial

## Report an issue

Select Visual Studio > Help > Send Feedback > Report a Problem. 
Or, create an issue directly in our GitHub repository:
https://github.com/OfficeDev/TeamsFx/issues
