# Welcome to Teams Toolkit!

## Quick Start

{{#enableTestToolByDefault}}
1. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
2. In Teams App Test Tool from the launched browser, type and send "helloWorld" to your app to trigger a response
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
1. In the debug dropdown menu, select Dev Tunnels > Create A Tunnel (set authentication type to Public) or select an existing public dev tunnel
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/create-devtunnel-button.png)
1. Right-click the 'TeamsApp' project in Solution Explorer and select **Teams Toolkit > Select Microsoft 365 Account**
2. Sign in to Teams Toolkit with a **Microsoft 365 work or school account**
3. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
4. In the opened web browser, select Add button to install the app in Teams
6. In the message input field, type and send "helloWorld" to your app to get a response
{{/enableTestToolByDefault}}

{{^enableTestToolByDefault}}
## Debug in Test Tool
Teams App Test Tool allows developers test and debug bots locally without needing Microsoft 365 accounts, development tunnels, or Teams app and bot registration. See https://aka.ms/teams-toolkit-vs-test-tool for more details.
{{/enableTestToolByDefault}}

## Run the app on other platforms

The Teams app can run in other platforms like Outlook and Microsoft 365 app. See https://aka.ms/vs-ttk-debug-multi-profiles for more details.

## Get more info

New to Teams app development or Teams Toolkit? Explore Teams app manifests, cloud deployment, and much more in the https://aka.ms/teams-toolkit-vs-docs.

Learn more advanced topic like how to customize your command bot code in 
tutorials at https://aka.ms/command-bot-tutorial

## Report an issue

Select Visual Studio > Help > Send Feedback > Report a Problem. 
Or, create an issue directly in our GitHub repository:
https://github.com/OfficeDev/TeamsFx/issues
