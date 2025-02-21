# Welcome to Teams Toolkit!

## Quick Start

> **Prerequisites**
>
> To run this app template in your local dev machine, you will need:
>
> - [Visual Studio 2022](https://aka.ms/vs) 17.11 or higher and [install Teams Toolkit](https://aka.ms/install-teams-toolkit-vs)
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)

### Add your own API Key

1. Open PowerShell, change the current working directory to this project root and run command `./TeamsApp/GenerateApiKey.ps1`
    ```
    > ./TeamsApp/GenerateApiKey.ps1
    ```

2. The above command will output something like "Generated a new API Key: xxx...".
3. Fill in API Key into `env/.env.*.user`.
    ```
    SECRET_API_KEY=<your-api-key>
    ```

### Start the app in Teams Web Client

1. If you haven't added your own API Key, please follow the above steps to add your own API Key.
2. In the debug dropdown menu, select Dev Tunnels > Create a Tunnel (set authentication type to Public) or select an existing public dev tunnel
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/create-devtunnel-button.png)
3. Right-click your project and select Teams Toolkit > Prepare Teams App Dependencies
4. If prompted, sign in with a Microsoft 365 account for the Teams organization you want to install the app to.
5. Press F5, or select the `Debug > Start Debugging` menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
6. When Teams launches in the browser, click the Apps icon from Teams client left rail to open Teams app store and search for Copilot.
7. Open the `Copilot` app, select `Plugins`, and from the list of plugins, turn on the toggle for your plugin. Now, you can send a prompt to trigger your plugin.
8. Send a message to Copilot to query the repair record. For example: List all repairs.
   > Note: Please make sure to switch to New Teams when Teams web client has launched

## Get more info

- [Declarative agents for Microsoft 365](https://aka.ms/teams-toolkit-declarative-agent)

## Report an issue

Select Visual Studio > Help > Send Feedback > Report a Problem.
Or, create an issue directly in our GitHub repository:
https://github.com/OfficeDev/TeamsFx/issues
