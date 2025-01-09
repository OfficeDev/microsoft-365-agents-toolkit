# Welcome to Teams Toolkit!

## Quick Start

> **Prerequisites**
>
> To run this app template in your local dev machine, you will need:
>
> - [Visual Studio 2022](https://aka.ms/vs) 17.11 or higher and [install Teams Toolkit](https://aka.ms/install-teams-toolkit-vs)
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts).
> - [Microsoft 365 Copilot license](https://learn.microsoft.com/microsoft-365-copilot/extensibility/prerequisites#prerequisites)

1. In the debug dropdown menu, select `Dev Tunnels > Create a Tunnel` (set authentication type to Public) or select an existing public dev tunnel.
2. Right-click your project and select `Teams Toolkit > Prepare Teams App Dependencies`.
3. If prompted, sign in with a Microsoft 365 account for the Teams organization you want to install the app to.
4. Press F5, or select the `Debug > Start Debugging` menu in Visual Studio
5. When Teams launches in the browser, click the Apps icon from Teams client left rail to open Teams app store and search for Copilot.
6. Open the `Copilot` app, select `Plugins`, and from the list of plugins, turn on the toggle for your plugin. Now, you can send a prompt to trigger your plugin.
7. Send a message to Copilot to query the repair record. For example: List all repairs.
   > Note: Please make sure to switch to New Teams when Teams web client has launched

## Learn more

{{^DeclarativeCopilot}}
- [Extend Microsoft 365 Copilot](https://aka.ms/teamsfx-copilot-plugin)
{{/DeclarativeCopilot}}
{{#DeclarativeCopilot}}
- [Declarative agents for Microsoft 365](https://aka.ms/teams-toolkit-declarative-agent)
{{/DeclarativeCopilot}}

## Report an issue

Select Visual Studio > Help > Send Feedback > Report a Problem.
Or, you can create an issue directly in our GitHub repository:
https://github.com/OfficeDev/TeamsFx/issues
