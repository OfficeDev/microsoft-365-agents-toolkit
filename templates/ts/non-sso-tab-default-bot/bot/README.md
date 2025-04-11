# How to use this Bots and Message Extensions HelloWorld app

A bot, chatbot, or conversational bot is an app that responds to simple commands sent in chat and replies in meaningful ways. Examples of bots in everyday use include: bots that notify about build failures, bots that provide information about the weather or bus schedules, or provide travel information. A bot interaction can be a quick question and answer, or it can be a complex conversation. Being a cloud application, a bot can provide valuable and secure access to cloud services and corporate resources.

A Message Extension allows users to interact with your web service while composing messages in the Microsoft Teams client. Users can invoke your web service to assist message composition, from the message compose box, or from the search bar.

Message Extensions are implemented on top of the Bot support architecture within Teams.

This is a simple hello world application with both Bot and Message extension capabilities.

## Prerequisites

- [Node.js](https://nodejs.org/), supported versions: 18, 20, 22
- An M365 account. If you do not have M365 account, apply one from [M365 developer program](https://developer.microsoft.com/en-us/microsoft-365/dev-program)
- [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)

## Debug

- From Visual Studio Code: Start debugging the project by hitting the `F5` key in Visual Studio Code.
- Alternatively use the `Run and Debug Activity Panel` in Visual Studio Code and click the `Run and Debug` green arrow button.
- From Microsoft 365 Agents Toolkit CLI:
  - Install [dev tunnel cli](https://aka.ms/teamsfx-install-dev-tunnel).
  - Login with your M365 Account using the command `devtunnel user login`.
  - Start your local tunnel service by running the command `devtunnel host -p 3978 --protocol http --allow-anonymous`.
  - In the `env/.env.local` file, fill in the values for `BOT_DOMAIN` and `BOT_ENDPOINT` with your dev tunnel URL.
    ```
    BOT_DOMAIN=sample-id-3978.devtunnels.ms
    BOT_ENDPOINT=https://sample-id-3978.devtunnels.ms
    ```
  - Executing the command `teamsapp provision --env local` in your project directory.
  - Executing the command `teamsapp deploy --env local` in your project directory.
  - Executing the command `teamsapp preview --env local` in your project directory.

## Edit the manifest

You can find the app manifest in `../appPackage` folder. The folder contains one manifest file:

- `manifest.json`: Manifest file for Teams app running locally or running remotely (After deployed to Azure).

This file contains template arguments with `${{...}}` statements which will be replaced at build time. You may add any extra properties or permissions you require to this file. See the [schema reference](https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema) for more information.

## Deploy to Azure

Deploy your project to Azure by following these steps:

| From Visual Studio Code                                                                                                                                                                                                                                                                                                                                                                                                                                             | From Microsoft 365 Agents Toolkit CLI                                                                                                                                               |
| :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | :------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| <ul><li>Open Microsoft 365 Agents Toolkit, and sign into Azure by clicking the `Sign in to Azure` under the `ACCOUNTS` section from sidebar.</li> <li>After you signed in, select a subscription under your account.</li><li>Open the Microsoft 365 Agents Toolkit and click `Provision` from DEPLOYMENT section or open the command palette and select: `Teams: Provision`.</li><li>Open the Microsoft 365 Agents Toolkit and click `Deploy` or open the command palette and select: `Teams: Deploy`.</li></ul> | <ul> <li>Run command `teamsapp auth login azure`.</li> <li>Run command `teamsapp provision --env dev`.</li> <li>Run command: `teamsapp deploy --env dev`. </li></ul> |

> Note: Provisioning and deployment may incur charges to your Azure Subscription.

## Preview

Once the provisioning and deployment steps are finished, you can preview your app:

- From Visual Studio Code

  1. Open the `Run and Debug Activity Panel`.
  1. Select `Launch Remote (Edge)` or `Launch Remote (Chrome)` from the launch configuration drop-down.
  1. Press the Play (green arrow) button to launch your app - now running remotely from Azure.

- From Microsoft 365 Agents Toolkit CLI: execute `teamsapp preview --remote` in your project directory to launch your application.

## Validate manifest file

To check that your manifest file is valid:

- From Visual Studio Code: open the command palette and select: `Teams: Validate Application`.
- From Microsoft 365 Agents Toolkit CLI: run command `teamsapp validate` in your project directory.

## Package

- From Visual Studio Code: open the Microsoft 365 Agents Toolkit and click `Zip app package` or open the command palette and select `Teams: Zip app package`.
- Alternatively, from the command line run `teamsapp package` in the project directory.

## Publish to Teams

Once deployed, you may want to distribute your application to your organization's internal app store in Teams. Your app will be submitted for admin approval.

- From Visual Studio Code: open the Microsoft 365 Agents Toolkit and click `Publish` or open the command palette and select: `Teams: Publish`.
- From Microsoft 365 Agents Toolkit CLI: run command `teamsapp publish` in your project directory.

## Play with Messging Extension

This template provides some sample functionality:

- You can search for `npm` packages from the search bar.

- You can create and send an adaptive card.

  ![CreateCard](https://github.com/OfficeDev/TeamsFx/assets/86260893/a0a8304b-3074-4eb8-9097-655cdda0b937)

- You can share a message in an adaptive card form.

  ![ShareMessage](https://github.com/OfficeDev/TeamsFx/assets/86260893/a7d4dd7b-6466-4e89-8f42-b93629a90bc8)

- You can paste a link that "unfurls" (`.botframework.com` is monitored in this template) and a card will be rendered.

  ![ComposeArea](https://github.com/OfficeDev/TeamsFx/assets/86260893/2b155dc8-9c01-4f14-8e2f-d179b81e97c6)

To trigger these functions, there are multiple entry points:

- Type a `/` in the command box and select your message extension.

  ![AtBotFromSearch](https://github.com/OfficeDev/TeamsFx/assets/86260893/d9ee7f72-0248-4a35-ae4d-e09d447614e6)

- Click the `...` under compose message area, find your message extension.

  ![ComposeArea](https://github.com/OfficeDev/TeamsFx/assets/86260893/f447f015-bb68-4ae2-9e0a-aae69c00c328)

- Click the `...` next to any messages you received or sent.

  ![ComposeArea](https://github.com/OfficeDev/TeamsFx/assets/86260893/0237dc5a-8b4d-4f52-a2fb-95ad17264c90)

## Further reading

### Bot

- [Bot Basics](https://docs.microsoft.com/azure/bot-service/bot-builder-basics?view=azure-bot-service-4.0)
- [Bot Framework Documentation](https://docs.botframework.com/)
- [Azure Bot Service Introduction](https://docs.microsoft.com/azure/bot-service/bot-service-overview-introduction?view=azure-bot-service-4.0)

### Message Extension

- [Search Command](https://docs.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/how-to/search-commands/define-search-command)
- [Action Command](https://docs.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/how-to/action-commands/define-action-command)
- [Link Unfurling](https://docs.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/how-to/link-unfurling?tabs=dotnet)
