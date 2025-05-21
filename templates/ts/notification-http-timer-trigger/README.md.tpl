# Overview of the Notification bot template

This template showcases an app that send a message to Teams with Adaptive Cards triggered by a HTTP post request or timer schedule. You can further extend the template to consume, transform and post events to individual, chat or channel in Teams.

The app template is built using the TeamsFx SDK, which provides a simple set of functions over the Microsoft 365 Agents SDK to implement this scenario.

## Get Started with the Notification bot

>
> **Prerequisites**
>
> To run the notification bot template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 18, 20, 22
> - An [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
> - [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
>
> **Note**
>
> Your app can be installed into a team, or a group chat, or as personal app. See [Installation and Uninstallation](https://aka.ms/teamsfx-notification-new#customize-installation).
> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

1. First, select the Microsoft 365 Agents Toolkit icon on the left in the VS Code toolbar.
2. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.
3. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)`.
4. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.
5. If you select `Timer Trigger`, wait for 30 seconds. If you select `HTTP Trigger`, send a POST request to `http://<endpoint>/api/notification` with your favorite tool (like `Postman`)

   - When your project is running locally, replace `<endpoint>` with `localhost:3978`
   - When your project is deployed to Azure App Service, replace `<endpoint>` with the url from Azure App Service

The bot will send an Adaptive Card to Teams:

![Notification Message in Teams](https://user-images.githubusercontent.com/7642967/223006044-5003574e-2aee-4a41-9b71-c103d0439012.png)

## What's included in the template

| Folder / File | Contents |
| - | - |
| `m365agents.yml` | Main project file describes your application configuration and defines the set of actions to run in each lifecycle stages |
| `m365agents.local.yml`| This overrides `m365agents.yml` with actions that enable local execution and debugging |
| `.vscode/` | VSCode files for local debug |
| `src/` | The source code for the notification application |
| `appPackage/` | Templates for the application manifest |
| `infra/` | Templates for provisioning Azure resources |

The following files can be customized and demonstrate an example implementation to get you started.

| File | Contents |
| - | - |
| `*Trigger/function.json` | Azure Function bindings for the notification trigger |
| `src/*Trigger.ts` | Notification trigger implementation |
| `src/teamsBot.ts`| An empty teams activity handler for bot customization |
| `src/adaptiveCards/notification-default.json` | A generated Adaptive Card that is sent to Teams |
| `src/cardModels.ts` | The default Adaptive Card data model |

The following files implement the core notification on the Microsoft 365 Agents SDK. You generally will not need to customize these files.

| File / Folder | Contents |
| - | - |
| `src/internal/initialize.ts` | Application initialization |
| `messageHandler/` | Azure Function bindings to implement Bot protocol |
| `src/internal/messageHandler.ts`<br/>`src/internal/responseWrapper.ts` | Bot protocol implementation |

The following files are project-related files. You generally will not need to customize these files.

| File / Folder | Contents |
| - | - |
| `.funcignore` | Azure Functions ignore file to exclude local files |
| `.gitignore` | Git ignore file |
| `host.json` | Azure Functions host file |
| `local.settings.json` | Azure Functions settings for local debugging |
| `package.json` | NPM package file |

## Extend the notification bot template

There are few customizations you can make to extend the template to fit your business requirements.

1. [Step 1: Customize the trigger point from event source](#step-1-customize-the-trigger-point-from-event-source)
2. [Step 2: Customize the notification content](#step-2-customize-the-notification-content)
3. [Step 3: Customize where notifications are sent](#step-3-customize-where-notifications-are-sent)

### Step 1: Customize the trigger point from event source

If you selected `timer` trigger, the default Azure Function timer trigger (`src/timerTrigger.ts`) implementation simply sends a hard-coded Adaptive Card every 30 seconds. You can edit the file `*Trigger/function.json` to customize the `schedule` property. Refer to the [Azure Function documentation](https://docs.microsoft.com/azure/azure-functions/functions-bindings-timer?tabs=in-process&pivots=programming-language-javascript#ncrontab-expressions) for more details.

If you selected `http` trigger, when this trigger is hit (via a HTTP request), the default implementation sends a hard-coded Adaptive Card to Teams. You can change this behavior by customizing `src/*Trigger.ts`. A typical implementation might make an API call to retrieve some events and/or data, and then send an Adaptive Card as appropriate.

You can also add any Azure Function trigger. For example:

- You can use an `Event Hub` trigger to send notifications when an event is pushed to Azure Event Hub.
- You can use a `Cosmos DB` trigger to send notifications when a Cosmos document has been created or updated.

See Azure Functions [supported triggers](https://docs.microsoft.com/azure/azure-functions/functions-triggers-bindings?tabs=javascript#supported-bindings).

## Step 2: Customize the notification content

`src/adaptiveCards/notification-default.json` defines the default Adaptive Card. You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to help visually design your Adaptive Card UI.

`src/cardModels.ts` defines a data structure that is used to fill data for the Adaptive Card. The binding between the model and the Adaptive Card is done by name matching (for example,`CardData.title` maps to `${title}` in the Adaptive Card). You can add, edit, or remove properties and their bindings to customize the Adaptive Card to your needs.

You can also add new cards if needed. Follow this [sample](https://aka.ms/teamsfx-adaptive-card-sample-new) to see how to build different types of adaptive cards with a list or a table of dynamic contents using `ColumnSet` and `FactSet`.

### Step 3: Customize where notifications are sent

Notifications can be sent to where the bot is installed:

- [Send notifications to a channel](https://aka.ms/teamsfx-notification-new#send-notifications-to-a-channel)
- [Send notifications to a group chat](https://aka.ms/teamsfx-notification-new#send-notifications-to-a-group-chat)
- [Send notifications to a personal chat](https://aka.ms/teamsfx-notification-new#send-notifications-to-a-personal-chat)

You can also send the notifications to a specific receiver:

- [Send notifications to a specific channel](https://aka.ms/teamsfx-notification-new#send-notifications-to-a-specific-channel)
- [Send notifications to a specific person](https://aka.ms/teamsfx-notification-new#send-notifications-to-a-specific-person)

Congratulations, you've just created your own notification! To learn more about extending the notification bot template, [visit the documentation on Github](https://aka.ms/teamsfx-notification-new). You can find more scenarios like:

- [Customize storage](https://aka.ms/teamsfx-notification-new#customize-storage)
- [Customize adapter](https://aka.ms/teamsfx-notification-new#customize-adapter)
- [Customize the way to initialize the bot](https://aka.ms/teamsfx-notification-new#customize-initialization)
- [Add authentication for your notification API](https://aka.ms/teamsfx-notification-new#add-authentication-for-your-notification-api)
- [Connect to existing APIs](https://aka.ms/teamsfx-notification-new#connect-to-existing-api)
- [Frequently asked questions](https://aka.ms/teamsfx-notification-new#frequently-asked-questions)

## Extend notification bot with other bot scenarios

Notification bot is compatible with other bot scenarios like command bot and workflow bot.

### Add command to your application

The command and response feature adds the ability for your application to "listen" to commands sent to it via a Teams message and respond to commands with Adaptive Cards. Follow the [steps here](https://aka.ms/teamsfx-command-new#How-to-add-more-command-and-response) to add the command response feature to your workflow bot. Refer [the command bot document](https://aka.ms/teamsfx-command-new) for more information.

### Add workflow to your notification bot

Adaptive cards can be updated on user action to allow user progress through a series of cards that require user input. Developers can define actions and use a bot to return an Adaptive Cards in response to user action. This can be chained into sequential workflows. Follow the [steps here](https://aka.ms/teamsfx-workflow-new#add-more-card-actions) to add workflow feature to your command bot. Refer [the workflow document](https://aka.ms/teamsfx-workflow-new) for more information.

## Additional information and references

- [Manage multiple environments](https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-multi-env)
- [Collaborate with others](https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-collaboration)
- [Microsoft 365 Agents Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [TeamsFx SDK](https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-sdk)
- [Microsoft 365 Agents Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)
