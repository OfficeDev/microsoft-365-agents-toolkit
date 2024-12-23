# Overview of the Command bot template

This template showcases an app that responds to chat commands by displaying UI using an Adaptive Card. This enables your users to type in simple messages in Teams and your application can provide an appropriate response based on the contents of the message.

The app template is built using the Teams AI SDK, which provides a simple set of functions over the Microsoft Bot Framework to implement this scenario.

## Get Started with the Command bot

> **Prerequisites**
>
> To run the command bot template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 18, 20
{{^enableTestToolByDefault}}
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
{{/enableTestToolByDefault}}
> - [Teams Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Teams Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
>
> **Note**
>
> Your app can be installed into a team, or a group chat, or as personal app. See [Installation and Uninstallation](https://aka.ms/teamsfx-command-new#customize-installation).
> For local debugging using Teams Toolkit CLI, you need to do some extra steps described in [Set up your Teams Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

1. First, select the Teams Toolkit icon on the left in the VS Code toolbar.
{{#enableTestToolByDefault}}
2. Press F5 to start debugging which launches your app in Teams App Test Tool using a web browser. Select `Debug in Test Tool`.
3. The browser will pop up to open Teams App Test Tool.
4. Type or select `helloWorld` in the chat to send it to your bot - this is the default command provided by the template.
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
2. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.
3. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)`.
4. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.
5. Type or select `helloWorld` in the chat to send it to your bot - this is the default command provided by the template.
{{/enableTestToolByDefault}}

The bot will respond to the `helloWorld` command with an Adaptive Card:

{{#enableTestToolByDefault}}
![Command and Response in Test Tool](https://github.com/OfficeDev/TeamsFx/assets/9698542/2636fd91-ec7f-4740-a5ea-b272575b0b7c)
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
![Command and Response in Teams](https://user-images.githubusercontent.com/11220663/165891754-16916b68-c1b5-499d-b6a8-bdfb195f1fd0.png)
{{/enableTestToolByDefault}}

## What's included in the template

| Folder / File | Contents |
| - | - |
| `teamsapp.yml` | Main project file describes your application configuration and defines the set of actions to run in each lifecycle stages |
| `teamsapp.local.yml`| This overrides `teamsapp.yml` with actions that enable local execution and debugging |
| `teamsapp.testtool.yml`| This overrides `teamsapp.yml` with actions that enable local execution and debugging in Teams App Test Tool |
| `env/`| Name / value pairs are stored in environment files and used by `teamsapp.yml` to customize the provisioning and deployment rules |
| `.vscode/` | VSCode files for debugging |
| `appPackage/` | Templates for the Teams application manifest |
| `infra/` | Templates for provisioning Azure resources |
| `src/` | The source code for the application |

The following files can be customized and demonstrate an example implementation to get you started.

| File | Contents |
| - | - |
| `src/index.ts` | Application entry point and `express` handlers for command and response |
| `src/teamsBot.ts`  | An empty teams activity handler for bot customization |
| `src/adaptiveCards/helloworldCommand.json` | A generated Adaptive Card that is sent to Teams |
| `src/helloworldCommandHandler.ts` | The business logic to handle a command |

## Extend the command bot template with more commands and responses

Follow the steps below to add more commands and responses to extend the command bot:

1. [Step 1: Add a command definition in manifest](#step-1-add-a-command-definition-in-manifest)
2. [Step 2: Respond with an Adaptive Card](#step-2-respond-with-an-adaptive-card)
3. [Step 3: Handle the command](#step-3-handle-the-command)
4. [Step 4: Register the new command](#step-4-register-the-new-command)

### Step 1: Add a command definition in manifest

You can edit the manifest template file `appPackage\manifest.json` to include definitions of a `doSomething` command with its title and description in the `commands` array:

```json
"commandLists": [
  {
    "commands": [
        {
            "title": "helloWorld",
            "description": "A helloworld command to send a welcome message"
        },
        {
            "title": "doSomething",
            "description": "A sample do something command"
        }
    ]
  }
]
```

### Step 2: Respond with an Adaptive Card

To respond with an Adaptive Card, define your card in its JSON format. Create a new file `src/adaptiveCards/doSomethingCommandResponse.json`:

```json
{
  "type": "AdaptiveCard",
  "body": [
    {
      "type": "TextBlock",
      "size": "Medium",
      "weight": "Bolder",
      "text": "Your doSomething Command is added!"
    },
    {
      "type": "TextBlock",
      "text": "Congratulations! Your hello world bot now includes a new DoSomething Command",
      "wrap": true
    }
  ],
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "version": "1.4"
}
```

You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to help visually design your Adaptive Card UI.

> Please note:

> - Respond with an Adaptive Card is optional, you can simply respond with plain texts.
> - If you'd like to send adaptive card with dynamic data, please refer to [this document](https://aka.ms/teamsfx-command-new#how-to-build-command-response-using-adaptive-card-with-dynamic-content).

### Step 3: Handle the command

Create a new file, `src/doSomethingCommandHandler.ts`:

```typescript
import { Activity, CardFactory, MessageFactory, TurnContext } from "botbuilder";
import { Selector } from "@microsoft/teams-ai";
import doSomethingCard from "./adaptiveCards/doSomethingCommandResponse.json";
import * as ACData from "adaptivecards-templating";

export class DoSomethingCommandHandler implements TeamsFxBotCommandHandler {
  triggerPatterns: string | RegExp | Selector | (string | RegExp | Selector)[] = "doSomething";

  async handleCommandReceived(
    context: TurnContext,
    state: ApplicationTurnState
  ): Promise<string | Partial<Activity>> {
    // verify the command arguments which are received from the client if needed.
    console.log(`App received message: ${context.activity.text}`);

    const cardJson = new ACData.Template(helloWorldCard).expand({
      $root: {
        title: "doSomething command is added",
        body: "Congratulations! You have responded to doSomething command",
      },
    });
    return MessageFactory.attachment(CardFactory.adaptiveCard(cardJson));
  }
}
```

You can customize what the command does here, including calling an API, process data, etc.

### Step 4: Register the new command

Each new command needs to be configured in the `ConversationBot`, which powers the conversational flow of the command bot template. Navigate to the `src/index.ts` file and register the trigger pattern to `app.message()`:

```typescript
const doSomethingCommandHandler = new DoSomethingCommandHandler();
app.message(
  doSomethingCommandHandler.triggerPatterns,
  async (context: TurnContext, state: ApplicationTurnState) => {
    const reply = await doSomethingCommandHandler.handleCommandReceived(context, state);

    if (reply) {
      await context.sendActivity(reply);
    }
  }
);
```

Congratulations, you've just created your own command! To learn more about the command bot template, [visit the documentation on GitHub](https://aka.ms/teamsfx-command-new). You can find more scenarios like:

- [Customize the trigger pattern](https://aka.ms/teamsfx-command-new#customize-the-trigger-pattern)
- [Customize the Adaptive Card with dynamic content](https://aka.ms/teamsfx-command-new#how-to-build-command-response-using-adaptive-card-with-dynamic-content)
- [Change the way to initialize the bot](https://aka.ms/teamsfx-command-new#customize-initialization)
- [Connect to an existing API](https://aka.ms/teamsfx-command-new#connect-to-existing-api)
- [Access Microsoft Graph](https://aka.ms/teamsfx-add-sso-new)

## Extend command bot with other bot scenarios

Command bot is compatible with other bot scenarios like notification bot and workflow bot.

### Add notifications to your command bot

The notification feature adds the ability for your application to send Adaptive Cards in response to external events. Follow the [steps here](https://aka.ms/teamsfx-command-new#how-to-extend-my-command-and-response-bot-to-support-notification) to add the notification feature to your command bot. Refer [the notification document](https://aka.ms/teamsfx-notification-new) for more information.

### Add workflow to your command bot

Adaptive cards can be updated on user action to allow user progress through a series of cards that require user input. Developers can define actions and use a bot to return an Adaptive Cards in response to user action. This can be chained into sequential workflows. Follow the [steps here](https://aka.ms/teamsfx-workflow-new#add-more-card-actions) to add workflow feature to your command bot. Refer [the workflow document](https://aka.ms/teamsfx-workflow-new) for more information.

## Additional information and references

- [Manage multiple environments](https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-multi-env)
- [Collaborate with others](https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-collaboration)
- [Teams Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Teams Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Teams AI SDK](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/teams-conversational-ai/teams-conversation-ai-overview)
- [Teams Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)
