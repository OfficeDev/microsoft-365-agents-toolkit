# Overview of the Workflow bot template

This template showcases an app that responds to chat commands by displaying UI using an Adaptive Card. The card has a button that demonstrates how to receive user input on the card, do something like call an API, and update the UI of that card. This can be further customized to create richer, more complex sequence of steps which forms a complete workflow.

The app template is built using the TeamsFx SDK, which provides a simple set of functions over the Microsoft Bot Framework to implement this scenario.

## Get Started with the Workflow bot

> **Prerequisites**
>
> To run the workflow bot template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 18, 20
{{^enableTestToolByDefault}}
> - A [Microsoft 365 account for development](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts)
{{/enableTestToolByDefault}}
> - [Teams Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Teams Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
>
> **Note**
>
> Your app can be installed into a team, or a group chat, or as personal app. See [Installation and Uninstallation](https://aka.ms/teamsfx-command-response#customize-installation).
> For local debugging using Teams Toolkit CLI, you need to do some extra steps described in [Set up your Teams Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

1. First, select the Teams Toolkit icon on the left in the VS Code toolbar.
{{#enableTestToolByDefault}}
2. Press F5 to start debugging which launches your app in Teams App Test Tool using a web browser. Select `Debug in Test Tool`.
3. The browser will pop up to open Teams App Test Tool.
4. Type or select `helloWorld` in the chat to send it to your bot - this is the default command provided by the template.
5. In the response from the bot, select the **DoStuff** button.
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
2. In the Account section, sign in with your [Microsoft 365 account](https://docs.microsoft.com/microsoftteams/platform/toolkit/accounts) if you haven't already.
3. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)`.
4. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.
5. Type or select `helloWorld` in the chat to send it to your bot - this is the default command provided by the template.
6. In the response from the bot, select the **DoStuff** button.
{{/enableTestToolByDefault}}

The bot will respond by updating the existing Adaptive Card to show the workflow is now complete! Continue reading to learn more about what's included in the template and how to customize it.

Here is a screen shot of the application running:

{{#enableTestToolByDefault}}
![Responds to command](https://github.com/OfficeDev/TeamsFx/assets/9698542/94173beb-2fff-44fa-87dc-f686677da631)

When you click the `DoStuff` button, the above adaptive card will be updated to a new card as shown below:

![Responds to card action](https://github.com/OfficeDev/TeamsFx/assets/9698542/521ff12d-726c-4087-825a-112547cad836)
{{/enableTestToolByDefault}}
{{^enableTestToolByDefault}}
![Responds to command](https://user-images.githubusercontent.com/10163840/192477792-dc447b3a-e304-4cd8-b4df-b1eb9d226292.png)

When you click the `DoStuff` button, the above adaptive card will be updated to a new card as shown below:

![Responds to card action](https://user-images.githubusercontent.com/10163840/192477148-29d9edfc-085b-4d02-b3de-b47b9a456108.png)
{{/enableTestToolByDefault}}

## What's included in the template

| Folder | Contents |
| - | - |
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
| `src/index.js`| Application entry point and `express` handlers for the Workflow bot |
| `src/teamsBot.js` | An empty teams activity handler for bot customization |
| `src/commands/helloworldCommandHandler.js` | Implementation that handles responding to a chat command |
| `src/adaptiveCards/helloworldCommandResponse.json` | Defines the Adaptive Card (UI) that is displayed in response to a chat command |
| `src/adaptiveCards/doStuffActionResponse.json` | A generated Adaptive Card that is sent to Teams for the response of "doStuff" action |
| `src/cardActions/doStuffActionHandler.js` | Implements the handler for the `doStuff` button displayed in the Adaptive Card |

## Extend the workflow bot template with more actions and responses

Follow steps below to add more actions and responses to extend the workflow bot:

1. [Step 1: Add an action to your Adaptive Card](#step-1-add-an-action-to-your-adaptive-card)
2. [Step 2: Respond with a new Adaptive Card](#step-2-respond-with-a-new-adaptive-card)
3. [Step 3: Handle the new action](#step-3-handle-the-new-action)
4. [Step 4: Register the new handler](#step-4-register-the-new-handler)

### Step 1: Add an action to your Adaptive Card

Adding new actions (buttons) to an Adaptive Card is as simple as defining them in the JSON file. Add a new `DoSomething` action to the `src/adaptiveCards/helloworldCommandResponse.json` file:

Here's a sample action with type `Action.Execute`:

```json
{
  "type": "AdaptiveCard",
  "body": [
    ...
    {
      "type": "ActionSet",
      "actions": [
        {
          "type": "Action.Execute",
          "title": "DoSomething",
          "verb": "doSomething"
        }
      ]
    },
    ...
  ]
}
```

Specifying the `type` as `Action.Execute` to define an universal action in the base card. User can click the button to perform some business task in Teams chat. Learn more about [Adaptive Card Universal Actions in the documentation](https://learn.microsoft.com/microsoftteams/platform/task-modules-and-cards/cards/universal-actions-for-adaptive-cards/overview?tabs=mobile#universal-actions).

> **_NOTE:_** the `verb` property is required here so that the TeamsFx conversation SDK can invoke the corresponding action handler when the action is invoked in Teams. You should provide a global unique string for the `verb` property, otherwise you may experience unexpected behavior if you're using a general string that might cause a collision with other bot actions.

### Step 2: Respond with a new Adaptive Card

For each action, you can display a new Adaptive Card as a response to the user. Create a new file, `src/adaptiveCards/doSomethingResponse.json` to use as a response for the `DoSomething` action created in the previous step:

```json
{
  "type": "AdaptiveCard",
  "body": [
    {
      "type": "TextBlock",
      "size": "Medium",
      "weight": "Bolder",
      "text": "A sample response to DoSomething."
    }
  ],
  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
  "version": "1.4"
}
```

You can use the [Adaptive Card Designer](https://adaptivecards.io/designer/) to help visually design your Adaptive Card UI.

### Step 3: Handle the new action

Create a new file, `src/cardActions/doSomethingActionHandler.js`:

```javascript
const ACData = require("adaptivecards-templating");
const responseCard = require("../adaptiveCards/doSomethingResponse.json");

class DoSomethingActionHandler {
  triggerVerb = "doSomething";

  async handleActionInvoked(context, message) {
    const cardJson = new ACData.Template(responseCard).expand({
      $root: {
        title: "doSomething command is added",
        body: "Congratulations! You have responded to doSomething command",
      },
    });
    return cardJson;
  }
}

module.exports = {
  DoSomethingActionHandler,
};
```

> Please note:
>
> - The `triggerVerb` is the `verb` property of your action.
>
> - The `actionData` is the data associated with the action, which may include dynamic user input or some contextual data provided in the `data` property of your action.
>
> - If an Adaptive Card is returned, then the existing card will be replaced with it by default.

You can customize what the action does here, including calling an API, processing data, etc.

### Step 4: Register the new handler

Navigate to the `src/index.js` file and register the trigger pattern to `app.adaptiveCards.actionExecute()`:

```javascript
const doSomethingActionHandler = new DoSomethingActionHandler();
app.adaptiveCards.actionExecute(doSomethingActionHandler.triggerVerb, async (context, state, data) => {
  return await doSomethingActionHandler.handleActionInvoked(context, data);
});
```

Congratulations, you've just created your own workflow! To learn more about extending the Workflow bot template, [visit the documentation on GitHub](https://aka.ms/teamsfx-workflow-new). You can find more scenarios like:

- [Customize the way to respond to an action](https://aka.ms/teamsfx-workflow-new#customize-the-action-response)
- [Customize the Adaptive Card content](https://aka.ms/teamsfx-workflow-new#customize-the-adaptive-card-content)
- [Create a user specific view](https://aka.ms/teamsfx-workflow-new#auto-refresh-to-user-specific-view)
- [Access Microsoft Graph](https://aka.ms/teamsfx-workflow-new#access-microsoft-graph)
- [Connect to existing APIs](https://aka.ms/teamsfx-workflow-new#connect-to-existing-apis)
- [Change the way to initialize the bot](https://aka.ms/teamsfx-workflow-new#customize-the-initialization)

## Extend workflow bot with other bot scenarios

Workflow bot is compatible with other bot scenarios like notification bot and command bot.

### Add notifications to your workflow bot

The notification feature adds the ability for your application to send Adaptive Cards in response to external events. Follow the [steps here](https://aka.ms/teamsfx-workflow-new#how-to-extend-workflow-bot-with-notification-feature) to add the notification feature to your workflow bot. Refer [the notification document](https://aka.ms/teamsfx-notification-new) for more information.

### Add command and responses to your workflow bot

The command and response feature adds the ability for your application to "listen" to commands sent to it via a Teams message and respond to commands with Adaptive Cards. Follow the [steps here](https://aka.ms/teamsfx-command-new#How-to-add-more-command-and-response) to add the command response feature to your workflow bot. Refer [the command bot document](https://aka.ms/teamsfx-command-new) for more information.

## Additional information and references

- [Manage multiple environments](https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-multi-env)
- [Collaborate with others](https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-collaboration)
- [Teams Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Teams Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Teams AI SDK](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/teams-conversational-ai/teams-conversation-ai-overview)
- [Teams Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)
