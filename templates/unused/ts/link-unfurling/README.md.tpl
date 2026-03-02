# Overview of the Link Unfurling app template

This template showcases an app that unfurls a link into an adaptive card when URLs with a particular domain are pasted into the compose message area in Microsoft Teams or email body in Outlook.

![hero-image](https://aka.ms/teams-app-test-tool-link-unfurling-hero-image)

## Get Started with the Link Unfurling app

> **Prerequisites**
>
> - [Node.js](https://nodejs.org/), supported versions: 18, 20, 22
> - [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)

> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

1. First, select the Microsoft 365 Agents Toolkit icon on the left in the VS Code toolbar.
2. Press F5 to start debugging which launches your app in Microsoft 365 Agents Playground using a web browser.
3. The browser will pop up to open Microsoft 365 Agents Playground.
4. Click the "+" button in the input box, select "Link Unfurling" and paste a link ending with `.botframework.com`. You should see an adaptive card unfurled. Click `Send to Conversation` to send it to the current chat or channel.

## What's included in the template

| Folder / File        | Contents                                                                                                                  |
| -------------------- | ------------------------------------------------------------------------------------------------------------------------- |
| `m365agents.yml`       | Main project file describes your application configuration and defines the set of actions to run in each lifecycle stages |
| `m365agents.local.yml` | This overrides `m365agents.yml` with actions that enable local execution and debugging                                      |
| `m365agents.playground.yml`| This overrides `m365agents.yml` with actions that enable local execution and debugging in Microsoft 365 Agents Playground.            |
| `.vscode/`           | VSCode files for local debug                                                                                              |
| `src/`               | The source code for the link unfurling application                                                                        |
| `appPackage/`        | Templates for the application manifest                                                                              |
| `infra/`             | Templates for provisioning Azure resources                                                                                |

The following files can be customized and demonstrate an example implementation to get you started.

| File                                    | Contents                                       |
| --------------------------------------- | ---------------------------------------------- |
| `src/index.ts`                          | Application entry point and `express` handlers |
| `src/linkUnfurlingApp.ts`               | The teams activity handler                     |
| `src/adaptiveCards/helloWorldCard.json` | The adaptive card                              |

## Extend this template

This section introduces how to customize or extend this template, including:

- [How to use Zero Install Link Unfurling in Teams](https://aka.ms/teamsfx-extend-link-unfurling#how-to-use-zero-install-link-unfurling-in-teams)
- [How to add link unfurling cache in Teams](https://aka.ms/teamsfx-extend-link-unfurling#how-to-add-link-unfurling-cache-in-teams)
- [How to customize Zero Install Link Unfurling's adaptive cards](https://aka.ms/teamsfx-extend-link-unfurling#how-to-customize-zero-install-link-unfurlings-adaptive-cards)
- [How to add stage view](https://aka.ms/teamsfx-extend-link-unfurling#how-to-add-stage-view)
- [How to add task module (Teams)](https://aka.ms/teamsfx-extend-link-unfurling#how-to-add-task-module-teams)
- [How to add adaptive card action (Teams)](https://aka.ms/teamsfx-extend-link-unfurling#how-to-add-adaptive-card-action-teams)
- [How to extend this template with Notification, Command and Workflow bot](https://aka.ms/teamsfx-extend-link-unfurling#how-to-extend-this-template-with-notification-command-and-workflow-bot)
