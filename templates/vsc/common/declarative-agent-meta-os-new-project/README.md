# Build a Declarative Agent in an Office add-in using Microsoft 365 Agents Toolkit

An Office Add-in can be a skill in a Copilot agent. Because Office Add-ins use the Office JavaScript Library (Office.js) to perform read and write operations on Office documents, these operations become actions in the Copilot agent. The skill is implemented as an API plugin that calls the APIs in Office.js. Alternatively, you can think of the agent as a natural language interface for the add-in's functionality.

> Note:
> The combination of a Copilot agent with an Office Add-in is in preview. The following are the limitations during the initial preview:
> 
> - The feature is only enabled for Office on Windows and Office on the web. We're working to bring support to Office on Mac.
> - The feature is only enabled for Excel, PowerPoint, and Word. We're working to bring support to Outlook.
> - An add-in must use the [unified manifest for Microsoft 365](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/unified-manifest-overview) to be configured as a skill in Copilot.

## Prerequisites

- The latest even-numbered version of [Node.js](https://nodejs.org/).
- Requirements specified in [Requirements for Copilot extensibility options](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/prerequisites#requirements-for-copilot-extensibility-options).
- An Microsoft 365 account. If you do not have M365 account, apply one from [M365 developer program](https://developer.microsoft.com/en-us/microsoft-365/dev-program).

## Test the agent

1. Close all Office applications.
1. Open Microsoft 365 Agent Toolkit.
1. In the **Lifecycle** pane, select **Provision**. Among other things, provisioning does the following:
1. In a command prompt or Visual Studio Code **TERMINAL** in the root of the project, run `npm run dev-server` to start the server on localhost. Wait until you see a line in the server window that the app compiled successfully. This means the server is running and serving the files.
    
    > Note: If this is the first time in over a month you have run a local server for an Office Add-in on your computer, you may be prompted to delete an old certificate and to install a new one. Agree to both prompts.

1. The first step in testing depends on the platform.

   - To test in Office on Windows, open Excel, PowerPoint, or Word. In a few moments, the **Contoso Add-in** control group appears on the **Home** ribbon with two buttons. (If it doesn't appear on the ribbon, select the **Add-ins** button on the ribbon, and then select the name of your add-in in the flyout that opens.)
   - To test in Office on the web, in a browser, navigate to `https://excel.cloud.microsoft.com/`, and then create a new workbook.

1. Select the down-pointing arrow head beneath th **Copilot** button on the ribbon and select **App Skills** from the drop down list.
1. In the hamburger control in the **Copilot** pane. The name of your agent should be in the list of agents. (You may need to select **See more** to ensure that all agents are listed.) If the agent isn't, try one or both of the following actions.

   - Wait a few minutes and reload Copilot.
   - With Copilot open to the list of agents, click the cursor on the Copilot window and press <kbd>Ctrl</kbd>+<kbd>R</kbd>.

1. When the agent is listed, select it. The agent's pane opens, and conversation starters appear in the pane.
1. Select a conversation starter that makes sense for the Office application that you are working in (Excel, PowerPoint, or Word), and then press the **Send** control in the conversation box at the bottom of the pane. Select **Confirm** in response to the confirmation prompt.

   > [!TIP]
   > If Copilot reports an error, repeat your prompt but add the following sentence to the prompt: "If you get an error, report the complete text of the error to me."

1. Try entering other prompts that your agent should be able to carry out. 

## Edit the manifest

If you need to make changes to the features of the add-in and agent that are configured in the manifest, edit the `manifest.json` in the `./appPackage` folder. The [schema reference](https://raw.githubusercontent.com/OfficeDev/microsoft-teams-app-schema/preview/op/extensions/MicrosoftTeams.schema.json) for more information.

## Deploy to Azure

You can deploy the project to Azure from either Visual Studio Code or by using the Microsoft 365 Agents Toolkit CLI.

### From Visual Studio Code

1. Open Microsoft 365 Agents Toolkit, and sign into Azure by clicking the **Sign in to Azure** under the **ACCOUNTS** section from sidebar.
1. After you signed in, select a subscription under your account.
1. Select **Provision** from **LIFECYCLE** section or open the command palette and select: **Microsoft 365 Agents: Provision**.
1. Select **Deploy** or open the command palette and select: **Microsoft 365 Agents: Deploy**.

### With Microsoft 365 Agents Toolkit CLI

1. Run the command `m365agents auth login azure`.
1. (Optional) In the in env/.env.dev file, set environment variable AZURE_SUBSCRIPTION_ID to your subscription id, or set the variable in your current shell environment if you are using non-interactive mode of `m365agents` CLI.
1. Run the command `m365agents provision`.
1. Run the command: `m365agents deploy`.

> Note: Provisioning and deployment may incur charges to your Azure Subscription.
