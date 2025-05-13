# Server-side code in applications

Azure Functions are a great way to add server-side behaviors to any application.

## Prerequisites

- [Node.js](https://nodejs.org/), supported versions: 18, 20, 22
- A Microsoft 365 account. If you do not have Microsoft 365 account, apply one from [Microsoft 365 developer program](https://developer.microsoft.com/en-us/microsoft-365/dev-program)
- [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)

## Develop

The Microsoft 365 Agents Toolkit IDE Extension and Microsoft 365 Agents Toolkit CLI provide template code for you to get started with Azure Functions for your application. Microsoft Teams Framework simplifies the task of establishing the user's identity within the Azure Functions.

The template handles calls from your Teams "custom tab" (client-side of your app), initializes the TeamsFx SDK to access the current user context, and demonstrates how to obtain a pre-authenticated Microsoft Graph Client. Microsoft Graph is the "data plane" of Microsoft 365 - you can use it to access content within Microsoft 365 in your company. With it you can read and write documents, SharePoint collections, Teams channels, and many other entities within Microsoft 365. Read more about [Microsoft Graph](https://docs.microsoft.com/en-us/graph/overview).

You can add your logic to the single Azure Functions created by this template, as well as add more functions as necessary. See [Azure Functions developer guide](https://docs.microsoft.com/en-us/azure/azure-functions/functions-reference) for more information.

### Call the Function

To call your Azure Functions, the client sends an HTTP request with an SSO token in the `Authorization` header. Here is an example:

```ts
import { TeamsUserCredentialAuthConfig, TeamsUserCredential } from "@microsoft/teamsfx";

const authConfig: TeamsUserCredentialAuthConfig = {
  clientId: "YOUR_CLIENT_ID",
  initiateLoginEndpoint: "YOUR_LOGIN_PAGE_URL",
};
const teamsUserCredential = new TeamsUserCredential(authConfig);
const accessToken = await teamsUserCredential.getToken(""); // Get SSO token
const endpoint = "https://YOUR_API_ENDPOINT";
const response = await axios.default.get(endpoint + "/api/" + functionName, {
  headers: {
    Authorization: `Bearer ${accessToken.token}`,
  },
});
```

### Add More Functions

- From Visual Studio Code, open the command palette, select `Microsoft 365 Agents: View How-to Guides` and select `Integrate with Azure Functions`.

## Change Node.js runtime version

By default, Microsoft 365 Agents Toolkit and Microsoft 365 Agents Toolkit CLI will provision an Azure functions app with function runtime version 3, and node runtime version 12. You can change the node version through Azure Portal.

- Sign in to [Azure Portal](https://azure.microsoft.com/).
- Find your application's resource group and Azure Functions app resource. The resource group name and the Azure functions app name are stored in your project configuration file `env.*.json`. You can find them by searching the key `AZURE_RESOURCE_GROUP_NAME` and `FUNCTION_APP_NAME` in that file.
- After enter the home page of the Azure Functions app, you can find a navigation item called `Configuration` under `settings` group.
- Click `Configuration`, you would see a list of settings. Then click `General settings` and update the `Node.js Version` value to `Node.js 18 LTS` or `Node.js 20 LTS` according to your requirement.
- After Click `OK` button, don't forget to click `Save` button on the top of the page.

Then following requests sent to the Azure Functions app will be handled by new node runtime version.

## Debug

- From Visual Studio Code: Start debugging the project by hitting the `F5` key in Visual Studio Code. Alternatively use the `Run and Debug Activity Panel` in Visual Studio Code and click the `Start Debugging` green arrow button.
- From Microsoft 365 Agents Toolkit CLI: Start debugging the project by executing the command `atk preview --local` in your project directory.

## Edit the manifest

You can find the app manifest in `./appPackage` folder. The folder contains one manifest file:

- `manifest.json`: Manifest file for app running locally or running remotely (After deployed to Azure).

This file contains template arguments with `${{...}}` statements which will be replaced at build time. You may add any extra properties or permissions you require to this file. See the [schema reference](https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema) for more information.

## Deploy to Azure

Deploy your project to Azure by following these steps:

| From Visual Studio Code                                                                                                                                                                                                                                                                                                                        | From Microsoft 365 Agents Toolkit CLI                                                                                                                          |
| :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | :---------------------------------------------------------------------------------------------------------------------------------------------- |
| <ul><li>Open Microsoft 365 Agents Toolkit, and sign into Azure by clicking the `Sign in to Azure` under the `ACCOUNTS` section from sidebar.</li> <li>After you signed in, select a subscription under your account.</li><li>Open the command palette and select: `Microsoft 365 Agents: Provision`.</li><li>Open the command palette and select: `Microsoft 365 Agents: Deploy`.</li></ul> | <ul> <li>Run command `atk auth login azure`.</li><li> Run command `atk provision`.</li> <li>Run command `atk deploy`. </li></ul> |

> Note: Provisioning and deployment may incur charges to your Azure Subscription.

## Preview

Once the provisioning and deployment steps are finished, you can preview your app:

- From Visual Studio Code

  1. Open the `Run and Debug Activity Panel`.
  1. Select `Launch Remote (Edge)` or `Launch Remote (Chrome)` from the launch configuration drop-down.
  1. Press the Play (green arrow) button to launch your app - now running remotely from Azure.

- From Microsoft 365 Agents Toolkit CLI: execute `atk preview --remote` in your project directory to launch your application.

## Validate manifest file

To check that your manifest file is valid:

- From Visual Studio Code: open the command palette and select: `Microsoft 365 Agents: Validate Application`.
- From Microsoft 365 Agents Toolkit CLI: run command `atk validate` in your project directory.

## Package

- From Visual Studio Code: open the command palette and select `Microsoft 365 Agents: Zip app package`.
- Alternatively, from the command line run `atk package` in the project directory.

## Publish to Teams

Once deployed, you may want to distribute your application to your organization's internal app store in Teams. Your app will be submitted for admin approval.

- From Visual Studio Code: open the command palette and select: `Microsoft 365 Agents: Publish`.
- From Microsoft 365 Agents Toolkit CLI: run command `atk publish` in your project directory.
