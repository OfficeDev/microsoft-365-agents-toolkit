# How to use this Teams Tab app HelloWorld app

Microsoft Teams supports the ability to run web-based UI inside "custom tabs" that users can install either for just themselves (personal tabs) or within a team or group chat context.

## Prerequisites

- [Node.js](https://nodejs.org/), supported versions: 18, 20, 22
- An M365 account. If you do not have M365 account, apply one from [M365 developer program](https://developer.microsoft.com/en-us/microsoft-365/dev-program)
- [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)

## Debug

- From Visual Studio Code: Start debugging the project by hitting the `F5` key in Visual Studio Code.
- Alternatively use the `Run and Debug Activity Panel` in Visual Studio Code and click the `Run and Debug` green arrow button.
- From Microsoft 365 Agents Toolkit CLI:
  - Executing the command `teamsapp provision --env local` in your project directory.
  - Executing the command `teamsapp deploy --env local` in your project directory.
  - Executing the command `teamsapp preview --env local` in your project directory.

## Edit the manifest

You can find the app manifest in `../appPackage` folder. The folder contains one manifest file:

- `manifest.json`: Manifest file for app running locally or running remotely (After deployed to Azure).

This file contains template arguments with `${{...}}` statements which will be replaced at build time. You may add any extra properties or permissions you require to this file. See the [schema reference](https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema) for more information.

## Deploy to Azure

Deploy your project to Azure by following these steps:

| From Visual Studio Code                                                                                                                                                                                                                                                                                                                                                                                                                                              | From Microsoft 365 Agents Toolkit CLI                                                                                                                                               |
| :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| <ul><li>Open Microsoft 365 Agents Toolkit, and sign into Azure by clicking the `Sign in to Azure` under the `ACCOUNTS` section from sidebar.</li> <li>After you signed in, select a subscription under your account.</li><li>Open the Microsoft 365 Agents Toolkit and click `Provision` from DEVELOPMENT section or open the command palette and select: `Teams: Provision`.</li><li>Open the Microsoft 365 Agents Toolkit and click `Deploy` or open the command palette and select: `Teams: Deploy`.</li></ul> | <ul> <li>Run command `teamsapp auth login azure`.</li> <li>Run command `teamsapp provision --env dev`.</li> <li>Run command: `teamsapp deploy --env dev`. </li></ul> |

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

## Add Single Sign On feature

Microsoft Teams provides a mechanism by which an application can obtain the signed-in Teams user token to access Microsoft Graph (and other APIs). Microsoft 365 Agents Toolkit facilitates this interaction by abstracting some of the Microsoft Entra flows and integrations behind some simple, high-level APIs. This enables you to add single sign-on (SSO) features easily to your application.

Please follow this [document](https://aka.ms/teamsfx-add-sso-new) to add single sign on for your project.
