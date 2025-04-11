# GitHub Issues Graph connector

## Summary

This sample project uses Microsoft 365 Agents Toolkit for Visual Studio Code to simplify the process of creating a [Microsoft Graph connector](https://learn.microsoft.com/graph/connecting-external-content-connectors-overview) that ingests data from the GitHub issues API to Microsoft Graph. It provides an end to end example of creating the connector, ingesting content and refreshing the ingested content.

![External content in Microsoft 365 Copilot](./assets/copilot-results.png)

## Features

This sample shows how to ingest data from a custom API into your Microsoft 365 tenant.

The sample illustrates the following concepts:

- Simplify debugging and provisioning of resources with Microsoft 365 Agents Toolkit for Visual Studio code
- Create external connection schema
- Support full ingestion of data
- Support incremental ingestion of data
- Visualize the external content in Microsoft 365 Copilot

## Contributors

- [Sébastien Levert](https://github.com/sebastienlevert)

## Version History

Version|Date|Comments
-------|----|--------
1.0|December 03, 2024|Initial release

## Prerequisites

- [Microsoft 365 Agents Toolkit for Visual Studio Code](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension)
- [Azure Functions Visual Studio Code extension](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-azurefunctions)
- [Microsoft 365 Developer tenant](https://developer.microsoft.com/microsoft-365/dev-program) with [uploading custom apps enabled](https://learn.microsoft.com/microsoftteams/platform/m365-apps/prerequisites#prepare-a-developer-tenant-for-testing)
- [Node@18](https://nodejs.org)

## Minimal path to awesome - Debug against a real Microsoft 365 tenant

- Clone repo
- Open repo in VSCode
- Create a GitHub fine-grained token
  - Go to [GitHub](https://github.com)
  - Click on your profile picture and select **Settings**
  - In the left sidebar, click on **Developer settings**
  - In the left sidebar, click on **Personal access tokens**
  - In the left sidebar, click on **Fine-grained tokens**
  - Click on **Generate new token**
  - Give it a name and an expiration
  - Select the **All repositories** access
  - In the **Repository permissions** section, select
    - Issues: Read-Only
    - Metadata: Read-Only
  - Click on **Generate token**
  - Copy the token
- Fill env file in `env` folder
  - Open the `.env.local`. Add the `CONNECTOR_ID` value and update the `CONNECTOR_REPOS` value
  - Open the `.env.local.user` and add the your GitHub token as the `SECRET_CONNECTOR_ACCESS_TOKEN` value
- Press <kbd>F5</kbd>, follow the sign in prompts
- When prompted, click on the link in the console to perform the tenant-wide admin consent
- Wait for all tasks to complete
- In the web browser navigate to the [Search & Intelligence](https://admin.microsoft.com/#/MicrosoftSearch/Connectors) area in the Microsoft 365 Admin Center
- A table will display available connections. Locate the **GitHub Issues** connection. In the **Required actions** column, select the link to **Include Connector Results** and confirm the prompt
- Navigate to [Microsoft 365 Copilot](https://m365.cloud.microsoft/chat)
- Using the search box on top, search for: `Summarize the latest GitHub issues`. You should see the following result:

![External content in Microsoft 365 Copilot](assets/copilot-results.png)

> [!NOTE]  
> It can take a moment for the results to appear. If you don't see the results immediately, wait a few moments and try again.
> If you are getting results from the web, you can turn off web for better isolation of your connector results.
