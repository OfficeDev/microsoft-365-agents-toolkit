# Changelog

> Note: This changelog only includes the changes for the stable versions of Teams Toolkit. For the changelog of pre-released versions, please refer to the [Teams Toolkit Pre-release Changelog](https://github.com/OfficeDev/TeamsFx/blob/dev/packages/vscode-extension/PRERELEASE.md).

## 5.12.0 - Dec 17, 2024

This update represents a minor version increment of the Teams Toolkit, introducing new features and addressing user-reported bugs. These incremental enhancements were previously documented in the prerelease version and a blog post: [Teams Toolkit for Visual Studio Code update – November 2024](https://devblogs.microsoft.com/microsoft365dev/teams-toolkit-for-visual-studio-code-update-november-2024/)

Below is a comprehensive list of new features, enhancements, and bug fixes implemented since the last stable release.

### New Features

- **Added Local Authentication Support for API ME and API Plugin**: We have introduced local authentication support for application templates of API-ME and API Plugin application with Microsoft Entra and OAuth. which eliminates the anonymous setup for APIs on local environment end to end and also helps developers verify the behavior when Teams and Microsoft 365 Copilot invokes an API without proper authorization. This feature includes built-in middleware that implements the token validation and a guided code tour explains how it works.
![Local Authentication](https://github.com/user-attachments/assets/dbf7173d-d11e-4a8d-b26d-fbfb6abfa976)

- **Process Termination For Port Conflict**: Port conflict is a common pain point for developers when debugging bot applications. It happens when the processes are not fully terminated when debug session stops. Now Teams Toolkit has added additional support to help terminate those unkilled processes to release ports.
![Port Conflicts](https://github.com/user-attachments/assets/208a21d5-0232-470b-a102-620e542a20cb)

- **Integrated with Microsoft Kiota for API Plugin Generation**: Now Teams Toolkit has integrated with [Microsoft Kiota](https://learn.microsoft.com/openapi/kiota/overview) when adding API Plugins into Declarative Agent. With this integration, developers are enabled to create API Plugins by browsing public API catalogs and Microsoft Kiota is capable of handling large OpenAPI Description Document. This is an experimental feature, for early access make sure to turn on [VSCode Settings](https://code.visualstudio.com/docs/getstarted/settings#_workspace-settings) with `Enable Microsoft Kiota`.

- **Using Assistant API on Azure OpenAI Service**: We have previously announced the support for Assistant API on Azure OpenAI Services in Teams Toolkit for Python language, and now it has expanded to TypeScript and JavaScript.

- **Manifest Schema Validation Enhancement**: Added support for manifest schema validation to cover localization files and default language files.

#### Enhancements

- Updated Bot application templates from `restify` to `express` server for better support of Node.js 20.
- Updated Tab application templates from `react-scripts` to `vite` for better performance and developer experience.
- Updated application templates and pre-requisite checker for Node.js 20 support.
- Updated application templates to the latest stable [App Manifest](https://learn.microsoft.com/microsoftteams/platform/resources/schema/manifest-schema#sample-app-manifest) version v1.19.

#### Bug Fixes

- Resolved a build error caused by changes to the `clientCitation` interface. [#12884](https://github.com/OfficeDev/teams-toolkit/pull/12884)
- Addressed persistent port conflict issues encountered by some users when stopping a debug session. [#12881](https://github.com/OfficeDev/teams-toolkit/pull/12881)
- Improved default instructions for Declarative Agent to ensure better compliance. [#12880](https://github.com/OfficeDev/teams-toolkit/pull/12880)
- Resolved an issue where Teams Toolkit could fail to generate authentication actions when creating a declarative agent with an API plugin using Microsoft Kiota. [#12764](https://github.com/OfficeDev/teams-toolkit/pull/12764)
- Update `@azure/storage-blob` to avoid AbortSignal error when deploying with `azureStorage/enableStaticWebsite`. [#12583](https://github.com/OfficeDev/teams-toolkit/pull/12583)
- Upgraded `path-to-regexp` and `body-parser` to fix vulnerability issues. [#12610](https://github.com/OfficeDev/teams-toolkit/pull/12610)
- Fixed an issue where `teamsapp.yml` when adding an API Plugin to the declarative agent. [#12656](https://github.com/OfficeDev/teams-toolkit/pull/12656)

## 5.10.1 - Nov 19, 2024

This patch version update of Teams Toolkit introduces several incremental enhancements:

- **Terminology Update**: The term `Copilot Agent` has been updated to `Agent` in the user interface. [#12672](https://github.com/OfficeDev/teams-toolkit/pull/12672)
- **Debug Session Improvement**: Resolved an issue where child processes were not fully terminated when a debug session was stopped. [#12673](https://github.com/OfficeDev/teams-toolkit/pull/12673)
- **Manifest Validation Fix**: Addressed a problem where manifest validation failed with multi-line user instructions for declarative agents. [#12677](https://github.com/OfficeDev/teams-toolkit/pull/12677)
- **Schema Validation Fix**: Corrected an issue where manifest schema validation was unsuccessful for projects that included an API Plugin. [#12679](https://github.com/OfficeDev/teams-toolkit/pull/12679)

## 5.10.0 - Oct 16, 2024

This update represents a minor version increment of the Teams Toolkit, introducing new features and addressing user-reported bugs. These incremental enhancements were previously documented in the prerelease version and a series of blog posts:

- [July Prerelease](https://devblogs.microsoft.com/microsoft365dev/teams-toolkit-for-visual-studio-code-update-july-2024/): You can now debug Teams apps directly in the desktop client, use managed identities for better security, and clean up development resources more easily.
- [August Prererelease](https://devblogs.microsoft.com/microsoft365dev/teams-toolkit-for-visual-studio-code-update-august-2024/): This update adds enhanced app validation, Python support for intelligent chatbots, and the ability to create custom copilots. It also introduces the Assistant API on Azure OpenAI Service for AI Agents.

Below is a comprehensive list of new features, enhancements, and bug fixes implemented since the last stable release.

### New Features

- **External File Support for Declarative Copilot Instructions**: Developers now have the ability to use an external file to author instructions for their declarative copilots and reference it in the manifest file. This greatly improves the authoring experience for longer instructions compared to using JSON files.
  ![External File](https://github.com/user-attachments/assets/fa13711c-fe8c-4155-bd7f-9e0a8e0ed606)

- **Plugin Integration for Declarative Copilot**: Teams Toolkit now allows developers to add a plugin as a skill to the declarative copilot. Developers can either add a new API plugin using an OpenAPI description document or reference an existing API plugin via its manifest file.
  ![Add Plugin](https://github.com/user-attachments/assets/009a63d0-8bc0-4449-8ba6-cef25779c140)

- **Enhanced App Validation**: Developers can now evaluate their app packages using the same test cases Microsoft employs during app review. The Enhanced App Validation feature in Teams Toolkit identifies any errors or warnings within your app package and provides clear guidelines for resolution. For more details on Microsoft test cases, refer to the [Teams Store validation guidelines](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/appsource/prepare/teams-store-validation-guidelines) and [Commercial marketplace certification policies](https://learn.microsoft.com/en-us/legal/marketplace/certification-policies).
  ![App Validation](https://github.com/user-attachments/assets/4c2b8c49-6a0a-4ea7-8796-a94464714463)

- **Generate an Intelligent Chatbot with Python**: Following the release of support for building [Custom Engine Copilot](https://learn.microsoft.com/microsoft-365-copilot/extensibility/overview-custom-engine-copilot) during Build 2024, which included the ability to "chat with" your own API, Teams Toolkit now extends this capability to the Python programming language.
  ![App Generator](https://github.com/user-attachments/assets/21efa344-aea5-4d44-bb78-aa8e26dc68a1)

- **Create Declarative Copilot**: Teams Toolkit now allows you to build a declarative copilot, enabling you to customize Microsoft 365 Copilot by declaring specific instructions, actions, and knowledge. Declarative copilots run on the same orchestrator, foundation models, and trusted AI services that power Microsoft Copilot. You can learn more about [declarative copilots here](https://learn.microsoft.com/microsoft-365-copilot/extensibility/overview-declarative-copilot). The toolkit supports the creation of both basic declarative copilots and those with an API plugin.
  ![Declarative Copilot](https://github.com/user-attachments/assets/37412cdd-c7e8-4e38-bd45-794997b050ec)

- **Using Assistant API on Azure OpenAI Service**: The Teams Toolkit has updated the `AI Agent` (Python) app template to support the Assistant API on Azure OpenAI Service. You can now build your own AI Agents on Microsoft 365 using Python, with the option to use either Azure OpenAI Service or OpenAI directly. Support for TypeScript and JavaScript is forthcoming.

- **Debug Apps in Teams Desktop Client**: The Teams desktop client now offers a faster and more reliable way to debug your Teams applications, with the same capabilities available in the Teams web client, such as breakpoints and hot reload. This feature is now available for Custom Engine Copilots, Bots, and Message Extensions apps.
  ![Debug in Desktop](https://github.com/OfficeDev/teams-toolkit/assets/11220663/dc85ee11-e847-40d7-bceb-b5dc3e83f040)

- **Use Managed Identity for Bot and Message Extension when deploying to Azure**: The Teams Toolkit has transitioned from client ID and secret-based identity to user-assigned managed identity for Bot and Message Extension application templates, enhancing security. [Learn more](https://learn.microsoft.com/entra/identity/managed-identities-azure-resources/overview) about the benefits of using managed identities for Azure resources.
  ![MSI](https://github.com/OfficeDev/teams-toolkit/assets/11220663/b2ffddb2-8c04-4ee4-aaaa-ae7c666af6e1)

- **Clean Up Resources Created After Development**: You can now safely clean up resources created after application development by deleting the application registration in the Teams Developer Portal and Bot Framework Portal, and removing uploaded custom apps in Microsoft 365 applications. This can be done via the `teamsapp uninstall` command, either by using the App ID in the Teams application manifest file or by specifying an environment if your project is managed by the Teams Toolkit.
  ![Uninstall](https://github.com/OfficeDev/teams-toolkit/assets/11220663/294447b7-d5f9-47cc-ab37-9235dbd5c111)

- **Integrated CodeTour Instructions for Using Graph Connector Data Source**: The `Chat With Your Data - Microsoft 365` app template in Teams Toolkit now includes interactive CodeTour instructions. By default, the app uses content uploaded to SharePoint, but with these instructions, you can easily switch to a Graph connector data source if you have external content. Learn more about using the [Graph connector](https://learn.microsoft.com/microsoft-365-copilot/extensibility/overview-graph-connector).
  ![Code Tour](https://github.com/OfficeDev/teams-toolkit/assets/11220663/be2eb3d6-0468-4316-8e6f-e8025408045a)

- **Build AI Agent With Assistant API and Python**: Previously we have included the AI Assistant Bot app template to help you get started with building a GPT-like chat bot with AI capabilities using `Teams AI Library`. Now we have added a new AI Agent app template to help you build an AI agent with Assistant API and Python. This template showcases how to build an intelligent chat bot in Teams capable of helping users accomplish a specific task using natural language right in the Teams conversations, such as solving a math problem.

### Enhancemens

- Teams Toolkit will continue to update scaffold app templates to ensure compliance with [Teams Store validation guidelines](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/appsource/prepare/teams-store-validation-guidelines). The first round of updates focuses on bot templates, including:
  - [PR#12063](https://github.com/OfficeDev/teams-toolkit/pull/12063): Updated `Basic Bot` and `Message Extension`
  - [PR#12096](https://github.com/OfficeDev/teams-toolkit/pull/12096): Updated `Chat Command`
  - [PR#12123](https://github.com/OfficeDev/teams-toolkit/pull/12123): Updated `Chat Notification Messages`
  - [PR#12119](https://github.com/OfficeDev/teams-toolkit/pull/12119): Updated `Sequential Workflow in Chat`
- Teams Toolkit now prompts users to generate an API key before debugging API ME or API Plugin with API Key authentication templates.
- Secret values have been redacted from the Visual Studio Code output channel.
- Updated application templates to use the latest [manifest schema version v1.17](https://learn.microsoft.com/microsoftteams/platform/resources/schema/manifest-schema).
- Improved the readability of error messages generated by the Teams Toolkit.

### Bug Fixes

- Upgraded the axios dependency used in Teams Toolkit to version 1.7.6 to fix a vulnerability issue. [#12306](https://github.com/OfficeDev/teams-toolkit/pull/12306)
- Changed a string for better clarity when creating an `AI Agent` without Assistant API. [#12266](https://github.com/OfficeDev/teams-toolkit/pull/12266)
- Fixed vulnerability issues in TeamsFx SDK. [#11973](https://github.com/OfficeDev/teams-toolkit/pull/11937)
- Resolved compatibility issues with `groupchat` and `groupChat` in the Teams app manifest. [#12028](https://github.com/OfficeDev/teams-toolkit/pull/12028)
- Corrected an issue where the link redirection for the lifecycle `Provision` button was incorrect. [#12120](https://github.com/OfficeDev/teams-toolkit/pull/12120)
- Fixed initialization failures of `publicClientApplication` in TeamsFx SDK. [#12159](https://github.com/OfficeDev/teams-toolkit/pull/12159)
- Addressed issues when creating SharePoint Framework-based tab apps. [#12173](https://github.com/OfficeDev/teams-toolkit/pull/12173)
- Resolved an issue where users still saw a pop-up window when logging into a Microsoft 365 account in non-interactive mode. [#11978](https://github.com/OfficeDev/teams-toolkit/pull/11978)
- Fixed an issue where importing an SPFx project failed due to case-sensitive file systems on Ubuntu. [#11972](https://github.com/OfficeDev/teams-toolkit/pull/11972)
- Addressed an issue where debugging an Outlook Add-in might fail with the error `Package is invalid`. [#11963](https://github.com/OfficeDev/teams-toolkit/pull/11963)
- Corrected unclear error messages for commands that only work for projects created by the Teams Toolkit. [#11945](https://github.com/OfficeDev/teams-toolkit/pull/11945)
- Fixed a vulnerability issue with `ws` affected by a DoS when handling a request with many HTTP headers. [#11937](https://github.com/OfficeDev/teams-toolkit/pull/11937)
- Fixed an issue where sometimes you may not be able to scroll down in Teams Toolkit CLI. [#11762](https://github.com/OfficeDev/teams-toolkit/pull/11762)
- Fixed an issue where Teams Toolkit generated Adaptive Cards may contain empty property. [#11759](https://github.com/OfficeDev/teams-toolkit/pull/11759)
- Fixed an issue where you may need to press enter twice after selecting resource group during provision using Teams Toolkit CLI. [#11724](https://github.com/OfficeDev/teams-toolkit/pull/11724)
- Fixed an issue to enable shell option in Windows platform to avoid [command injection via args parameters](https://nodejs.org/en/blog/vulnerability/april-2024-security-releases-2#command-injection-via-args-parameter-of-child_processspawn-without-shell-option-enabled-on-windows-cve-2024-27980---high). [#11699](https://github.com/OfficeDev/teams-toolkit/pull/11699)
- Fixed an issue where provision summary logs are printed twice. [#11658](https://github.com/OfficeDev/teams-toolkit/pull/11658)

## 5.8.2 - Sep 26, 2024

Hotfix version.

- Fixed an issue ([#12143](https://github.com/OfficeDev/teams-toolkit/issues/12143)) where users may encounter an error (`SPFx.LatestPackageInstallFailed`) when creating a Tab with SPFx project.
- Resolved an issue ([#12384](https://github.com/OfficeDev/teams-toolkit/issues/12384)) where users may encounter an error (`AppStudioPlugin.TeamsAppCreateConflict`) during local debugging.
- Updated TypeScript from `4.4.4` to `5.5.4` to avoid build failures for `Custom Engine Agent` templates in [#12462](https://github.com/OfficeDev/teams-toolkit/pull/12462).

## 5.8.1 - May 27, 2024

Hotfix version.

- Resolved an issue that occurred when the Teams Toolkit extension was used with VS Code versions v1.87.2 or earlier. See issue [#11679](https://github.com/OfficeDev/teams-toolkit/issues/11679) for more details.
- Fixed the launch URL issue in the API-based message extension template with Microsoft Entra authentication.
- Fixed an issue where a `Create` label is unexpectedly shown in the Create New App dialog.

## 5.8.0 - May 14, 2024

This update, marking a minor version increment of the Teams Toolkit, incorporates new features and resolves bugs based on user input. Previously, these incremental modifications were detailed in the prerelease version and a series of blog posts:

- [March Prerelease](https://devblogs.microsoft.com/microsoft365dev/teams-toolkit-for-visual-studio-code-update-march-2024/): Included enhancements to the user interface and application templates for custom copilots, introducing a basic AI chatbot and AI agent, along with Python language support for constructing custom copilots using the Teams AI Library.
- [April Prerelease](https://devblogs.microsoft.com/microsoft365dev/teams-toolkit-for-visual-studio-code-update-april-2024/): Introduced the ability to utilize authentication-protected APIs in API-based message extensions, debug message extensions using the Teams App Test Tool, and build custom copilots with retrieval augmented generation employing Azure AI Search, custom APIs, and Microsoft 365 data.

### New Features

- **Create API based Message Extensions using auth-protected API**: Teams Toolkit now supports two forms of API authentication protection for your API-based Message Extension app:
  ![add-auth-api-me](https://github.com/OfficeDev/TeamsFx/assets/113089977/c5faea2f-676b-4a8c-82d6-f3b037e54f0e)

  - API-Key: You can either add the API key of your existing API or, if you don't have one, Teams Toolkit will generate one to demonstrate authentication flow.
  - Microsoft Entra (Azure AD): Teams Toolkit facilitates the creation of Microsoft Entra IDs to authenticate your new API.

- **Debug Message Extensions in Teams App Test Tool**: Developers can now debug and test the message extensions in a web-based environment using the Teams App Test Tool, which emulates Microsoft Teams without requiring tunnels or a Microsoft 365 account. This version extends Teams App Test Tool support to search-based, action-based, and link-unfurling message extension apps.
  ![ME-test-tool](https://github.com/OfficeDev/TeamsFx/assets/113089977/2b55996f-87a9-4683-abaf-3089b7ea878e)

- **Create Intelligent Chatbot with RAG Capability**: The Custom Copilot with RAG capability is an AI-powered chatbot capable of understanding natural language and retrieving domain-specific data to answer questions. Teams Toolkit now supports accessing custom data in Custom Copilot apps. When creating a Custom Copilot app, developers can select 'Chat with your Data' and then choose the desired data source.
  ![access-data-custom-copilot](https://github.com/OfficeDev/TeamsFx/assets/113089977/d40cfc84-8cb8-4816-b587-668a2bcf9560)
  Four types of data sources are available:
  ![data-source-custom-copilot](https://github.com/OfficeDev/TeamsFx/assets/113089977/2d010366-96a0-4f8b-861d-28d5bb9e36b8)

  - `Customize`: Load data from a custom data source, for example, the file system or vector Database.
  - `Azure AI Search`: Load data from Azure AI Search service and use it in conversation with users.
  - `Custom API`: Invoke the API defined in the OpenAPI description document to retrieve domain data from the API service.
  - `Microsoft 365`: Query Microsoft 365 data from the Microsoft Graph Search API as a data source in the conversation.

- **Build Your Own Copilots in Teams with Teams AI Library**: Enhancements have been made to the user experience for developers to create their custom copilots, AI-powered intelligent chatbots for Teams, with the following improvements:
  ![Custom Copilots](https://github.com/OfficeDev/TeamsFx/assets/11220663/0387a2ce-ec39-4c72-aabc-1ec2b9e85d59)
  - Streamlined UX for scaffolding, including top-level entry points and easy configuration of LLM services and credentials during the scaffolding flow.
  - New application templates allowing developers to build an AI Agent from scratch.
  - Python language support for building a `Basic AI Chatbot` and `AI Agent`.

### Enhancements

- The `aadApp/create` action now supports setting `serviceManagementReference` for Entra apps, facilitating referencing application or service contact information from a Service or Asset Management database.
- The `office-addin-manifest` and `office-addin-project` have been upgraded to the latest version to enable the use of the latest version of the manifest schema when importing an existing XML-based Outlook Add-in project.
- The `@azure/identity` and `@azure/msal-node` packages have been updated to the latest version to ensure the security of the authentication process.
- The `teamsjs` dependency in `teamsfx` and `teamsfx-react` has been upgraded to the latest version to ensure compatibility with the latest TeamsJS SDK.

### Bug Fixes

- Fixed an issue in the `teamsapp` CLI, where scrolling down would result in an error when selecting an item of a large index from a list. ([#11521](https://github.com/OfficeDev/TeamsFx/pull/11521))
- Fixed a bug in API-based message extension app templates, where the launch URL pointed to the home page instead of the app installation page. ([#11461](https://github.com/OfficeDev/TeamsFx/pull/11461))
- Fixed a vulnerability issue introduced by the `axios` package. ([#11449](https://github.com/OfficeDev/TeamsFx/pull/11449))

## 5.6.0 - Mar 12, 2024

This minor version update of Teams Toolkit includes new features and bug fixes based on your feedback. The new features include Deploy Tab Apps to Static Web App, Teams Toolkit CLI v3 and many other enhancements. We previously shared these incremental changes in the prerelease version and through a blog post:

- [Janaury Prerelease](https://devblogs.microsoft.com/microsoft365dev/teams-toolkit-for-visual-studio-code-update-january-2024/): Deploy Tab Apps to Static Web App, Teams Toolkit CLI v3, new Link Unfurling sample app and many other enhancements.

We've listened to your feedback and included these additional new features, enhancements, and bug fixes to this release.

New features:

- **Deploy Tab Apps to Static Web App**: Azure Static Web Apps, an automatic service for building and deploying full-stack web apps to Azure from a code repository, is now the default solution for deploying Tab-based applications in Teams Toolkit. If you prefer the old way using Azure Storage, please refer to this [sample](https://github.com/OfficeDev/TeamsFx-Samples/tree/dev/hello-world-tab-codespaces).
- **Teams Toolkit CLI ([`@microsoft/teamsapp-cli`](https://www.npmjs.com/package/@microsoft/teamsapp-cli)) `v3.0.0`**. Teams Toolkit CLI version 3 is now released in stable version. Major changes include:
  ![Teams Toolkit CLI](https://camo.githubusercontent.com/67608a468cbd406d6ff18585c8bc3b34d3d97d0a8ef525bdf516ca23fd5e32dd/68747470733a2f2f616b612e6d732f636c692d6865726f2d696d616765)
  - **New Command Signature**: Teams Toolkit CLI now starts with `teamsapp` as the root command signature for more clarity. We recommend changing your scripts to use `teamsapp` as the command prefix.
  - **New Command Structure**: Teams Toolkit CLI now has a new command structure that is more intuitive and easier to use. You can find the new command structure in the [Teams Toolkit CLI Command Reference](https://aka.ms/teamsfx-toolkit-cli).
  - **New Doctor Command**: `teamsapp doctor` command is a new command that helps diagnose and fix common issues with Teams Toolkit and Teams application development.

Enhancements:

- **Format Reddit Link into Adaptive Card Sample**: This sample application demonstrates how to format a Reddit link into an Adaptive Card in Microsoft Teams conversations.
  ![Link Unfurling Sample](https://github.com/OfficeDev/TeamsFx/assets/11220663/0d44f8c3-d02e-4912-bfa2-6ed3fdb29c1b)
- **Clean up `.deployment` Folder in between Deployments**: Teams Toolkit now cleans up the `.deployment` folder in the build directory before each deployment, addressing a [known issue](https://github.com/OfficeDev/TeamsFx/issues/10075) and reducing deployment time.
- **Optimized Dev Tunnel Expiration**: Inactive Dev Tunnel instances will now be automatically cleaned up after an hour, mitigating Dev Tunnel instance limitation errors.
- **Log Level Settings**: Added log level settings for controlling the verbosity of Teams Toolkit logs. You can find the settings in the [User and Workspace Settings](https://code.visualstudio.com/docs/getstarted/settings) under the `Teams Toolkit` section.
  ![Logs](https://github.com/OfficeDev/TeamsFx/assets/11220663/3a1fc3a0-d69b-446e-8db2-0c756a18f95e)
- **Richer Information in Sample App Details Page**: The Sample app detail page now includes additional details from the project README file, such as the project description, prerequisites, and steps to run the project.
- **Improved Troubleshooting for Multi-tenant Scenario**: Teams Toolkit now provides a [troubleshooting guide](https://aka.ms/teamsfx-multi-tenant) for scenarios where `aadApp/update` action fails with a `HostNameNotOnVerifiedDomain` error in multi-tenant setups.
- **Optimized SPFx Solution Version Handling**: Teams Toolkit now compares the SPFx solution version between global installations and the one used by Teams Toolkit when developers add additional web parts. Developers will be prompted if there's a need to install or upgrade the solution version when differences are detected.

## 5.4.1 - Feb 07, 2024

Hotfix version.

We have made UI and docs updates to multiple places according to the [latest updates to the Microsoft 365 Developer Program](https://devblogs.microsoft.com/microsoft365dev/stay-ahead-of-the-game-with-the-latest-updates-to-the-microsoft-365-developer-program/).

## 5.4.0 - Dec 18, 2023

This minor version update of Teams Toolkit includes new features and bug fixes based on your feedback. The new features include the Adaptive Card Previewer, Teams App Test Tool, new project templates for AI Assistant Bot, and a refreshed Sample Gallery. We previously shared these incremental changes in the prerelease version and through a series of blog posts:

- [October Prerelease](https://devblogs.microsoft.com/microsoft365dev/teams-toolkit-update-october-2023): Search sample app, new Contoso Retail Dashboard sample app, Teams Toolkit CLI improvement.
- [November Prerelease](https://devblogs.microsoft.com/microsoft365dev/teams-toolkit-for-visual-studio-code-update-november-2023): Brand-new AI Assistant Bot App template, Teams app test tool, integrated Adaptive Card Previewer, Refreshed look for Sample Gallery etc.

We've listened to your feedback and included these additional new features, enhancements, and bug fixes to this release.

New features:

- [Preview Feature] AI Assistant Bot App template using Teams AI library and the OpenAI Assistants API. This template demonstrates how to build an intelligent chatbot within Teams.
  ![AI-assistant-bot](https://github.com/OfficeDev/TeamsFx/assets/113089977/13877e7b-cffb-4717-9d1e-587478a43e5d)
- [Preview Feature] Teams App Test Tool removes the need for a lot of the setup to run and debug a Teams bot in a web-based chat environment that is similar to Teams.
  ![Teams app test tool](https://github.com/OfficeDev/TeamsFx/assets/113089977/7f6bcf56-27b3-4a7b-a160-64ab847efd92)
- [Preview Feature] Preview and edit Adaptive Cards effortlessly using the Adaptive Card Previewer now included in Teams Toolkit.
  ![Adaptive Card Previwer](https://github.com/OfficeDev/TeamsFx/assets/113089977/8062f237-d640-45b2-a7b2-81a0a4b0aa10)
- [Generally Available] Navigate quickly through samples using list or grid views. You can also search samples or filter samples by type, capability, and programming language.
  ![New sample gallery UI](https://github.com/OfficeDev/TeamsFx/assets/113089977/cbb18477-1f6d-4ab5-a947-1840d3290e40)
- [Preview Feature] Developing with Microsoft Copilot? We’ve added a helpful license check UI to ensure that your account has the necessary Microsoft Copilot licenses.
  ![Copilot license check](https://github.com/OfficeDev/TeamsFx/assets/113089977/78fa1fc5-0e19-40f0-b324-65583485f743)

Enhancement:

- New samples in the Sample Gallery:
  ![new samples](https://github.com/OfficeDev/TeamsFx/assets/113089977/2af41ec4-ee19-4b66-a58a-d2d8bdbbbd60)

  - Large Scale Notification Bot: send individual chat messages to a large number of users in a tenant
  - Graph Connector Bot: Teams command bot that queries custom data ingested into Microsoft Graph using Graph connector.
  - Contoso Retail Dashboard: demonstrates how to build a dashboard for control layouts. It runs in Teams, Outlook, and the Microsoft 365 app.
  - Teams Conversation Bot using Python: Python example for Teams Bot app.
  - Teams Messaging Extensions Search using Python: Python example for Teams Messaging Extensions app.
  - SSO Enabled Tab via APIM Proxy: Teams tab app that can display user login information with SSO.
  - Ingest Custom API Data into Microsoft 365 with Microsoft Graph Connector: simplify the process of creating a Microsoft Graph connector that ingests data from a custom API to Microsoft Graph

- Recommended Regions in Provision
  ![Provision Region](https://github.com/OfficeDev/TeamsFx/assets/113089977/97867d08-b7af-4eae-b1e7-d0102e1a1361)
- Automatic `npm install` for SPFx Tab App
  ![npm install for SPFx](https://github.com/OfficeDev/TeamsFx/assets/113089977/514d262d-9695-40dc-91aa-5c35044a319d)
- Teams Toolkit CLI Enhancement including: Commands have been reorganized into a hierarchical structure, added a teamsfx list command, improve the help command readability, outputs have been refreshed and log levels have been streamlined for clarity.
- Update Teams AI chat bot template to use latest teams-ai library.

Bug Fixes:

- Resolved an issue causing an No localized strings file found error in Visual Studio Code output. ([#10090](https://github.com/OfficeDev/TeamsFx/pull/10090))
- Fixed a flickering issue when selecting an option in Quick Pick. ([#10100](https://github.com/OfficeDev/TeamsFx/pull/10100))
- Corrected a string typo in the Create a New App dialog. ([#10197](https://github.com/OfficeDev/TeamsFx/pull/10197))
- Fixed the notification store path missing issue when using Test tool for Teams notification bot app. ([#10499](https://github.com/OfficeDev/TeamsFx/pull/10499))

## 5.2.0 - Sep 19, 2023

This minor version update of Teams Toolkit includes new features and bug fixes based on your feedback. The new features include the ability to import an existing SharePoint Framework solution, new project templates for Link Unfurling and AI Chat Bot, and more. We previously shared these incremental changes in the prerelease version and through a series of blog posts:

- [July Prerelease](https://devblogs.microsoft.com/microsoft365dev/teams-toolkit-for-visual-studio-code-update-july-2023): Import an existing SharePoint Framework solution, implement link unfurling in Teams and Outlook.
- [August Prerelease](https://devblogs.microsoft.com/microsoft365dev/teams-toolkit-for-visual-studio-code-update-with-new-ai-chat-bot-template/): Create an AI chat bot with the new AI Chat Bot template.

We've listened to your feedback and included these additional new features, enhancements, and bug fixes to this release.

New features:

- Import an existing SharePoint Framework solution and continue development with Teams Toolkit.
  ![SPFx Existing App](https://github.com/OfficeDev/TeamsFx/assets/11220663/3944f5c8-6c8c-4b4d-8df8-dc4f45b5967f)
- A new link unfurling project template to help you get started with displaying rich content from links in Teams messages and Outlook emails.
  ![Link Unfurling](https://github.com/OfficeDev/TeamsFx/assets/11220663/6e8b982a-0531-4ec1-8420-f6f17955ff40)
- A new AI Chat Bot project template to help you get started with building a GPT-like chat bot with AI capabilities using the [Teams AI Library](https://github.com/microsoft/teams-ai).
  ![AI Bot](https://github.com/OfficeDev/TeamsFx/assets/11220663/86a90d2a-efc3-4d8b-9e8c-5d34a1e8c081)
- The Sample Gallery has a new sample, One Productivity Hub using Graph Toolkit with SPFx, that shows you how to build a Tab for viewing your calendar events, to-do tasks, and files using Microsoft Graph Toolkit components and a SharePoint provider.
  ![SPFx Sample](https://github.com/OfficeDev/TeamsFx/assets/11220663/084ac508-49ea-4b30-854c-8b4d578ff6ee)
- Run life-cycle commands like Provision, Deploy, and Publish using new CodeLens hints added in-line to `teamsapp.yml`` when editing the file.
  ![Inline Commands](https://github.com/OfficeDev/TeamsFx/assets/11220663/f6897b26-0e3c-441c-b028-32093e8322a7)

Bug fixes:

- Fixed an issue where the `Preview` tag for the `AI Chat Bot` project template was not included in previous version. ([#9901](https://github.com/OfficeDev/TeamsFx/pull/9901))
- Fixed an issue where sometimes Teams Toolkit may fail to activate due to an unexpected package dependency on `types/keytar`. ([#9910](https://github.com/OfficeDev/TeamsFx/pull/9910))

## 5.0.1 - June 20, 2023

Incremental version for Teams Toolkit with multiple bug fixes:

- Fixed an issue where older versions of Teams Toolkit CLI and Node.js were referenced in CI/CD templates. ([#8972](https://github.com/OfficeDev/TeamsFx/pull/8972))
- Fixed an issue [#8929](https://github.com/OfficeDev/TeamsFx/issues/8929) in [`teamsApp/create`](https://github.com/OfficeDev/TeamsFx/wiki/Available-actions-in-Teams-Toolkit#teamsappcreate) and [`teamsApp/update`](https://github.com/OfficeDev/TeamsFx/wiki/Available-actions-in-Teams-Toolkit#teamsappupdate) actions where detailed error messages were not printed in output channel when Teams Toolkit failed to create or update app in Teams Developer Portal. ([#8967](https://github.com/OfficeDev/TeamsFx/pull/8967))
- Fixed an issue in [`script`](https://github.com/OfficeDev/TeamsFx/wiki/Available-actions-in-Teams-Toolkit#script) action where not all env output pairs were parsed. ([#8811](https://github.com/OfficeDev/TeamsFx/pull/8811))
- Fixed an issue in [`script`](https://github.com/OfficeDev/TeamsFx/wiki/Available-actions-in-Teams-Toolkit#script) action where charset encoding were not properly handled. ([#8769](https://github.com/OfficeDev/TeamsFx/pull/8769))
- Fixed an issue where the app might fail to start the dev tunnel for local development due to an unexpected token error. ([#8980](https://github.com/OfficeDev/TeamsFx/pull/8980))
- Fixed an issue where migration failed due to the expected path of the `.gitignore` file. ([#8912](https://github.com/OfficeDev/TeamsFx/pull/8912))
- Fixed an issue [#8853](https://github.com/OfficeDev/TeamsFx/issues/8853) in [`botAadApp/create`](https://github.com/OfficeDev/TeamsFx/wiki/Available-actions-in-Teams-Toolkit#botaadappcreate) action where no detailed errors were printed when Teams Toolkit failed to create the app in Microsoft Entra. ([#8910](https://github.com/OfficeDev/TeamsFx/pull/8910))
- Fixed an issue in [`aadApp/create`](https://github.com/OfficeDev/TeamsFx/wiki/Available-actions-in-Teams-Toolkit#aadappcreate) and [`aadApp/update`](https://github.com/OfficeDev/TeamsFx/wiki/Available-actions-in-Teams-Toolkit#aadappupdate) actions where no detailed error messages were printed when Teams Toolkit failed to create or update the app in Microsoft Entra. ([#8911](https://github.com/OfficeDev/TeamsFx/pull/8911))
- Fixed an issue where you might see multiple login prompts when Teams Toolkit failed to retrieve the token by automatically clearing the cache. ([#9026](https://github.com/OfficeDev/TeamsFx/pull/9026))

## 5.0.0 - May 16, 2023

This major version update of Teams Toolkit addresses your top feedback requests with new features and bug fixes, including a new way to customize the automation with composable actions, integrated tunneling support for debugging using Dev Tunnels, simpler project structure and template options, and much more. We previously shared these incremental changes in the prerelease version and through a series of blog posts:

- [February Prerelease](https://devblogs.microsoft.com/microsoft365dev/teams-toolkit-for-visual-studio-code-v5-0-prerelease/): Update the fundamental new design for Teams Toolkit that allows you to use existing cloud resources, integrate Teams Toolkit with existing projects and customize Teams Toolkit to fit your needs.
- [March Prerelease](https://devblogs.microsoft.com/microsoft365dev/teams-toolkit-for-visual-studio-code-update-march-2023/): Create, debug and deploy an Outlook Add-in project with Teams Toolkit.
- [April Prerelease](https://devblogs.microsoft.com/microsoft365dev/teams-toolkit-for-visual-studio-code-update-april-2023/): Switch to an integrated and secure tunneling for bot and message extension apps. Introduced GitHub Codespaces for a new getting started experience.

We've listened to your feedback and included these additional new features, enhancements, and bug fixes to this release.

New features:

- Added a how-to guide with instructions for running and debugging apps on mobile clients.
- Simplified the Basic Tab project template by removing the dependency on React, single sign-on, and complicated example code. Use this template like an empty starting point for Tab apps.
- You can customize which version of Azure Functions Core Tools is used with the `devTool/install` action. If not specified, the default version used is `4.0.4670` and [supports Node.js 18](https://learn.microsoft.com/azure/azure-functions/functions-versions?tabs=v4&pivots=programming-language-typescript#languages).
- We've re-categorized the project templates to use familiar terminology that matches the documentation and platform.
  ![create-new-app](https://github.com/OfficeDev/TeamsFx/assets/11220663/fe3ac358-775d-4deb-9b1e-a9eb4d932e56)

Enhancements:

- Removed SPFx version selection when adding additional web parts.
- Updated the `Tab App with Azure Backend` sample to use an on-behalf-of flow for authentication.
- Added a warning message if you're using Node.js 14 because it's [no longer actively supported](https://nodejs.dev/en/about/releases/).

Sample additions:

- `Dice Roller in meeting`: The dice roller example is a simple app that allows multiple users to roll a dice and see the results. This sample is a great way to get started with Live Share and Fluid.
- `Set signature using Outlook add-in`: An Outlook add-in that demonstrates how to set an email signature.

Bug fixes:

- Fixed an issue with `xml2js` package versions that may cause security vulnerabilities. ([#8390](https://github.com/OfficeDev/TeamsFx/pull/8390))
- Fixed an issue where `manifest-file-path` will be asked twice even if it was specified in the parameter for `teamsfx update aad-app` command. ([#8435](https://github.com/OfficeDev/TeamsFx/pull/8435))
- Fixed an issue where the `Validate Application` continues with errors. ([#8467](https://github.com/OfficeDev/TeamsFx/pull/8467))
- Fixed an issue where `teamsfx validate` command will default to `dev` environment even for multi-environment projects in non-interactive cli mode when not specifying `--env` parameters. ([#8499](https://github.com/OfficeDev/TeamsFx/pull/8499))
- Fixed an issue where sometimes Teams Toolkit tree view will not show up if project upgrade detects error.([#8538](https://github.com/OfficeDev/TeamsFx/pull/8538))

## 4.2.5 - May 04, 2023

Incremental version for Teams Toolkit with an enhancement.

- Added a note on pre-requisite checker for usage of [NGROK](https://ngrok.com/).

## 4.2.4 - Feb 28, 2023

Incremental version for Teams Toolkit with an enhancement.

- Updated SharePoint Framework from `1.16.0` to `1.16.1`.

## 4.2.3 - Feb 17, 2023

Incremental version for Teams Toolkit with multiple bug fixes.

- Fixed an unexpected string change in the `create new app` flow.
- Fixed an issue where there's no response when clicking on the `Get Started` button.
- Fixed an issue where you might see `Something went wrong` when adding the bot app on Teams during a debug session.

## 4.2.2 - Feb 7, 2023

Incremental version for Teams Toolkit with multiple bugs fixes and enhancements.

Enhancements:

- Updated pre-requisite checker, applicable app templates and samples to support Node.js 18.
- Updated comments on Notification bot template for how to send notifications to an individual person.
- Updated TeamsFx-React SDK to support React 17.
- Updated TeamsFx SDK, relevant bot app templates and samples to use [CloudAdapter](https://learn.microsoft.com/javascript/api/botbuilder/cloudadapter).
- Updated app templates to use TeamsJS SDK v2.7.1.

Bug fixes:

- Fxied an issue where login page will pop up even after signed into Microsoft 365 account from Teams Toolkit.
- Fixed an issue where it may fail to run ngrok because of whitespaces in path.

## 4.2.0 - Dec 20, 2022

New features:

- Use the new Dashboard tab template to create a personal and collaborative canvas in Teams. Learn more about this new template in the [documentation](https://aka.ms/teamsfx-dashboard-app).

  ![image](https://user-images.githubusercontent.com/11220663/205831064-26f94d61-7bb7-4bc1-8678-c52be6cb27b3.png)

Enhancements:

- Step-by-step guidance in Visual Studio Code to assist you while creating a free Microsoft 365 testing tenant.
- Scaffolds that use SharePoint Framework (SPFx) have been updated to use [SharePoint Framework version 1.16](https://learn.microsoft.com/sharepoint/dev/spfx/release-1.16) which introduces the ability to build experiences for Outlook and Office.com (Microsoft 365 app) in addition to Teams.

## 4.1.3 - Nov 17, 2022

Enhancements:

- Bump up version of [@microsoft/teamsfx](https://www.npmjs.com/package/@microsoft/teamsfx) used in templates and samples to `2.0.0` that depends on [@microsoft/teams-js](https://www.npmjs.com/package/@microsoft/teams-js) version `2.0.0` and above. Starting from this version, it enables [Teams app to run in Outlook and Office.](https://learn.microsoft.com/microsoftteams/platform/m365-apps/overview)
- A/B testing on in-product documentation to display tutorial articles inside Visual Studio Code.

## 4.1.2 - Nov 10, 2022

Hotfix version for Teams Toolkit.

- Fixed bugs in Sample Gallery.

## 4.1.1 - Nov 02, 2022

Enhancements:

- Improved experience to obtain a Microsoft Graph client in [@microsoft/teamsfx](https://www.npmjs.com/package/@microsoft/teamsfx) with separated configration interfaces for different authentication flows.
- Included `useTeams` hook in [@microsoft/teamsfx-react](https://www.npmjs.com/package/@microsoft/teamsfx-react) package. Thanks to its original author [Wictor Wilén](https://github.com/wictorwilen) who first developed this hook in [msteams-react-base-component](https://github.com/wictorwilen/msteams-react-base-component) package.
- Upgraded [@microsoft/teamsfx](https://www.npmjs.com/package/@microsoft/teamsfx) to `2.0.0` that's compatible with [@microsoft/teamjs](https://www.npmjs.com/package/@microsoft/teams-js) `2.0.0` and above.
- Added a shortcut to debug and preview your Teams application from left side tree view under `DEVELOPMENT` section.

## 4.1.0 - Oct 17, 2022

New features:

- Use the new Workflow bot template to create sequential workflows where Adaptive Cards can be returned in response to user input. Learn more about this new template in the [documentation](https://aka.ms/teamsfx-card-action-response).
- Configure each step of the debugging experience in Teams Toolkit by customizing the `.vscode/tasks.json` file, which now [surfaces parameters for finer-tuned control](https://aka.ms/teamsfx-debug-tasks).
- Use the new Add SSO command for a simple way to [configure single sign-on for Message extension and Command bot projects](https://aka.ms/teamsfx-add-sso).
- SPFx developers can now combine web parts as a [multi-Tab Teams app](https://learn.microsoft.com/sharepoint/dev/spfx/build-for-teams-me-experience#build-a-multi-tab-personal-teams-app) using the Add Feature menu.

Enhancements:

- You can now send notifications to a specific channel or person with TeamsFx SDK. Checkout [code snippets](https://aka.ms/teamsfx-send-notification#send-notifications-to-a-specific-channel) here.
- We've changed which windows are shown by default when an app starts debugging so you'll no longer see red error messages coming from the browser that are mistaken as errors for your app.

## 4.0.6 - Sep 23, 2022

Incremental version for Teams Toolkit with multiple bugs fix, enhancements and new features.

New sample:

- Stock Update: Keep up to date with the latest stock price in Microsoft Teams

SDK:

- Simplified single sign-on implementation for message extension with TeamsFx SDK. Take a look at the [sample code](https://aka.ms/teamsfx-me-sso-sample).
- Simplified single sign-on implementation for command bot with TeamsFx SDK. Take a look at the [sample code](https://aka.ms/teamsfx-command-bot-sso).

## 4.0.5 - Aug 22, 2022

Incremental version for Teams Toolkit with multiple bugs fix, enhancements and a new feature to allow switch Microsoft 365 tenant and Azure subscription for local debugging and cloud provision.

## 4.0.4 - Aug 09, 2022

Incremental version for Teams Toolkit with multiple bugs fix, enhancements and new additions to sample application gallery.

New samples:

- Proactive Messaging: Save users' conversation reference to send proactive reminder messages using bots.
- One Productivity Hub: View calendar events, to-do tasks and files in Teams tab.

Enhancement:

- Enhanced experience of building Teams tab application with SharePoint Framework v1.15.0

## 4.0.3 - Jul 26, 2022

Incremental version for Teams Toolkit with multiple bugs fix and enhancements.

Enhancement:

- User can now switch Azure account or Azure subscription to provision cloud resource.
- Now support SPFx version to v1.15.0.

## 4.0.2 - Jul 12, 2022

Incremental version for Teams Toolkit with multiple bugs fix and enhancements.

## 4.0.1 - Jun 16, 2022

Incremental version for Teams Toolkit with multiple bugs fix and enhancements.

## 4.0.0 - May 24, 2022

Major version for Teams Toolkit with new features to support more Teams app scenario like notification bot and command bot. What's more, this version adds support to extend Teams app across Microsoft 365 platform like Office.com and Outlook.

New Features:

- User can create more business-oriented Teams app template using Teams Toolkit. For example, user can now create not only bot app but also notification bot or command bot. User can see more options are there to choose when create a new Teams app.
- User can create Teams app that can launch and preview in other Microsoft 365 platform like Office.com and Outlook. The options are offered when user create a new Teams app.
- User can incrementally add features to their Teams app using `Add features` in Teams Toolkit during the development process. For example, adding additional Teams extending capability, adding Azure resources like SQL Database or Azure Function etc., adding Single Sign on or API connections and so on.
- User can preview the Teams manifest file and only deploy the manifest file without deploy the whole project.
- User can customized Microsoft Entra manifest file.
- Add tutorials in the Teams Toolkit, user can find them by typing the command `Teams: View Guided Tutorials` in the command palette (Ctrl+Shift+P).
- A new sample which use Graph Connector get on board to the Sample Gallery. Click `View Samples` in Teams Toolkit to browse Sample Gallery.

Enhancement:

- UI improvement of `Create a new Teams app` and `Start from a sample`.
- UI improvement of the Teams Toolkit menus in the sidebar.
- Optimize and simplify the Sample apps. Improve the experience of Sample apps.
- Improved the experience of Teams Toolkit CLI tool.

## 3.8.0 - Apr 22, 2022

Incremental version for Teams Toolkit with multiple bugs fixes and the following updates:

Enhancement:

- Optimize the configuration schema and manifest template of project created by Teams Toolkit.
- Support to use CodeLens to preview variables value in manifest template file.
- Optimize the In-meeting Sample App in sample gallery, shorten the time to run the sample.
- Improved "Start from a sample" UI, show more information of each sample.

## 3.7.0 - Apr 06, 2022

Incremental version for Teams Toolkit with multiple bugs fixes and the following updates:

New Features:

- Provide multiple entry points of creating testing tenant/accounts to unblock user from Microsoft 365 account issues, like Microsoft 365 account does not have custom app upload permission or user does not have Microsoft 365 account at all. These entry points include an Add(+) button besides ACCOUNTS in sidebar, an new "Create an account" option in `Teams: Accounts` Command and improved Get Started page.

Enhancement:

- Improved SPFx Project scaffolding experience by using Yeoman Generator.

## 3.6.1 - Mar 23, 2022

Hotfix version for Teams Toolkit with multiple bugs fixed.

## 3.6.0 - Mar 21, 2022

Incremental version for Teams Toolkit with multiple bugs fixes and the following updates:

New Features:

- Optimized Get Started page for Teams Toolkit. User can check environment prerequisites from Get started now.
- User can use Teams Toolkit to create workflow automation templates for Github, Azure DevOps and Jenkins.

Enhancement:

- Enhance TeamsFx SDK.

## 3.5.0 - Mar 07, 2022

Incremental version for Teams Toolkit with multiple bugs fixes and the following updates:

New Features:

- New sample app - Teams tab app without SSO.

Enhancement:

- Teams tab app generated from "create a new Teams app" can now use graph toolkit to retrieve user data.

## 3.4.0 - Feb 21, 2022

Incremental version for Teams Toolkit with multiple bugs fixes and the following updates:

Enhancement:

- Improved local debug experience, more light weighted and more graceful.

## 3.3.0 - Feb 07, 2022

Incremental version for Teams Toolkit with multiple bugs fixes and the following updates:

New Features:

- "Add cloud resources" feature now supports adding multiple instances of the same cloud resource type. For example, add multiple instance of SQL DB at the same time.

Enhancement:

- Teams Tab project created by Teams Toolkit now is updated to use Auth Code Flow with PKCE for SPA authentication. You can find more details [here](https://aka.ms/teamsfx-auth-code-flow). Please be noted that Tab project created by Teams Toolkit of this version will not be supported by previous versions of Teams Toolkit.

## 3.2.0 - Jan 10, 2022

Incremental version for Teams Toolkit with multiple bugs fixes and the following updates:

New Features:

- Use Service Principle to login Azure account in CICD template.
- Support building React Tab app by different environment variables for multiple environments.

Enhancement:

- Provide guidance to install development certificate on WSL. See guidance [here](https://github.com/OfficeDev/TeamsFx/blob/dev/docs/fx-core/localdebug-help.md#how-to-manually-install-the-development-certificate-for-windows-subsystem-for-linux-wsl-users)
- Support .NET SDK 6.0.
- Improve the experience to preview manifest file and update manifest file to Developer Portal.
- Improve CICD template by reducing dependency on project metadata file.

## 3.1.1 - Dec 27, 2021

This is a hotfix version.

The Azure App service is upgraded and does not support some older NodeJs versions in some regions any more. This hotfix solves the problem that Azure App service is not working in those regions which does not support older NodeJs versions.

## 3.1.0 - Dec 13, 2021

Incremental version for Teams Toolkit with multiple bugs fixes and the following updates:

New Features:

- Integrate with Azure Key Vault to secure your application secrets at runtime.
- View state file and edit environment configurations from manifest with code lens.

Enhancement:

- Support Node.js 16 and NPM 7 except for SPFx based tab application or projects including Azure Functions.

## 3.0.0 - Nov 29, 2021

Major version for Teams Toolkit with new features to support cloud resources customization, multiple cloud environments, collaborations and some bug fix.

New Features:

- Adopt ARM templates to provision Azure cloud resources, support customization of cloud resources. Refer to [Provision cloud resources](https://aka.ms/provision-doc) for more information.
- Developers can create and manage multiple cloud environments with different customizations for each environment. Refer to [Manage multiple environment](https://aka.ms/multi-env-doc) for more information.
- Developers can collaborate with others on the same project. Refer to [Collaborations in Teams Toolkit](https://aka.ms/collaboration-doc) for more information.
- Support manifest customization for both local and remote environment. Refer to [Customize manifest](https://aka.ms/customize-manifest-doc) for more information.
- Provide flexibility to add cloud resources to your project using ARM template. Refer to [Add cloud resources](https://aka.ms/add-resources-doc) for more information.
- Add more Teams Sample Apps which support local run with no need to manually set up environment.

Enhancement:

- Improve UI: In sample gallery, add time and effort estimation for each sample.
- Improve UI: multiple enhancement to the Tree View. For example, provide documents links in Tree View, and enrich the tooltip descriptions.
- Reduce the required user inputs in order to create new project.
- Enhance the status and messages showed in Teams Toolkit.
- Upgrade samples to adopt new features in Teams Toolkit.

## 2.10.0 - Nov 15, 2021

Incremental version for Teams Toolkit with multiple bugs fixes and the following updates:

New Features:

- Enable developers with the capability to extend Teams apps to run across Microsoft 365, get instructions from our [documentation](https://aka.ms/teamsfx-extend-m365).
- Provide Teams sample apps "Todo List (Works in Teams, Outlook and Office)" and "NPM Search Connector" which can run across Microsoft 365. Users can get an initial experience of running Teams apps in Microsoft 365.

## 2.9.0 - Nov 01, 2021

Incremental version for Teams Toolkit with multiple bugs fixes and the following updates:

Preview Features:

- Enable CI/CD for multiple environments scenario.
- Insider Preview features in 2.8.0 release are still in preview. Refer to [Enable insider preview features](https://github.com/OfficeDev/TeamsFx/wiki/Enable-Preview-Features-in-Teams-Toolkit) for how to use.
- Upgrade existing projects to support preview features, refer to [Upgrade existing project to use latest features](https://github.com/OfficeDev/TeamsFx/wiki/Upgrade-project-to-use-latest-Toolkit-features) for more information.

Enhancement:

- Improve UI: more friendly user experience to create a new Teams project.
- Improve UI: add links to source code for samples.
- Support one-click deployment of SharePoint framework based Teams app.
- Integrate Adaptive Card Studio with previewing and debugging Adaptive Card in VS Code.

## 2.8.0 - Oct 18, 2021

Incremental version for Teams Toolkit with multiple bugs fixes and the following updates:

Preview Features:

- Support management of multiple environments.
- Support project collaborations among multiple developers.
- Improve the experience to customize Azure resource provision using ARM(Azure Resource Manager).

To enable the preview features, refer to the [preview guidance](https://github.com/OfficeDev/TeamsFx/wiki/Enable-Preview-Features-in-Teams-Toolkit)

Enhancement:

- Improve UI: making the Tree View and Command Pallete text consistent.
- UX A/B testing:
  - Your Tree View(sidebar) may include or exclude quick start page.
  - You may or may not be invited to do local debug after create new project.

## 2.7.0 - Sep 17, 2021

Incremental version for Teams Toolkit with multiple bugs fixes and the following updates:

Enhancement:

- Improved version upgrade experience by adding "What is New?" info.
- Simplified welcome view when clicking the Toolkit logo on the sidebar.

## 2.6.0 - Sep 06, 2021

Incremental version for Teams Toolkit with multiple bugs fixed and the following updates：

New Feature:

- Support projects migration from Teams Toolkit V1 to V2. If your Teams projects are created using Teams Toolkit V1, try migrate your project follow the [migration instructions](https://aka.ms/teamsfx-migrate-v1).
- Support local debug experience for Teams Tab/Bot/Messaging extension project migrated from Teams Toolkit V1.
- Check permission to turn-on Teams custom app uploading when user sign-in to Microsoft 365 account. Learn more about [custom app upload permission](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/build-and-test/prepare-your-o365-tenant#enable-custom-teams-apps-and-turn-on-custom-app-uploading).
- (Preview Feature) Support provision cloud resources using Azure Resource Manager. To enable this feature, please follow [instructions](https://github.com/OfficeDev/TeamsFx/wiki/Enable-Preview-Features-in-Teams-Toolkit).

Enhancement:

- UI Enhancement:
  - Hide Azure account login for SharePoint projects.
  - Tree View A/B testing: with non-Teams projects open, you may randomly see either one of two different Tree View layouts in the side bar.
- Optimization of created project folder structure.
- Improved getting start experience of creating new project for Bot.
- Upgrade Sample apps. Enable CI/CD workflow for Todo-List with Azure backend sample. More samples are coming, view them at [Sample GitHub open source repo](https://github.com/OfficeDev/TeamsFx-samples)

## 2.5.0 - Aug 23 2021

Incremental version for Teams Toolkit with multiple bugs fixed and the following updates：

New Feature:

- Enable create project from quick start page.
- Enable report issue from local debug pop-up error message.
- Enable CI/CD workflow with Github Actions. Check [CI/CD instructions](https://github.com/OfficeDev/TeamsFx/tree/dev/docs/cicd) to learn how to use.

Enhancement:

- Update new CLI progress bar.
- Improve some UI experience.
- Add more information in the output error message of debug.

## 2.4.1 - Aug 10 2021

Incremental version for Teams Toolkit with multiple bugs fixed and the following updates：

New Feature:

- Add extension accessiblity for the disabled.
- Add CLI command 'teamsfx preview' to preview SPFx apps.

Enhancement:

- A/B testing for Tree View design. You may randomly see either one of two different Tree View layouts in the side bar.

## 2.3.1 - July 26 2021

Incremental version for Teams Toolkit with multiple bugs fixed and the following updates：

New Feature:

- Add CLI command 'teamsfx preview' to directly preview Teams apps after scaffolding.
- For multiple developers collaborate on one project, Teams Toolkit enables developers to create their own local environment by pressing F5.
- Add encryption to secret data in .userdata file, and support view/edit secret data through VS Code UI and CLI tool.

Enhancement:

- Speed up the installation of npm dependency for debugging experience.

## 2.2.0 - July 12 2021

Incremental version for Teams Toolkit with following updates：

Added:

- UI Run button for debug and preview
- Provide 'Report Issue' button in the pop-up error message to log issues when npm install fails or system error occurs
- A/B testing:  
  Please be noted that in this release we are conducting A/B testing for 'Tree View', the left panel of Teams Toolkit. Thus you may see different layout of Tree View, that is not an issue but totally by design.

## 2.1.1 - June 30 2021

Multiple bug fixings

## 2.1.0 - June 15 2021

Incremental version for Teams Toolkit with the following updates:

Added:

- enable customer survey from toolkit
- add FAQ plus to the samples
- better local debug experience

## 2.0.1 - May 25 2021

2.0.1 is a major version upgrade for Teams Toolkit with many new improvements and supports.

Added:

- new design of UI and command palette list
- new getting started guide, samples and doc link from toolkit
- new sign-in experience for Microsoft 365 and Azure
- new interactive flow for creating new app
- add messaging extension capability to Teams app
- sample codes for in-meeting app.
- support e2e dev experience for tab app hosted by either Azure or SPFx
- support adding backend API (Azure Functions) and SQL DB to tab app
- preview support for tab app, bot app, and messaging extension
- dev environment checking and auto-setup
- cloud provision and deploy for Teams app from treeview UI and command palette
- support simplified auth code and graph client through scaffolding
- Mac/Linux support of new toolkit
- integration with GitHub Codespaces for dev/test

Enhanced:

- improved helloworld app through scaffolding for tab app, bot app, and messaging extension
- imoroved local frontend/backend debugging support for tab app, bot app, and messaging extension
- improved error messages and logging
- improved publish to teams flow

Fixed:

- cumulated bug fixes

## 1.2.3 - April 20 2021

- Add scaffold option for a Teams messaging extension with SSO (Single Sign-on)

## 1.2.2 - April 1 2021

- Increase the timout limit when creating a Microsoft Entra password for a bot registration.

## 1.2.1 - March 15 2021

- Bug fix for env/manifest file automatic replacements not happening for some scaffolds.
- Bug fix for Microsoft Entra app creation where consent should be admin and users and not just admin.

## 1.2.0 - March 2021

- Pull scaffolds from public github repository for always up to date content.
- Updated scaffold selection wizard with more details on what is available.
- More scaffolding and language options.
- Publishing package support when manifest.json is updated locally the zip and cloud configuration in App Studio is created/updated.
- New scaffolds repository https://github.com/OfficeDev/Microsoft-Teams-Samples

## 1.1.8 - 9 Dec 2020

- Microsoft Entra single sign-on bug fixes for Group tabs

## 1.1.4 - 1 Dec 2020

- Microsoft Entra single sign-on documentation updates
- Bug fixes

## 1.1.3 - 24 Nov 2020

- Microsoft Entra single sign-on scaffolding for Tab apps
- Download application publishing package from VS Code
- Bug fixes

## 1.1.1 & 1.1.2

- Bug fixes

## 1.1.0 - 27 Oct 2020

- Onboard to the new VS Code auth
- New F5
- Deprecation of .publish folder - all updates you do through App Studio tooling
- Pre 1.1.0 project upgrade to new project format. We detect if they don't have a teamsAppId in the project...and if they have a .publish folder we create an app registration for them and update the project.
- New bot creation wizard updates
- New project creation flow
- Bug fixes

## 1.0.4 - 3 Aug 2020

Publish to your org catalog from the toolkit

## 1.0.2 - 15 Jul 2020

Bug fixes

## 1.0.1 - 10 Jul 2020

Quality improvements

## 1.0.0 - 7 Jul 2020

Bug fixes

## 0.9.6 - 30 June 2020

Added

- A new bot service instance is automatically created in Azure when a bot/messaging extension project is provisioned and requires one.
- Run your project in the Teams client by hitting F5.

## 0.9.5 - 19 May 2020

Extension released in preview.
