# Microsoft 365 Agents Toolkit - an evolution of Teams Toolkit

## What is Microsoft 365 Agents Toolkit?

The Microsoft 365 Agents Toolkit, [an evolution of Teams Toolkit](https://aka.ms/M365AgentsToolkit), is designed to help developers create and deploy agents or apps for multiple Microsoft 365 platforms including Microsoft 365 Copilot, Microsoft Teams and Microsoft 365. It can significantly ease your development life by providing integrated [Microsoft 365](https://docs.microsoft.com/microsoftteams/platform/concepts/build-and-test/prepare-your-o365-tenant) identity, cloud storage access, data from [Microsoft Graph](https://docs.microsoft.com/graph/teams-concept-overview), and other services in [Azure](https://docs.microsoft.com/microsoftteams/platform/build-your-first-app/build-bot) with a "zero-configuration" approach.

![ATK animation](https://github.com/user-attachments/assets/96e34da7-06a5-4632-9243-db823e9edb51)

## Extensive Templates Spectrum

No matter you are building custom extensions for Microsoft 365 Copilot, Microsoft Teams or Microsoft 365. It is easy to find the right template to start with in Microsoft 365 Agents Toolkit.

### Build for Microsoft 365 Copilot

You can bring your custom knowledge, skills, and process automation into Microsoft 365 Copilot for your specific needs. You can build your own agents for specific tasks, such as retrieving information, summarizing data, and taking actions like sending emails or updating records. There are two types of agents:

[Declarative Agents](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/overview-declarative-agent) are leveraging the models and AI services built in Copilot. You tailor Copilot by declaring instructions, actions and retrieving knowledge from specified APIs to meet your unique business needs.
[Custom Engine Agents](https://learn.microsoft.com/en-us/microsoft-365-copilot/extensibility/overview-custom-engine-agent) are specialized agent experiences built on any large language model (LLM) and tailored for a specific domain or workflows. You can build your custom orchestrators, foundation models, and you can customize your agent business logic.

### Build for Microsoft Teams

Microsoft Teams platform provides [diverse extensible features](https://learn.microsoft.com/en-us/microsoftteams/platform/overview-solution) for developers to customize for their own business scenarios. 

[Teams Bot](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/overview) allows users to interact with your web service through text, interactive cards, and task modules. There are chat bots, notification bots, and various bots for different scenarios. And you can evolve Teams bots to [conversational agents for Teams](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/teams-conversational-ai/ai-ux) with the [Teams AI library](https://learn.microsoft.com/en-us/microsoftteams/platform/bots/how-to/teams-conversational-ai/teams-conversation-ai-overview?tabs=javascript%2Cjavascript1). Agents in Teams can understand and respond to user chat in natural language to handle QnA interactions or engage in commands/actions interacting with back-end services.
[Teams Tabs](https://learn.microsoft.com/en-us/microsoftteams/platform/tabs/what-are-tabs?tabs=personal) are webpages embedded directly within Teams. They can be integrated as part of a channel inside a team, in group chats, or as personal apps designed for individual users. 

[Teams Message extensions](https://learn.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/what-are-messaging-extensions?tabs=desktop) allow users to interact with your web service directly from the Teams client. They support searching for information or triggering actions in external systems. The outcome can be delivered as a richly formatted card within the conversation.

### Build for Office and Outlook.

You can build your custom [Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins) that extend and interact with content in Excel, Word, PowerPoint, OneNote and Outlook. You only need to build one solution and it can run in Office across all these apps.
## Getting started

Open Microsoft 365 Agents Toolkit to create a new agent/app and start coding!
<img width="1111" alt="image" src="https://github.com/user-attachments/assets/add124dc-44db-4c10-b84c-243af89aa0f3" />

Follow the [Get Started](https://aka.ms/teamsfx-build-first-app) instructions to start with a new project.

You can easily discover all available commands by activating Microsoft 365 Agents Toolkit in the sidebar:
<img width="406" alt="image" src="https://github.com/user-attachments/assets/dcde0d16-5757-4e42-9aa5-e731d76ebb0e" />

or search keyword " > Microsoft 365 Agents" in VS Code. (Invoke search pad with the shortcut `ctrl+shift+P`)
<img width="590" alt="image" src="https://github.com/user-attachments/assets/507a3ef4-a740-48b2-afbc-b3b0ee32ad1a" />


### Useful links

[Microsoft 365 Agents Toolkit Documentation](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)

[Microsoft 365 Agents Toolkit CLI](https://www.npmjs.com/package/@microsoft/m365agentstoolkit-cli) and [CLI documentation](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/teams-toolkit-cli?pivots=version-three)

[Create new Agent/App](https://docs.microsoft.com/microsoftteams/platform/toolkit/create-new-project)

[Preview and customize the manifest file](https://docs.microsoft.com/microsoftteams/platform/toolkit/teamsfx-preview-and-customize-app-manifest)

[Debug and preview](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/debug-overview)

[Deploy to Azure](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/deploy)

[Upload app and publish to organization](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/publish)

[Publish app to Teams Store](https://learn.microsoft.com/en-us/microsoftteams/platform/concepts/deploy-and-publish/appsource/publish)


## Prerequisites

Verify you have the right prerequisites for building apps and install some recommended development tools. [Read more details](https://docs.microsoft.com/microsoftteams/platform/build-your-first-app/build-first-app-overview).

<table>
    <tr>
        <td><img src="https://raw.githubusercontent.com/OfficeDev/TeamsFx/main/packages/vscode-extension/img/landingPage_nodejs.png"></td>
        <td><h3>Node.js</h3>As a fundamental runtime context for app, Node.js is required. If you develop SPFx Tab app, please install v18.x.</td>
    </tr>
    <tr>
        <td><img src="https://raw.githubusercontent.com/OfficeDev/TeamsFx/main/packages/vscode-extension/img/landingPage_m365.png"></td>
        <td><h3>Microsoft 365</h3> A Microsoft 365 work or organization account is required where your app is registered and deployed.</td>
    </tr>
    <tr>
        <td><img src="https://raw.githubusercontent.com/OfficeDev/TeamsFx/main/packages/vscode-extension/img/landingPage_azure.png"></td>
        <td><h3>Azure</h3> The Microsoft 365 Agents Toolkit may require an Azure account and subscription to create the Azure resources. </td>
    </tr>
</table>

> Don’t have a Microsoft 365 account to experience building app? Sign up for [Microsoft Developer Program](https://developer.microsoft.com/microsoft-365/dev-program), which allows you to have a testing tenant with preconfigured permissions.


## Explore Code Samples

Explore our [samples](https://github.com/OfficeDev/TeamsFx-Samples) to help you quickly get started with the basic app concepts and code structures.

## Contributing

There are many ways in which you can participate in the project, for example:

- [Download our latest builds](https://github.com/OfficeDev/TeamsFx/releases).
- [Submit bugs and feature requests](https://github.com/OfficeDev/TeamsFx/issues), and help us verify as they are checked in 
- Review [source code changes](https://github.com/OfficeDev/TeamsFx/pulls)
- Review the [documentation](CONTRIBUTING.md) and make pull requests for anything from typos to new content

## Reporting security issues

Give security researchers information on how to privately report security vulnerabilities found in your open-source project. See more details [Reporting security issues](https://docs.opensource.microsoft.com/content/releasing/security.html).

## Telemetry

The software may collect information about you and your use of the software and send it to Microsoft. Microsoft may use this information to provide services and improve our products and services. You may turn off the telemetry as described in the repository. There are also some features in the software that may enable you and Microsoft to collect data from users of your applications. If you use these features, you must comply with applicable law, including providing appropriate notices to users of your applications together with a copy of Microsoft's privacy statement. Our privacy statement is located at [Microsoft Privacy Statement](https://go.microsoft.com/fwlink/?LinkID=824704). You can learn more about data collection and use in the help documentation and our privacy statement. Your use of the software operates as your consent to these practices.

### Telemetry Configuration

Telemetry collection is on by default. To opt out, please set the `telemetry.enableTelemetry` setting to `false`. Learn more in our [FAQ](https://code.visualstudio.com/docs/supporting/faq#_how-to-disable-telemetry-reporting).

## Code of conduct

See [Microsoft Open Source code of conduct](https://opensource.microsoft.com/codeofconduct).

## Trademark

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/legal/intellectualproperty/trademarks/usage/general). Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.

## License

Copyright (c) Microsoft Corporation. All rights reserved.

Licensed under the [MIT](LICENSE) license.
