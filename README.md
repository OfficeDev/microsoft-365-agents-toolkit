# Microsoft 365 Agents Toolkit

[![DotNet SDK CI workflow](https://github.com/OfficeDev/TeamsFx/actions/workflows/dotnetsdk-ci.yml/badge.svg)](https://github.com/OfficeDev/TeamsFx/actions/workflows/dotnetsdk-ci.yml)
[![Function Extension CI workflow](https://github.com/OfficeDev/TeamsFx/actions/workflows/FunctionExtensionCI.yml/badge.svg)](https://github.com/OfficeDev/TeamsFx/actions/workflows/FunctionExtensionCI.yml)
[![CodeQL](https://github.com/OfficeDev/TeamsFx/actions/workflows/codeql-analysis.yml/badge.svg)](https://github.com/OfficeDev/TeamsFx/actions/workflows/codeql-analysis.yml)
[![codecov](https://codecov.io/gh/OfficeDev/teams-toolkit/branch/dev/graph/badge.svg?token=QQX8WVOEC3)](https://codecov.io/gh/OfficeDev/teams-toolkit)

Microsoft 365 Agents Toolkit for Visual Studio, Visual Studio Code, and Command Line Interface (CLI) are tools for building agents or apps for Microsoft 365 Copilot, Microsoft Teams and Microsoft 365 easy and fast. Whether you are new or a seasoned developer, Microsoft 365 Agents Toolkit is the best way for you to start creating, building, debugging, testing, and deploying.

<img width="2405" alt="DarkBg_Slogan_250429" src="https://github.com/user-attachments/assets/84fa9339-cc14-43fc-bef2-22fc8513e57c" />


Microsoft 365 Agents Toolkit provides support for the end-to-end development journey, including:

- Support for all Microsoft 365 platform extensibility surfaces, including Copilot for Microsoft 365, Agent/chatbot, tabs and message extensions for Microsoft Teams as well as Office Add-ins.
- Seamless integration with [Microsoft Agents SDK](https://learn.microsoft.com/en-us/microsoft-365/agents-sdk/agents-sdk-overview?tabs=csharp) to build self-hosted agents.
- Seamless integration with [Teams AI Library](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/teams%20conversational%20ai/conversation-ai-quick-start?pivots=qs-javascript) to build intelligent chatbot in Microsoft Teams with ease.
- Integrations with the tools, languages, and frameworks you know and love.
- Scaffolds for getting started fast with the extensibility interfaces and common scenarios such as Agents for Microsoft 365 Copilot, intelligent chat bot for multiple platforms.
- Rapid iteration with full stack debugging, hot reload, secure tunneling and [Microsoft 365 Agents Playground](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/debug-your-teams-app-test-tool?tabs=vscode%2Cclijs).
- Simplified SSO authentication.
- Integrated support for hosting, data storage, and serverless functions.
- CI/CD actions for GitHub and Azure DevOps to deliver apps with confidence.

## Get Started

Pick your preferred tool to get started:

- For JavaScript and TypeScript developers, install [Microsoft 365 Agents Toolkit for Visual Studio Code](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals#install-teams-toolkit-for-visual-studio-code).
- For .NET developers, install [Microsoft 365 Agents Toolkit for Visual Studio](https://docs.microsoft.com/microsoftteams/platform/toolkit/visual-studio-overview#install-teams-toolkit-for-visual-studio).
- For command line users, install [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli#get-started): `npm install -g @microsoft/teamsapp-cli`

Visit [Microsoft 365 Agents Toolkit documentation](https://aka.ms/teamsfx-docs) to get started with building agents and apps.

## Roadmap

Microsoft 365 Agents Toolkit for Visual Studio, Visual Studio Code, and Command Line Interface (CLI) will be updated regularly with new features and bug fixes to continuously improve end-to-end Teams development experience. Visit our [Changelog](https://github.com/OfficeDev/teams-toolkit/blob/dev/packages/vscode-extension/CHANGELOG.md) and [Prerelease Changelog](https://github.com/OfficeDev/teams-toolkit/blob/dev/packages/vscode-extension/PRERELEASE.md) to see what's available for you now and see [product roadmap](https://aka.ms/teamsfx-roadmap) to find out what's coming.

## Support Policy

Microsoft 365 Agents Toolkit products will follow [Modern Lifecycle Policy](https://docs.microsoft.com/lifecycle/policies/modern) and extended support as described in our [lifecycle and support document](https://aka.ms/teamsfx-support).

## Feedback

- Ask a question on [Stack Overflow](https://stackoverflow.com/questions/tagged/teams-toolkit)
- [Request a new feature](https://github.com/OfficeDev/TeamsFx/issues/new?assignees=&labels=&template=feature_request.md&title=)
- [File an issue](https://github.com/OfficeDev/TeamsFx/issues/new?assignees=&labels=&template=bug_report.md&title=)
- Send an email to ttkfeedback@microsoft.com to chat with the product team
- Report security issues and bugs to the Microsoft Security Response Center (MSRC) via secure@microsoft.com. Further information can be found in the [Security TechCenter](https://www.microsoft.com/msrc/faqs-report-an-issue?rtc=1).

## Repository

This repository contains the following packages:
| Package | Description |
| ----------- | ----------- |
| **Microsoft 365 Agents Toolkit for Visual Studio Code** [packages/vscode-extension](https://github.com/OfficeDev/TeamsFx/tree/main/packages/vscode-extension) | Teams Toolkit for Visual Studio Code enables you to scaffold, run, debug, and deploy custom Teams apps directly from Visual Studio Code. It provides all the features of the Teams Toolkit CLI tool integrated into the IDE, as well as easy access to more samples, docs and tools. |
| **Microsoft 365 Agents Toolkit CLI** [packages/cli](https://github.com/OfficeDev/TeamsFx/tree/main/packages/cli) | Whether you prefer keyboard-centric developer operations, or you are automating your CI/CD pipeline, the Teams Toolkit CLI tool offers the same features as the IDE extensions. |
| **SDK** [packages/sdk](https://github.com/OfficeDev/TeamsFx/tree/main/packages/sdk) | The main code library encapsulating simple authentication for both client and server-side code tailored for Teams developers. |
| **API** [packages/api](https://github.com/OfficeDev/TeamsFx/tree/main/packages/api) | The API package is a collection of contracts supported by the IDE Extensions and CLI. It enables developers to write plugins to extend TeamsFx with new capabilities. |
| **Core** [packages/fx-core](https://github.com/OfficeDev/TeamsFx/tree/main/packages/fx-core) | The Core package centralizes implementation of capabilities shared by the IDE Extensions and the CLI. |
| **Azure Functions Support** [packages/function-extension](https://github.com/OfficeDev/TeamsFx/tree/main/packages/function-extension) | Teams Toolkit helps developers include server-side code in their Teams application backed by [Azure Functions](https://docs.microsoft.com/azure/azure-functions/). This plugin adds support to simplify the integration of an authentication-aware Azure Function into your Teams app. |
| **Spec Parser** [packages/spec-parser](https://github.com/OfficeDev/TeamsFx/tree/main/packages/spec-parser) | Teams Toolkit automates the process of generating [API-based message extensions](https://learn.microsoft.com/en-us/microsoftteams/platform/messaging-extensions/api-based-overview) and Adaptive Cards by parsing the OpenAPI description document. |

## Contributions

This project welcomes contributions and suggestions. Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit [Contributor License Agreement](https://cla.opensource.microsoft.com).

When you submit a pull request, a CLA bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., status check, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

- Download our latest daily Builds [here](https://github.com/OfficeDev/TeamsFx/releases)
- Check out our [contribution](https://github.com/OfficeDev/TeamsFx/blob/main/CONTRIBUTING.md) page for more information

## Telemetry

Microsoft 365 Agents Toolkit collects usage data and sends it to Microsoft to help improve our products and services. Read our [Privacy Statement](https://go.microsoft.com/fwlink/?LinkId=521839) and [Data Collection Notice](https://docs.opensource.microsoft.com/content/releasing/telemetry.html) to learn more. Learn more in our [FAQ](https://code.visualstudio.com/docs/supporting/faq#_how-to-disable-telemetry-reporting).

## Trademarks

This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft
trademarks or logos is subject to and must follow
[Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/legal/intellectualproperty/trademarks/usage/general).
Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship.
Any use of third-party trademarks or logos are subject to those third-party's policies.

## Code of Conduct

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
