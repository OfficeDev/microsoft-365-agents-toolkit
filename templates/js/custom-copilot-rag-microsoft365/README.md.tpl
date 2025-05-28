# Overview of the Chat With Your Data (Using Microsoft 365 Data) template

This app template showcases how to build one of the most powerful applications enabled by LLM - sophisticated question-answering (Q&A) chat bots that can answer questions about specific source information right in the Microsoft Teams.
This app template also demonstrates usage of techniques like: 
- [Retrieval Augmented Generation](https://python.langchain.com/docs/use_cases/question_answering/#what-is-rag), or RAG.
- [Microsoft Graph Search API](https://learn.microsoft.com/graph/search-concept-overview)
- [Teams AI Library](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/teams%20conversational%20ai/teams-conversation-ai-overview)

## Get started with the template

> **Prerequisites**
>
> To run the template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 18, 20, 22
> - A Microsoft 365 tenant in which you have permission to upload apps. You can get a free Microsoft 365 developer tenant by joining the [Microsoft 365 developer program](https://developer.microsoft.com/en-us/microsoft-365/dev-program).
> - [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli).
{{#useOpenAI}}
> - An account with [OpenAI](https://platform.openai.com/).
{{/useOpenAI}}
{{#useAzureOpenAI}}
> - Prepare your own [Azure OpenAI](https://aka.ms/oai/access) resource.
{{/useAzureOpenAI}}

> [!TIP]
> You can adjust this template to use data from a Microsoft Copilot connector. Follow the steps in the [CodeTour](https://marketplace.visualstudio.com/items?itemName=vsls-contrib.codetour) included in the project to apply the necessary changes. To use data from a Microsoft Copilot connector, you need a Copilot connector deployed to your tenant. For testing, we recommend using the [Ingest custom API data using TypeScript, Node.js and Microsoft 365 Agents Toolkit for Visual Studio Code](https://adoption.microsoft.com/sample-solution-gallery/sample/pnp-graph-connector-nodejs-typescript-food-catalog) sample.

> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

1. First, select the Microsoft 365 Agents Toolkit icon on the left in the VS Code toolbar.
{{#useOpenAI}}
1. In file *env/.env.local.user*, fill in your OpenAI key `SECRET_OPENAI_API_KEY=<your-key>`.
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. In file *env/.env.local.user*, fill in your Azure OpenAI key `SECRET_AZURE_OPENAI_API_KEY=<your-key>`, endpoint `AZURE_OPENAI_ENDPOINT=<your-endpoint>` and deployment name `AZURE_OPENAI_DEPLOYMENT_NAME=<your-deployment>`.
{{/useAzureOpenAI}}
1. Microsoft Graph Search API is available for searching SharePoint content, thus you just need to ensure your document in *src/data/\*.txt* is [uploaded to SharePoint / OneDrive](https://support.microsoft.com/office/upload-files-and-folders-to-a-library-da549fb1-1fcb-4167-87d0-4693e93cb7a0), no extra data ingestion required.
1. Press F5 to start debugging which launches your app in Teams using a web browser. Select `Debug in Teams (Edge)` or `Debug in Teams (Chrome)`.
1. When Teams launches in the browser, select the Add button in the dialog to install your app to Teams.
1. You can send any message to get a response from the agent.

![M365 RAG Bot](https://github.com/OfficeDev/TeamsFx/assets/13211513/c2fff68c-53ce-445a-a101-97f0c127b825)

## What's included in the template

| Folder       | Contents                                            |
| - | - |
| `.vscode`    | VSCode files for debugging                          |
| `appPackage` | Templates for the application manifest        |
| `env`        | Environment files                                   |
| `infra`      | Templates for provisioning Azure resources          |
| `src`        | The source code for the application                 |

The following files can be customized and demonstrate an example implementation to get you started.

| File                                 | Contents                                           |
| - | - |
|`src/index.js`| Sets up the agent app server.|
|`src/adapter.js`| Sets up the agent adapter.|
|`src/config.js`| Defines the environment variables.|
|`src/prompts/chat/skprompt.txt`| Defines the prompt.|
|`src/prompts/chat/config.json`| Configures the prompt.|
|`src/app/app.js`| Handles business logics for the RAG agent.|
|`src/app/m365DataSource.js`| Defines the m365 data source.|
|`src/data/*.txt`| Raw text data sources.|
|`src/public/*.html`| Auth start page and an auth end page to be used by the user sign in flow.|

The following are Microsoft 365 Agents Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Microsoft 365 Agents Toolkit works.

| File                                 | Contents                                           |
| - | - |
|`m365agents.yml`|This is the main Microsoft 365 Agents Toolkit project file. The project file defines two primary things:  Properties and configuration Stage definitions. |
|`m365agents.local.yml`|This overrides `m365agents.yml` with actions that enable local execution and debugging.|

## Extend the template

- Follow [Build a Basic AI Chatbot in Teams](https://aka.ms/teamsfx-basic-ai-chatbot) to extend the template with more AI capabilities.
- Understand more about [how to add additional APIs](https://aka.ms/teamsfx-rag-bot#add-more-api-for-custom-api-as-data-source).

## Additional information and references

- [Microsoft 365 Agents Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Microsoft 365 Agents Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)

## Known issue

- This template doesn't work in Microsoft 365 Copilot currently.