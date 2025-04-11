# Overview of the Chat With Your Data (Using Azure AI Search) template

This app template showcases how to build one of the most powerful applications enabled by LLM - sophisticated question-answering (Q&A) chat bots that can answer questions about specific source information right in the Microsoft Teams.
This app template also demonstrates usage of techniques like: 
- [Retrieval Augmented Generation](https://python.langchain.com/docs/use_cases/question_answering/#what-is-rag), or RAG.
- [Azure AI Search](https://learn.microsoft.com/azure/search/search-what-is-azure-search)
- [Teams AI Library](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/teams%20conversational%20ai/teams-conversation-ai-overview)

## Get started with the template

> **Prerequisites**
>
> To run the template in your local dev machine, you will need:
>
> - [Node.js](https://nodejs.org/), supported versions: 18, 20, 22
> - [Microsoft 365 Agents Toolkit Visual Studio Code Extension](https://aka.ms/teams-toolkit) version 5.0.0 and higher or [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
{{#useOpenAI}}
> - An account with [OpenAI](https://platform.openai.com/) and [Azure AI Search](https://azure.microsoft.com/en-us/products/ai-services/ai-search).
{{/useOpenAI}}
{{#useAzureOpenAI}}
> - Prepare your own [Azure OpenAI](https://aka.ms/oai/access) resource and [Azure AI Search](https://azure.microsoft.com/en-us/products/ai-services/ai-search).
{{/useAzureOpenAI}}

> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

1. First, select the Microsoft 365 Agents Toolkit icon on the left in the VS Code toolbar.
{{#useOpenAI}}
1. In file *env/.env.testtool.user*, fill in your OpenAI key `SECRET_OPENAI_API_KEY=<your-key>`. And fill in your Azure AI search key `SECRET_AZURE_SEARCH_KEY=<your-ai-search-key>` and endpoint `AZURE_SEARCH_ENDPOINT=<your-ai-search-endpoint>`.
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. In file *env/.env.testtool.user*, fill in your Azure OpenAI key `SECRET_AZURE_OPENAI_API_KEY=<your-key>`, endpoint `AZURE_OPENAI_ENDPOINT=<your-endpoint>`, deployment name `AZURE_OPENAI_DEPLOYMENT_NAME=<your-deployment>`, and embedding deployment name `AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME=<your-embedding-deployment>`. And fill in your Azure AI search key `SECRET_AZURE_SEARCH_KEY=<your-ai-search-key>` and endpoint `AZURE_SEARCH_ENDPOINT=<your-ai-search-endpoint>`.
{{/useAzureOpenAI}}
{{#useOpenAI}}
1. Do `npm install` and `npm run indexer:create -- <your-ai-search-key> <your-openai-api-key>` to create the my documents index. Once you're done using the sample it's good practice to delete the index. You can do so with the `npm run indexer:delete -- <your-ai-search-key>` command.
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. Do `npm install` and `npm run indexer:create -- <your-ai-search-key> <your-azure-openai-api-key>` to create the my documents index. Once you're done using the sample it's good practice to delete the index. You can do so with the `npm run indexer:delete -- <your-ai-search-key>` command.
{{/useAzureOpenAI}}
1. Press F5 to start debugging which launches your app in Microsoft 365 Agents Playground using a web browser. Select `Debug in Microsoft 365 Agents Playground`.
1. You can send any message to get a response from the bot.

**Congratulations**! You are running an application that can now interact with users in Microsoft 365 Agents Playground:

![AI Search Bot](https://github.com/OfficeDev/TeamsFx/assets/13211513/f56e7602-a5d3-436a-ae01-78546d61717d)

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
|`src/index.ts`| Sets up the bot app server.|
|`src/adapter.ts`| Sets up the bot adapter.|
|`src/config.ts`| Defines the environment variables.|
|`src/prompts/chat/skprompt.txt`| Defines the prompt.|
|`src/prompts/chat/config.json`| Configures the prompt.|
|`src/app/app.ts`| Handles business logics for the RAG bot.|
|`src/app/azureAISearchDataSource.ts`| Defines the Azure AI search data source.|
|`src/indexers/data/*.md`| Raw text data sources.|
|`src/indexers/utils.ts`| Basic index tools. |
|`src/indexers/setup.ts`| A script to create index and upload documents. |
|`src/indexers/delete.ts`| A script to delete index and documents. |

The following are Microsoft 365 Agents Toolkit specific project files. You can [visit a complete guide on Github](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5-Guide#overview) to understand how Microsoft 365 Agents Toolkit works.

| File                                 | Contents                                           |
| - | - |
|`m365agents.yml`|This is the main Microsoft 365 Agents Toolkit project file. The project file defines two primary things:  Properties and configuration Stage definitions. |
|`teamsapp.local.yml`|This overrides `m365agents.yml` with actions that enable local execution and debugging.|
|`teamsapp.testtool.yml`| This overrides `m365agents.yml` with actions that enable local execution and debugging in Microsoft 365 Agents Playground.|

## Extend the template

- Follow [Build a Basic AI Chatbot in Teams](https://aka.ms/teamsfx-basic-ai-chatbot) to extend the template with more AI capabilities.
- Follow [Build a RAG Bot in Teams](https://aka.ms/teamsfx-rag-bot) to extend the template with more RAG capabilities.
- Understand more about [Azure AI Search as data source](https://aka.ms/teamsfx-rag-bot#azure-ai-search-as-data-source).

## Additional information and references

- [Microsoft 365 Agents Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Microsoft 365 Agents Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)