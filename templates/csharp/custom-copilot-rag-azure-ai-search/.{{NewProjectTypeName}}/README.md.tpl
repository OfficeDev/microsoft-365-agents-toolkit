# Overview of the Chat With Your Data (Custom Data Source) template

This app template showcases how to build one of the most powerful applications enabled by LLM - sophisticated question-answering (Q&A) chat bots that can answer questions about specific source information right in the Microsoft Teams.
This app template also demonstrates usage of techniques like: 
- [Retrieval Augmented Generation](https://python.langchain.com/docs/use_cases/question_answering/#what-is-rag), or RAG.
- [Teams AI Library](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/teams%20conversational%20ai/teams-conversation-ai-overview)

## Get started with the template

> **Prerequisites**
>
> To run the template in your local dev machine, you will need:
>
> - Prepare your own [Azure AI Search](https://azure.microsoft.com/en-us/products/ai-services/ai-search).
{{#useOpenAI}}
> - Prepare an account with [OpenAI](https://platform.openai.com).
{{/useOpenAI}}
{{#useAzureOpenAI}}
> - Prepare [Azure OpenAI](https://aka.ms/oai/access) resource
{{/useAzureOpenAI}}

## Create your Azure AI Search document index 
**Before running or debugging your bot, please follow these steps to create your document index in Auzre AI Search.**

> This app template provides script `Indexer.ps1` to help create document index. You can change the instructions and settings in the script to customize the document index.
{{#useOpenAI}}
1. Ensure your OpenAI and Azure AI search settings filled in `appsettings.TestTool.json`.
    ```
    "OpenAI": {
      "ApiKey": "<your-openai-api-key>",
      "EmbeddingModel": "<your-openai-embedding-model>"
    },
    "Azure": {
      "AISearchApiKey": "<your-azure-ai-search-api-key>",
      "AISearchEndpoint": "<your-azure-ai-search-endpoint>"
    }
    ```
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. Ensure your Azure OpenAI and Azure AI search settings filled in `appsettings.TestTool.json`.
    ```
    "Azure": {
      "OpenAIApiKey": "<your-azure-openai-api-key>",
      "OpenAIEndpoint": "<your-azure-openai-endpoint>",
      "OpenAIDeploymentName": "<your-azure-openai-deployment-name>",
      "OpenAIEmbeddingDeploymentName": "<your-azure-openai-embedding-deployment-name>",
      "AISearchApiKey": "<your-azure-ai-search-api-key>",
      "AISearchEndpoint": "<your-azure-ai-search-endpoint>"
    }
    ```
{{/useAzureOpenAI}}
1. Open PowerShell, change the current working directory to this project root and run command `. ./Indexer.ps1 -run create`.
   ```
   > . ./Indexer.ps1 -run create
   ```
1. Once you're done using the sample it's good practice to delete the index. You can do so with the `. ./Indexer.ps1 -run delete`.

### Debug bot app in Microsoft 365 Agents Playground
1. Create your Azure AI Search document index as mentioned above.
1. Set `Startup Item` as `Microsoft 365 Agents Playground (browser)`.
![image](https://github.com/user-attachments/assets/41121a95-5c13-4bda-8f69-3b970a4e8f78)
1. Press F5, or select the Debug > Start Debugging menu in Visual Studio.
1. In Microsoft 365 Agents Playground from the launched browser, type and send anything to your bot to trigger a response.

**Congratulations**! You are running an application that can now interact with users in Microsoft 365 Agents Playground:

![RAG Bot](https://github.com/OfficeDev/TeamsFx/assets/13211513/f56e7602-a5d3-436a-ae01-78546d61717d)

### Debug bot app in Teams Web Client
1. Create your Azure AI Search document index as mentioned above.
{{#useOpenAI}}
1. Ensure your OpenAI and Azure AI search settings filled in `env/.env.local.user`.
    ```
    SECRET_OPENAI_API_KEY="<your-openai-api-key>"
    OPENAI_EMBEDDING_MODEL="<your-openai-embedding-model>"
    SECRET_AI_SEARCH_API_KEY="<your-azure-ai-search-api-key>"
    AI_SEARCH_ENDPOINT="<your-azure-ai-search-endpoint>"
    ```
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. Ensure your Azure OpenAI and Azure AI search settings filled in `env/.env.local.user`.
    ```
    SECRET_AZURE_OPENAI_API_KEY="<your-azure-openai-api-key>"
    AZURE_OPENAI_ENDPOINT="<your-azure-openai-endpoint>"
    AZURE_OPENAI_DEPLOYMENT_NAME="<your-azure-openai-deployment-name>"
    AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME="<your-azure-openai-embedding-deployment-name>"
    SECRET_AI_SEARCH_API_KEY="<your-azure-ai-search-api-key>"
    AI_SEARCH_ENDPOINT="<your-azure-ai-search-endpoint>"
    ```
{{/useAzureOpenAI}}
1. Create your Azure AI Search document index as mentioned above.
2. In the debug dropdown menu, select Dev Tunnels > Create A Tunnel (set authentication type to Public) or select an existing public dev tunnel.
3. Right-click the '{{NewProjectTypeName}}' project in Solution Explorer and select **Microsoft 365 Agents Toolkit > Select Microsoft 365 Account**
4. Sign in to Microsoft 365 Agents Toolkit with a **Microsoft 365 work or school account**
5. Set `Startup Item` as `Microsoft Teams (browser)`.
6. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
7. In the opened web browser, select Add button to install the app in Teams
8. In the chat bar, type and send anything to your bot to trigger a response.

> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

## Extend the template

- Follow [Build a Basic AI Chatbot in Teams](https://aka.ms/teamsfx-basic-ai-chatbot) to extend the template with more AI capabilities.
- Understand more about [build your own data ingestion](https://aka.ms/teamsfx-rag-bot#build-your-own-data-ingestion).
- Understand more about [Azure AI Search as data source](https://aka.ms/teamsfx-rag-bot#azure-ai-search-as-data-source).

## Additional information and references

- [Microsoft 365 Agents Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Microsoft 365 Agents Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)

## Report an issue

Select Visual Studio > Help > Send Feedback > Report a Problem. 
Or, you can create an issue directly in our GitHub repository: 
https://github.com/OfficeDev/TeamsFx/issues.
