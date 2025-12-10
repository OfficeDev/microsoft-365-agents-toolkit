# Overview of the Travel Agent template

This template demonstrates how to build an intelligent travel agent using the Microsoft 365 Agents Toolkit. The agent provides comprehensive travel assistance by answering travel-related questions, helping users understand company travel policies, and finding flights and hotels that comply with organizational guidelines.

The app template is built using the Microsoft 365 Agents SDK and Agent Framework, which provides the capabilities to build AI-based applications.

## Quick Start

**Prerequisites**
> To run the Weather Agent template in your local dev machine, you will need:
>
{{#useOpenAI}}
> - an account with [OpenAI](https://platform.openai.com).
{{/useOpenAI}}
{{#useAzureOpenAI}}
> - [Azure OpenAI](https://aka.ms/oai/access) resource
{{/useAzureOpenAI}}

### Debug agent in Teams Web Client

{{#useOpenAI}}
1. Ensure your OpenAI API Key is filled in `env/.env.local.user`.
    ```
    SECRET_OPENAI_API_KEY="<your-openai-api-key>"
    ```
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. Ensure your Azure OpenAI settings are filled in `env/.env.local.user`.
    ```
    SECRET_AZURE_OPENAI_API_KEY="<your-azure-openai-api-key>"
    AZURE_OPENAI_ENDPOINT="<your-azure-openai-endpoint>"
    AZURE_OPENAI_DEPLOYMENT_NAME="<your-azure-openai-deployment-name>"
    ```
{{/useAzureOpenAI}}
1. In the debug dropdown menu, select Dev Tunnels > Create A Tunnel (set authentication type to Public) or select an existing public dev tunnel.
1. Right-click the '{{NewProjectTypeName}}' project in Solution Explorer and select **Microsoft 365 Agents Toolkit > Select Microsoft 365 Account**
1. Sign in to Microsoft 365 Agents Toolkit with a **Microsoft 365 work or school account**
1. Set `Startup Item` as `Microsoft Teams (browser)`.
1. Press F5, or select Debug > Start Debugging menu in Visual Studio to start your app
</br>![image](https://raw.githubusercontent.com/OfficeDev/TeamsFx/dev/docs/images/visualstudio/debug/debug-button.png)
1. In the opened web browser, select Add button to install the app in Teams
1. In the chat bar, type and send anything to your agent to trigger a response.

> For local debugging using Microsoft 365 Agents Toolkit CLI, you need to do some extra steps described in [Set up your Microsoft 365 Agents Toolkit CLI for local debugging](https://aka.ms/teamsfx-cli-debugging).

### (Optional) Enable Full Capabilities with Microsoft 365 Retrieval API

The following steps are **optional** and only needed if you want to enable the full capabilities of the Travel Agent, including accessing company travel documents and policies stored in SharePoint or OneDrive for Business through the Microsoft 365 Retrieval API.

#### Step 1: Clean Up Previous Debug Configuration

If you have already debugged the agent with basic functionality, you need to clean up the generated configuration files before enabling SSO:

1. Open `{{NewProjectTypeName}}/env/.env.local` and remove all generated values except keep the file structure
{{#useOpenAI}}
1. Open `{{NewProjectTypeName}}/env/.env.local.user` and remove all values except the OpenAI configuration:
    ```
    SECRET_OPENAI_API_KEY="<your-openai-api-key>"
    ```
{{/useOpenAI}}
{{#useAzureOpenAI}}
1. Open `{{NewProjectTypeName}}/env/.env.local.user` and remove all values except the Azure OpenAI configuration:
   ```
   SECRET_AZURE_OPENAI_API_KEY="<your-azure-openai-api-key>"
   AZURE_OPENAI_ENDPOINT="<your-azure-openai-endpoint>"
   AZURE_OPENAI_DEPLOYMENT_NAME="<your-azure-openai-deployment-name>"
   ```
{{/useAzureOpenAI}}

> **Note**: This step ensures a clean state for provisioning with SSO enabled. Skip this step if you haven't debugged the agent yet.

#### Step 2: Upload Sample Documents

1. Navigate to the `{{ProjectName}}/SampleDocuments` folder
1. Upload all sample documents to your **OneDrive for Business** following the steps in the ["Prepare the grounding data"](https://learn.microsoft.com/en-us/training/modules/copilot-declarative-agents-build-your-first/5-exercise-custom-knowledge) section of the tutorial
1. Documents uploaded to OneDrive for Business are available instantly for the Microsoft 365 Retrieval API

> **Note**: SharePoint is also an option for document storage, but it requires several hours for document indexing to complete. If you choose SharePoint, expect significant delays before the documents become available for retrieval. OneDrive for Business is recommended for immediate availability.

#### Step 3: Configure Retrieval Plugin

1. Open `{{ProjectName}}/Bot/Plugins/RetrievalPlugin.cs`
1. Review the [Retrieval API documentation](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/ai-services/retrieval/copilotroot-retrieval)
1. Replace the `DataSource` and `FilterExpression` with your actual configuration. Example values are as below:
   - For **OneDrive for Business** (recommended):
     - `DataSource`: `RetrievalDataSource.OneDriveBusiness`
     - `FilterExpression`: `"(path:\"https://{tenant}-my.sharepoint.com/personal/{user}/Documents/{foldername}\")"`
   - For **SharePoint**:
     - `DataSource`: `RetrievalDataSource.SharePoint`
     - `FilterExpression`: `"(path:\"https://{tenant}.sharepoint.com/\")"`

#### Step 4: Enable SSO in Provisioning Configuration

Open `{{NewProjectTypeName}}/m365agents.local.yml` and replace the following section:

```yaml
- uses: botFramework/create
  with:
    botId: ${{AAD_APP_CLIENT_ID}}
    name: {{appName}}
    messagingEndpoint: ${{BOT_ENDPOINT}}/api/messages
    description: ""
    channels:
      - name: msteams
- uses: script
  with:
    run:
      echo "::set-teamsfx-env BOT_ID=${{AAD_APP_CLIENT_ID}}";
```

With this:

```yaml
- uses: arm/deploy
  with:
    subscriptionId: ${{AZURE_SUBSCRIPTION_ID}}
    resourceGroupName: ${{AZURE_RESOURCE_GROUP_NAME}}
    templates:
      - path: ./infra/azure.bicep
        parameters: ./infra/azure.parameters.local.json
        deploymentName: Create-resources-for-bot-local
    bicepCliVersion: v0.9.1
```

#### Step 5: Enable SSO in Bot Code

Open `{{ProjectName}}/Bot/{{SafeProjectName}}Bot.cs` and replace this line in the constructor:

```csharp
OnActivity(ActivityTypes.Message, MessageActivityAsync, rank: RouteRank.Last);
```

With:

```csharp
OnActivity(ActivityTypes.Message, MessageActivityAsync, rank: RouteRank.Last, autoSignInHandlers: ["graph"]);
```

#### Step 6: Grant Admin Consent

1. After completing the debug steps in Teams Web Client, navigate to the [Microsoft Entra ID portal](https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/Overview)
1. Search for the application using the `AAD_APP_CLIENT_ID` value from `{{NewProjectTypeName}}/env/.env.local`
1. Navigate to the **API permissions** tab
1. Click **Grant admin consent**

## Additional information and references
- [Microsoft 365 Agents SDK](https://github.com/microsoft/Agents)
- [Microsoft 365 Agents Toolkit Documentations](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Microsoft 365 Agents Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)

## Learn more

New to app development or Microsoft 365 Agents Toolkit? Learn more about app manifests, deploying to the cloud, and more in the documentation 
at https://aka.ms/teams-toolkit-vs-docs.

## Report an issue

Select Visual Studio > Help > Send Feedback > Report a Problem. 
Or, you can create an issue directly in our GitHub repository: 
https://github.com/OfficeDev/TeamsFx/issues.
