# Getting Started with Travel Agent

This sample demonstrates how to build an intelligent travel agent using the Microsoft 365 Agents Toolkit. The agent provides comprehensive travel assistance by answering travel-related questions, helping users understand company travel policies, and finding flights and hotels that comply with organizational guidelines.

The Travel Agent leverages Azure OpenAI and the Microsoft 365 Retrieval API to access company travel documents and policies stored in SharePoint or OneDrive for Business, providing contextual and policy-compliant travel recommendations directly within Microsoft Teams.

![Sample response from agent](./assets/sampleResponse.png)

## This sample illustrates

- How to build an intelligent agent using Microsoft 365 Agents Toolkit and .NET 9.0
- Integration with Azure OpenAI for natural language processing and conversation handling
- Use of Microsoft 365 Retrieval API to access and search company documents
- Implementation of custom plugins for specialized travel assistance
- Authentication and authorization with Microsoft 365 services
- Deployment and debugging of agents in Microsoft Teams

## Prerequisites to use this sample

- [.NET 9.0 SDK](https://dotnet.microsoft.com/download/dotnet/9.0)
- [Visual Studio 2026](https://visualstudio.microsoft.com/)
- [Microsoft 365 Agents Toolkit for Visual Studio](https://aka.ms/teams-toolkit-vs)
- A Microsoft 365 tenant with administrative permissions to:
  - Create Azure Active Directory applications
  - Grant admin consent for API permissions
  - Upload files to SharePoint or OneDrive for Business
- **Azure OpenAI resource**: [Access Azure OpenAI](https://aka.ms/oai/access)
- **Azure subscription**: Active subscription with appropriate permissions

## Minimal path to awesome

### Run the app locally

#### Step 1: Configure Azure OpenAI Settings

1. Open the project in Visual Studio 2026
2. Update your Azure OpenAI configuration in the `env/.env.local.user` file:

```env
SECRET_AZURE_OPENAI_API_KEY="<your-azure-openai-api-key>"
AZURE_OPENAI_ENDPOINT="<your-azure-openai-endpoint>"
AZURE_OPENAI_DEPLOYMENT_NAME="<your-azure-openai-deployment-name>"
```

#### Step 2: Upload Sample Documents

1. Navigate to the `TravelAgent/SampleDocuments` folder
2. Upload all sample documents to your **OneDrive for Business** following the steps in the ["Prepare the grounding data"](https://learn.microsoft.com/en-us/training/modules/copilot-declarative-agents-build-your-first/5-exercise-custom-knowledge) section of the tutorial
3. Documents uploaded to OneDrive for Business are available instantly for the Microsoft 365 Retrieval API

> **Note**: SharePoint is also an option for document storage, but it requires several hours for document indexing to complete. If you choose SharePoint, expect significant delays before the documents become available for retrieval. OneDrive for Business is recommended for immediate availability.

#### Step 3: Configure Retrieval Plugin

1. Open `TravelAgent/Bot/Plugins/RetrievalPlugin.cs`
2. Review the [Retrieval API documentation](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/ai-services/retrieval/copilotroot-retrieval)
3. Replace the `DataSource` and `FilterExpression` with your actual configuration. Example values are as below:
   - For **OneDrive for Business** (recommended):
     - `DataSource`: `RetrievalDataSource.OneDriveBusiness`
     - `FilterExpression`: `"(path:\"https://{tenant}-my.sharepoint.com/personal/{user}/Documents/{foldername}\")"`
   - For **SharePoint**:
     - `DataSource`: `RetrievalDataSource.SharePoint`
     - `FilterExpression`: `"(path:\"https://{tenant}.sharepoint.com/\")"`

#### Step 4: Set Up Development Tunnel

1. In Visual Studio, open the debug dropdown menu
2. Select **Dev Tunnels** > **Create A Tunnel**
3. Set authentication type to **Public**, or select an existing public dev tunnel

#### Step 5: Configure Microsoft 365 Account

1. Right-click the **M365Agent** project in Solution Explorer
2. Select **Microsoft 365 Agents Toolkit** > **Select Microsoft 365 Account**
3. Sign in with your **Microsoft 365 work or school account**

#### Step 6: Start Debugging

1. Set the **Startup Item** to **Microsoft Teams (browser)**
2. Press **F5** or select **Debug** > **Start Debugging** to launch the application
3. Follow the instructions for provisioning necessary infrastructure when prompted

#### Step 7: Grant Admin Consent

1. After deployment completes, navigate to the [Microsoft Entra ID portal](https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/Overview)
2. Search for the application using the `AAD_APP_CLIENT_ID` value from `M365Agent/env/.env.local`
3. Navigate to the **API permissions** tab
4. Click **Grant admin consent**

### Deploy the app to Azure

>Here are the instructions to run the sample in **Visual Studio 2026**.

1. Clone the repo to your local workspace or directly download the source code
2. Download [Visual Studio 2026](https://visualstudio.microsoft.com/) and install [Microsoft 365 Agents Toolkit for Visual Studio](https://aka.ms/teams-toolkit-vs).
3. Open the project in Visual Studio
4. Open **env/.env.dev.user** file, set value for `SECRET_AZURE_OPENAI_API_KEY`, `AZURE_OPENAI_ENDPOINT` and `AZURE_OPENAI_DEPLOYMENT_NAME`
5. Open **m365agents.yml** file, go to `Project` menu and select `Microsoft 365 Agents Toolkit` -> `Provision in the Cloud`
6. Once provision is completed, open `Project` menu and select `Microsoft 365 Agents Toolkit` -> `Deploy to the Cloud`
7. Grant admin consent for application with `AAD_APP_CLIENT_ID` in `M365Agent/env/.env.dev` same as step 7 in local debug.

### Preview the app in Teams

#### Step 1: Install and Test the Agent

1. After the previous debugging step completes, a Teams web browser window will pop up automatically
2. In the Teams web browser window, select **Add** to install the application
3. In the chat interface, type any message to interact with your travel agent
4. Try asking questions about travel policies, flight bookings, or hotel recommendations

> **Note**: For local debugging using the Microsoft 365 Agents Toolkit CLI, refer to the [CLI debugging setup guide](https://aka.ms/teamsfx-cli-debugging).

## Advanced usage of this sample

### Customizing the Travel Agent

You can extend the Travel Agent's capabilities by:

1. **Adding Custom Plugins**: Create additional plugins in the `TravelAgent/Bot/Plugins/` directory to handle specific travel scenarios
2. **Modifying Company Policies**: Update the sample documents in `TravelAgent/SampleDocuments/` with your organization's actual travel policies
3. **Integrating External APIs**: Replace the mocked methods in `TravelAgent/Bot/Plugins/DataPlugin.cs` with actual connections to airline, hotel, or car rental APIs for real-time booking capabilities
4. **Enhancing Conversation Flow**: Modify the conversation logic to handle more complex travel planning scenarios

### Configuration Options

- **Azure OpenAI Models**: The sample supports different Azure OpenAI deployment models. Configure in the environment settings.
- **Retrieval Sources**: Customize the data sources by modifying the Retrieval Plugin configuration to point to specific SharePoint sites or document libraries.
- **Authentication Scopes**: Adjust the required permissions in the application manifest based on your specific use case.

## Version History

|Date| Author| Comments|
|---|---|---|
|Oct 31, 2025| qinzhouxu | Initial release with .NET 9.0 support and Microsoft 365 Agents SDK|

## Additional information and references

- [Microsoft 365 Agents SDK](https://github.com/microsoft/Agents)
- [Microsoft 365 Agents Toolkit Documentation](https://docs.microsoft.com/microsoftteams/platform/toolkit/teams-toolkit-fundamentals)
- [Microsoft 365 Agents Toolkit CLI](https://aka.ms/teamsfx-toolkit-cli)
- [Microsoft 365 Agents Toolkit Samples](https://github.com/OfficeDev/TeamsFx-Samples)
- [Retrieval API Documentation](https://learn.microsoft.com/microsoft-365-copilot/extensibility/api/ai-services/retrieval/copilotroot-retrieval)
- [Agent Framework](https://learn.microsoft.com/agent-framework/)

## Feedback

We really appreciate your feedback! If you encounter any issue or error, please report issues to us following the [Supporting Guide](https://github.com/OfficeDev/TeamsFx-Samples/blob/dev/SUPPORT.md). Meanwhile you can make [recording](https://aka.ms/teamsfx-record) of your journey with our product, they really make the product better. Thank you!
