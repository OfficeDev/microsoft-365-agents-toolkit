# Collaborator Agent for Microsoft Teams

This intelligent collaboration assistant is built with the [Teams AI Library v2](https://aka.ms/teamsai-v2), and showcases how to create a sophisticated bot that can analyze conversations, manage tasks, and search through chat history using advanced AI capabilities and natural language processing.

This agent can listen to all messages in a group chat (even without being @mentioned) using RSC (Resource Specific Control) permissions defined in [App Manifest](appPackage/manifest.json). For more details, see the documentation [RSC Documentation](https://staticsint.teams.cdn.office.net/evergreen-assets/safelinks/2/atp-safelinks.html).

## Key Features

- 📋 **Intelligent Summarization** - Analyze conversations and provide structured summaries with proper participant attribution and topic identification
- ✅ **Action Items** - Automatically identify and create action items from team discussions with smart assignment
- 🔍 **Conversation Search** - Search through chat history using natural language queries with time-based filtering and deep linking to original messages

## Agent Architecture

The Collaborator agent uses a sophisticated multi-capability architecture:

- **Manager**: Coordinates between specialized capabilities and handles natural language time parsing
- **Summarizer**: Analyzes conversation content and provides structured summaries
- **Action Items**: Identifies tasks, manages assignments, and tracks completion
- **Search**: Performs semantic search across conversation history with citation support
- **Context Management**: Global message context handling for concurrent request support

## Running the Sample

### Prerequisites

- [.NET 9.0 SDK](https://dotnet.microsoft.com/download)
- [Visual Studio 2022](https://visualstudio.microsoft.com/) 17.14 or later
- [Microsoft 365 Agents Toolkit for Visual Studio](https://aka.ms/install-m365-agents-toolkit-vs)
- A Microsoft Teams account with the ability to upload custom apps
- Azure OpenAI resource with GPT-4 deployment

### Environment Variables

Update the configuration in `appsettings.Development.json`:
  - `Azure:OpenAIApiKey`: Your Azure OpenAI API key
  - `Azure:OpenAIEndpoint`: Your Azure OpenAI endpoint URL
  - `Azure:OpenAIDeploymentName`: Your GPT-4 model deployment name
  - `Storage:Type`: Storage type (sqlite or sqlserver)
  - `Storage:ConnectionString`: Database connection string

### Running the Bot

1. Open the solution in Visual Studio 2022
2. Press F5 to start debugging
3. The agent will launch in Microsoft Teams or Microsoft 365 Agents Playground
4. @mention the bot in any conversation to start using its capabilities!

#### Sample Questions

You can ask the Collaborator agent questions like:

**Summarization:**
- "@Collaborator summarize yesterday's discussion"
- "@Collaborator what were the main topics from last week?"
- "@Collaborator give me an overview of recent messages"

**Action Items:**
- "@Collaborator find action items from the past 3 days"
- "@Collaborator create a task to review the proposal by Friday"
- "@Collaborator what tasks are assigned to me?"

**Search:**
- "@Collaborator find messages about the project deadline"
- "@Collaborator search for conversations between Alice and Bob"
- "@Collaborator locate discussions from this morning about the budget"

## Deployment

The agent can be deployed to Azure App Service for production use:
- Host your app in Azure by [provision cloud resources](https://learn.microsoft.com/microsoftteams/platform/toolkit/provision) and [deploy the code to cloud](https://learn.microsoft.com/microsoftteams/platform/toolkit/deploy)
- Azure SQL Database is recommended for production environments. Update `Storage:Type` to `sqlserver` and provide the connection string.
