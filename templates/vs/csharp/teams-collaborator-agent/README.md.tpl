# Teams Collaborator Agent

An intelligent agent for Microsoft Teams that helps teams collaborate more effectively by:

- **Summarizing Conversations**: Generate concise summaries of chat discussions
- **Tracking Action Items**: Automatically identify and track action items from conversations
- **Searching Chat History**: Search through conversation history with natural language

## Features

### Conversation Summarization
Ask the agent to summarize a chat conversation, and it will provide a concise overview of key points discussed.

### Action Item Extraction
The agent can identify action items, tasks, and commitments from conversations and present them in an organized format.

### Conversation Search
Search through your chat history using natural language queries to quickly find relevant information.

## Getting Started

### Prerequisites
- .NET 8.0 SDK or later
- Microsoft 365 Agents Toolkit for Visual Studio
- Azure OpenAI or OpenAI API key

### Configuration

1. Update `appsettings.Development.json` with your credentials:
   - `Teams.ClientId`: Your Bot ID
   - `Teams.ClientSecret`: Your Bot Password
   - `OpenAI.ApiKey` or `Azure.OpenAIApiKey`: Your AI model credentials
   - `Storage.Type`: Choose "sqlite" or "sqlserver"
   - `Storage.ConnectionString`: Database connection string

2. For local development, the agent uses SQLite by default. For production, configure SQL Server.

### Running Locally

1. Press F5 in Visual Studio to start debugging
2. The agent will be available in Teams or Bot Framework Emulator
3. Start a conversation and @mention the agent to interact with it

## Architecture

The agent uses:
- **Microsoft.SemanticKernel** for AI orchestration
- **Microsoft.Teams.Apps** for Teams integration
- **SQLite/SQL Server** for conversation storage
- **Azure OpenAI** or **OpenAI** for language models

## Project Structure

```
├── Agent/              # Agent manager and routing logic
├── Capabilities/       # Individual capability implementations
│   ├── Summarizer/    # Conversation summarization
│   ├── ActionItems/   # Action item extraction
│   └── Search/        # Conversation search
├── Storage/           # Database and conversation storage
├── appPackage/        # Teams app manifest and icons
├── infra/             # Azure deployment templates
└── Program.cs         # Application entry point
```

## Deployment

Deploy to Azure using the provided Bicep templates in the `infra/` directory.

## Learn More

- [Microsoft 365 Agents Toolkit Documentation](https://aka.ms/m365-agents-toolkit)
- [Semantic Kernel Documentation](https://learn.microsoft.com/semantic-kernel)
- [Teams AI Library](https://learn.microsoft.com/microsoftteams/platform/bots/how-to/teams-conversational-ai/teams-conversation-ai-overview)
