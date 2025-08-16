# Default Bot

This template contains a simple echo bot for Microsoft Teams using the Microsoft Agents Hosting library for Python.

## Prerequisites

- Python 3.8 or higher
- Microsoft Teams
- [Microsoft 365 Agents Toolkit](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension)

## Getting Started

1. In the Command Palette, select **Microsoft 365 Agents Toolkit: Create a new App** and choose **Default Bot** template
2. Provide your app name and save to a folder
3. Open the app folder in VS Code
4. Press **F5** to start debugging which launches your app in Microsoft Teams using Teams Toolkit

## What's Included

This template includes:

- Echo bot functionality that repeats user messages with a count
- Welcome message for new conversation members
- Command handlers for bot management:
  - `/reset` - Reset conversation state
  - `/count` - Show current message count
  - `/diag` - Show diagnostic information
  - `/state` - Show current conversation state
  - `/runtime` - Show runtime information
- Conversation state management
- Microsoft Agents Hosting library integration

## Project Structure

```
├── .vscode/
│   ├── launch.json          # Launch configurations for debugging
│   ├── settings.json        # VS Code settings
│   └── tasks.json          # VS Code tasks
├── appPackage/             # Teams app manifest
│   ├── color.png           # App icon (color)
│   ├── outline.png         # App icon (outline)
│   └── manifest.json       # Teams app manifest
├── env/                    # Environment files
├── infra/                  # Infrastructure files for deployment
├── src/
│   ├── app.py              # Main application entry point
│   ├── bot.py              # Bot logic and handlers
│   ├── config.py           # Configuration settings
│   └── requirements.txt    # Python dependencies
├── .gitignore
├── .webappignore
├── README.md
└── m365agents.yml          # Project configuration
```

## Key Features

### Echo Functionality
The bot echoes back any message sent to it with a message counter.

### Command Handling
Special commands are handled with dedicated handlers:
- Type `/reset` to clear conversation state
- Type `/count` to see the current message count
- Type `/diag` to view diagnostic information
- Type `/state` to see current conversation state
- Type `/runtime` to view runtime information

### State Management
Conversation state is maintained using the Microsoft Agents Hosting library's built-in state management.

## Customization

You can extend this bot by:

1. Adding new command handlers using the `@teams_bot.message()` decorator
2. Implementing custom activity handlers using `@teams_bot.activity()`
3. Adding custom state properties to the `ConversationState` class
4. Integrating with external APIs or services

## Learn More

- [Microsoft Agents Documentation](https://github.com/microsoft/Agents)
- [Microsoft Teams Platform Documentation](https://docs.microsoft.com/microsoftteams/platform/)
- [Microsoft 365 Agents Toolkit Documentation](https://docs.microsoft.com/microsoftteams/platform/toolkit/)
