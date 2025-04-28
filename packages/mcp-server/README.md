# Microsoft 365 Agents Toolkit MCP Server

The Microsoft 365 Agents Toolkit MCP Server is a [Model Context Protocol (MCP)](https://modelcontextprotocol.io/introduction)
server that provides a seamless connection between AI agents and developers for building apps and agents for Microsoft 365 and Microsoft 365 Copilot.

## Overview

### What can you do with it?

M365 Agents Toolkit MCP Server is designed to help you: 
- Build and deploy AI agents for Microsoft 365
- Integrate with Microsoft 365 Copilot features
- Access and manage app and agent templates
- Troubleshoot common issues effectively

## Currently Supported Tools
- Schema Fetcher for:
    - App Manifest
    - Declarative Agent Manifest
    - API Plugin Manifest
- Microsoft 365 and Microsoft 365 Copilot Knowledge Retriever
- Apps and Agents Samples and Templates Retriever
- Troubleshooting Retriever

## Getting Started

The Microsoft 365 Agents Toolkit MCP Server requires Node.js to install and run the server. If you don't have it installed, follow the instructions [here](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).

### Prerequisites

1. Install either the stable or Insiders release of VS Code:
   * [💫 Stable release](https://code.visualstudio.com/download)
   * [🔮 Insiders release](https://code.visualstudio.com/insiders)
2. Install the [GitHub Copilot](https://marketplace.visualstudio.com/items?itemName=GitHub.copilot) and [GitHub Copilot Chat](https://marketplace.visualstudio.com/items?itemName=GitHub.copilot-chat) extensions
3. Open VS Code in an empty folder

### Setup

#### Personal Access Token

[Create a GitHub Personal Access Token](https://docs.github.com/en/authentication/keeping-your-account-and-data-secure/managing-your-personal-access-tokens) with appropriate permissions:
   - Go to [Personal access tokens](https://github.com/settings/tokens) (in GitHub Settings > Developer settings)
   - Select which repositories you'd like this token to have access to (Public, All, or Select)
   - Generate token
   - Copy the generated token

#### Manual Install

For a step-by-step guide to install the Microsoft 365 Agents Toolkit MCP Server, follow these instructions:

- Add `.vscode/mcp.json`:
    ```json
    {
        "servers": {
            "M365AgentsToolkit MCP Server": {
                "command": "npx",
                "args": [
                    "-y",
                    "@microsoft/m365agentstoolkit-mcp"
                ],
                "env": {
                    "GITHUB_PERSONAL_ACCESS_TOKEN": "<YOUR_TOKEN>"
                }
            }
        }
    }
    ```

## License
This project is licensed under the [MIT License](LICENSE).