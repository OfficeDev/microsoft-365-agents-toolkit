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
- Apps and Agents Samples and Templates Code Snippets Retriever
- Troubleshooting Retriever

## Prerequisites

1. The Microsoft 365 Agents Toolkit MCP Server requires Node.js to install and run the server. If you don't have it installed, follow the instructions [here](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm).
2. Install either the stable or Insiders release of VS Code:
   * [💫 Stable release](https://code.visualstudio.com/download)
   * [🔮 Insiders release](https://code.visualstudio.com/insiders)
3. Install the [GitHub Copilot](https://marketplace.visualstudio.com/items?itemName=GitHub.copilot) and [GitHub Copilot Chat](https://marketplace.visualstudio.com/items?itemName=GitHub.copilot-chat) extensions

#### Manual Install

For a step-by-step guide to install the Microsoft 365 Agents Toolkit MCP Server, follow these instructions:

- Add `.vscode/mcp.json`:
    ```json
    {
        "servers": {
            "M365AgentsToolkit Server": {
                "command": "npx",
                "args": [
                    "-y",
                    "@microsoft/m365agentstoolkit-mcp@latest",
                    "server",
                    "start"
                ]
            }
        }
    }
    ```

#### Install From Command Palette
- [![Install with NPX in VS Code](https://img.shields.io/badge/VS_Code-Install_all-0098FF?style=flat-square&logo=visualstudiocode&logoColor=white)](https://insiders.vscode.dev/redirect/mcp/install?name=M365AgentsToolkit%20Server&config=%7B%22command%22%3A%22npx%22%2C%22args%22%3A%5B%22-y%22%2C%22@microsoft%2Fm365agentstoolkit-mcp%40latest%22%2C%22server%22%2C%22start%22%5D%7D) [![Install with NPX in VS Code Insiders](https://img.shields.io/badge/VS_Code_Insiders-Install_all-24bfa5?style=flat-square&logo=visualstudiocode&logoColor=white)](https://insiders.vscode.dev/redirect/mcp/install?name=M365AgentsToolkit%20Server&config=%7B%22command%22%3A%22npx%22%2C%22args%22%3A%5B%22-y%22%2C%22@microsoft%2Fm365agentstoolkit-mcp%40latest%22%2C%22server%22%2C%22start%22%5D%7D&quality=insiders)

#### List The Tools

- Open GitHub Copilot in VS Code and [switch to Agent mode](https://code.visualstudio.com/docs/copilot/chat/chat-agent-mode)

- Click `refresh` on the tools list.

#### For Visual Studio

Manual configuration required, please follow: [Visual Studio MCP Official Guide](https://learn.microsoft.com/en-us/visualstudio/ide/mcp-servers?view=vs-2022)

## License
This project is licensed under the [MIT License](LICENSE).