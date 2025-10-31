// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { exec } from "child_process";
import { promisify } from "util";

export interface ODRServer {
  name: string;
  display_name: string;
  description: string;
  version: string;
  identifier: string;
  tools: ODRTool[];
}

export interface ODRTool {
  name: string;
  description: string;
  inputSchema: any;
  outputSchema?: any;
}

export class ODRProvider {
  /**
   * Parse the output of 'odr list' command
   * @param jsonOutput The JSON output from 'odr list' command
   * @returns Array of parsed ODR servers with their tools
   */
  static parseODRListOutput(jsonOutput: any): ODRServer[] {
    const servers: ODRServer[] = [];

    if (!jsonOutput.servers || !Array.isArray(jsonOutput.servers)) {
      return servers;
    }

    for (const server of jsonOutput.servers) {
      const manifest =
        server._meta?.["io.modelcontextprotocol.registry/publisher-provided"]?.[
          "com.microsoft.windows"
        ]?.manifest;
      const staticResponses = manifest?._meta?.["com.microsoft.windows"]?.static_responses;
      const toolsList = staticResponses?.["tools/list"]?.tools || [];

      servers.push({
        name: server.name,
        display_name: manifest?.display_name || server.name,
        description: server.description || "",
        version: server.version || "1.0.0",
        identifier: server.packages?.[0]?.identifier || "",
        tools: toolsList.map((tool: any) => ({
          name: tool.name,
          description: tool.description || "",
          inputSchema: tool.inputSchema,
          outputSchema: tool.outputSchema,
        })),
      });
    }

    return servers;
  }

  /**
   * Get dummy data for testing (until odr.exe integration is ready)
   * This uses sample data from actual Windows ODR servers
   */
  static getDummyServers(): ODRServer[] {
    return [
      {
        name: "applaunch-mcp-server",
        display_name: "App Launch MCP Server",
        description: "MCP server for launching apps.",
        version: "1.0.1",
        identifier:
          "MicrosoftWindows.Client.Core_cw5n1h2txyewy_com.microsoft.windows.ai.mcpServer_applaunch-mcp-server",
        tools: [
          {
            name: "launch_application",
            description:
              "Launches the specified application using its AppId. Use GetInstalledApplications first to see available applications and their AppIds.",
            inputSchema: {
              type: "object",
              properties: {
                appId: {
                  type: "string",
                  description: "The AppId of the application to launch",
                },
              },
              required: ["appId"],
            },
            outputSchema: {
              type: "object",
              properties: {
                result: {
                  type: "string",
                },
              },
              required: ["result"],
            },
          },
          {
            name: "get_installed_applications",
            description: "Get all installed applications for the user in the system.",
            inputSchema: {
              type: "object",
              properties: {},
            },
            outputSchema: {
              type: "object",
              properties: {
                result: {
                  type: "array",
                },
              },
              required: ["result"],
            },
          },
        ],
      },
      {
        name: "systeminfo-mcp-server",
        display_name: "System Info MCP Server",
        description: "MCP server for getting system information.",
        version: "1.0.1",
        identifier:
          "MicrosoftWindows.Client.Core_cw5n1h2txyewy_com.microsoft.windows.ai.mcpServer_systeminfo-mcp-server",
        tools: [
          {
            name: "get_system_overview",
            description:
              "Get comprehensive system information including CPU, memory, disk, and uptime for troubleshooting",
            inputSchema: {
              type: "object",
              properties: {},
            },
          },
          {
            name: "get_performance_metrics",
            description:
              "Get the latest snapshot of system performance metrics including CPU usage, memory, disk activity",
            inputSchema: {
              type: "object",
              properties: {},
            },
          },
          {
            name: "get_local_users",
            description: "Get local user accounts and their properties for security auditing",
            inputSchema: {
              type: "object",
              properties: {},
            },
          },
        ],
      },
      {
        name: "file-mcp-server",
        display_name: "Files Connector",
        description: "MCP server for performing file operations and file search.",
        version: "1.0.2",
        identifier:
          "MicrosoftWindows.Client.Core_cw5n1h2txyewy_com.microsoft.windows.ai.mcpServer_file-mcp-server",
        tools: [
          {
            name: "get_file_details",
            description: "Get file details such as name, extension, size, creation time.",
            inputSchema: {
              type: "object",
              properties: {
                path: {
                  type: "string",
                  description: "File path",
                },
              },
              required: ["path"],
            },
          },
          {
            name: "create_directory",
            description: "Create a directory on the given path",
            inputSchema: {
              type: "object",
              properties: {
                path: {
                  type: "string",
                  description: "Directory path",
                },
                directoryName: {
                  type: "string",
                  description: "Directory name",
                },
              },
              required: ["path", "directoryName"],
            },
          },
          {
            name: "search_files",
            description:
              "Search for files or directories on all available folders by name or extension",
            inputSchema: {
              type: "object",
              properties: {},
            },
          },
        ],
      },
      {
        name: "windowing-mcp-server",
        display_name: "Windowing MCP Server",
        description: "MCP server for performing windowing operations.",
        version: "1.0.1",
        identifier:
          "MicrosoftWindows.Client.Core_cw5n1h2txyewy_com.microsoft.windows.ai.mcpServer_windowing-mcp-server",
        tools: [
          {
            name: "get_window_list",
            description: "Get a list of windows that can be interacted with using other tools",
            inputSchema: {
              type: "object",
              properties: {},
            },
          },
          {
            name: "snap_window",
            description: "Snap a list of 2/3/4 windows",
            inputSchema: {
              type: "object",
              properties: {
                request: {
                  type: "object",
                },
              },
              required: ["request"],
            },
          },
          {
            name: "close_window",
            description: "Close a window",
            inputSchema: {
              type: "object",
              properties: {
                windowId: {
                  type: "string",
                },
              },
              required: ["windowId"],
            },
          },
        ],
      },
    ];
  }

  /**
   * List all available local MCP servers
   * TODO: Implement actual 'odr list' command execution when ODR provider is ready
   * For now, returns dummy data for development and testing
   */
  static async listServers(): Promise<ODRServer[]> {
    const execAsync = promisify(exec);
    const execOutcome = async () => {
      try {
        const { stdout } = await execAsync("odr list");
        const jsonOutput = JSON.parse(stdout);
        return ODRProvider.parseODRListOutput(jsonOutput);
      } catch (error) {
        console.error("Error executing odr list:", error);
        return [];
      }
    };

    return execOutcome().then((output) => {
      return output;
    });
  }
}
