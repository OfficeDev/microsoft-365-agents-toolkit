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
  packageFamily: string;
  command: string;
  args: string[];
}

export interface ODRTool {
  name: string;
  description: string;
  inputSchema: any;
  outputSchema?: any;
  _meta?: any;
}

export class ODRProvider {
  /**
   * Parse the output of 'odr list' command
   * @param jsonOutput The JSON output from 'odr list' command
   * @returns Array of parsed ODR servers with their tools
   */
  static parseODRListOutput(jsonOutput: any): ODRServer[] {
    const servers: ODRServer[] = [];

    if (!jsonOutput || !jsonOutput.servers || !Array.isArray(jsonOutput.servers)) {
      return servers;
    }

    for (const server of jsonOutput.servers) {
      const manifest =
        server._meta?.["io.modelcontextprotocol.registry/publisher-provided"]?.[
          "com.microsoft.windows"
        ]?.manifest;
      const staticResponses = manifest?._meta?.["com.microsoft.windows"]?.static_responses;
      const toolsList = staticResponses?.["tools/list"]?.tools || [];
      const packageFamily = manifest?._meta?.["com.microsoft.windows"]?.package_family_name;
      const mcpConfig = manifest?.server?.mcp_config;

      if (packageFamily && mcpConfig) {
        servers.push({
          name: server.name,
          packageFamily: packageFamily,
          display_name: manifest?.display_name || server.name,
          description: server.description || "",
          version: server.version || "1.0.0",
          identifier: server.packages?.[0]?.identifier || "",
          command: mcpConfig.command || "",
          args: mcpConfig.args || [],
          tools: toolsList.map((tool: any) => ({
            name: tool.name,
            description: tool.description || "",
            inputSchema: tool.inputSchema,
            outputSchema: tool.outputSchema,
          })),
        });
      }
    }

    return servers;
  }

  /**
   * List all available local MCP servers from Windows ODR.
   * This is used during project creation to show available servers to the user.
   * @returns Array of ODR servers. Returns empty array if:
   *   - Not on Windows platform
   *   - ODR command is not available (not installed)
   *   - ODR command fails or returns invalid output
   */
  static async listServers(): Promise<ODRServer[]> {
    if (process.platform !== "win32") {
      return [];
    }

    const execAsync = promisify(exec);
    try {
      const { stdout } = await execAsync("odr list");

      if (!stdout) {
        return [];
      }

      const jsonOutput = JSON.parse(stdout);
      return ODRProvider.parseODRListOutput(jsonOutput);
    } catch (error) {
      console.error("Error executing odr list:", error);
      return [];
    }
  }
}
