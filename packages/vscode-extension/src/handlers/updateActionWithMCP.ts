// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { err, FxError, ok, Result, Stage, UserError } from "@microsoft/teamsfx-api";
import { getSystemInputs } from "../utils/systemEnvUtils";
import path from "path";
import * as fs from "fs-extra";
import { QuestionNames } from "@microsoft/teamsfx-core";
import * as vscode from "vscode";
import axios from "axios";
import { runCommand } from "./sharedOpts";

export async function updateActionWithMCP(args?: any[]): Promise<Result<any, FxError>> {
  const inputs = getSystemInputs();
  const mcpFile = path.join(inputs.projectPath!, ".vscode", "mcp.json");
  if (!fs.pathExistsSync(mcpFile)) {
    return err(new UserError("da-mcp", "MCPFileNotFound", "MCP file not found"));
  }
  const mcpContent = await fs.readJSON(mcpFile);
  if (!mcpContent || !mcpContent.servers) {
    return err(new UserError("da-mcp", "MCPContentInvalid", "MCP content is invalid"));
  }

  // TODO: support multiple MCP servers
  const mcpNames = Object.keys(mcpContent.servers);
  if (mcpNames.length === 0 || mcpNames.length > 1) {
    return err(new UserError("da-mcp", "MCPServerCountInvalid", "MCP server count is invalid"));
  }

  const mcpName = mcpNames[0];
  const server = mcpContent.servers[mcpName].url;
  inputs[QuestionNames.MCPForDAServerUrl] = server;
  inputs[QuestionNames.MCPForDAServerName] = mcpName;
  const allMcpTools = vscode.lm.tools;
  const tools = allMcpTools
    .filter((tool: any) => (tool.name as string).includes(`mcp_${mcpName}`))
    .map((tool: any) => {
      const index = tool.name.indexOf(mcpName);
      const newName = (tool.name as string).substring((index as number) + mcpName.length + 1);
      return {
        name: newName,
        description: tool.description,
        inputSchema: tool.inputSchema,
        tags: tool.tags,
      };
    });
  if (tools.length === 0) {
    return err(
      new UserError(
        "da-mcp",
        "MCPToolsNotFound",
        "No tools found for the MCP server. Please run the server first."
      )
    );
  }
  inputs[QuestionNames.MCPForDAAvailableTools] = tools;

  let auth: "OAuthPluginVault" | "NoneAuth" = "NoneAuth";
  try {
    const response = await axios.get(server);
  } catch (error) {
    if (error.status == 401) {
      auth = "OAuthPluginVault";
    }
  }
  inputs[QuestionNames.MCPForDAAuth] = auth;
  const result = await runCommand(Stage.updateActionWithMCP, inputs);
  return result;
}
