// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  err,
  FxError,
  ok,
  Result,
  SingleSelectConfig,
  Stage,
  UserError,
} from "@microsoft/teamsfx-api";
import { getSystemInputs } from "../utils/systemEnvUtils";
import path from "path";
import * as fs from "fs-extra";
import { QuestionNames } from "@microsoft/teamsfx-core";
import * as vscode from "vscode";
import axios from "axios";
import { runCommand } from "./sharedOpts";
import { VS_CODE_UI } from "../qm/vsc_ui";

export async function updateActionWithMCP(args?: any[]): Promise<Result<any, FxError>> {
  const inputs = getSystemInputs();
  let mcpName = args && args.length > 0 ? args[0].serverName : undefined;
  let server = args && args.length > 0 ? args[0].serverConfig.url : undefined;
  if (!mcpName && !server) {
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
    if (mcpNames.length === 0) {
      return err(
        new UserError("da-mcp", "MCPServerNotFound", "No MCP server found in the MCP file")
      );
    }
    if (mcpNames.length === 1) {
      mcpName = mcpNames[0];
      server = mcpContent.servers[mcpName].url;
    } else {
      const mcpNameSelection: SingleSelectConfig = {
        name: "mcpName",
        title: "Select MCP Server",
        options: mcpNames.map((name) => ({
          id: name,
          label: name,
          detail: mcpContent.servers[name].url,
        })),
      };
      const result = await VS_CODE_UI.selectOption(mcpNameSelection);
      if (result.isErr()) {
        return err(result.error);
      }
      mcpName = result.value.result;
      server = mcpContent.servers[mcpName].url;
    }
  } else if (!mcpName || !server) {
    return err(
      new UserError("da-mcp", "MCPNameOrServerUrlMissing", "MCP name or server URL is missing")
    );
  }

  inputs[QuestionNames.MCPForDAServerUrl] = server;
  inputs[QuestionNames.MCPForDAServerName] = mcpName;
  const allMcpTools = vscode.lm.tools;
  const tools = allMcpTools
    .filter((tool: any) => (tool.name as string).includes(`mcp_${mcpName as string}`))
    .map((tool: any) => {
      const index = tool.name.indexOf(mcpName);
      const newName = (tool.name as string).substring(
        (index as number) + (mcpName as string).length + 1
      );
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
  let oauthMetadataUrl = undefined;
  try {
    const response = await axios.get(server);
  } catch (error) {
    if (error.status == 401) {
      auth = "OAuthPluginVault";
    }
    const errorDetails = error.response.headers["www-authenticate"];
    if (!errorDetails) {
      auth = "NoneAuth";
    } else {
      const match = errorDetails.match(/resource_metadata=\s*"([^"]+)"/);
      if (match) {
        oauthMetadataUrl = match[1];
      }
    }
  }
  inputs[QuestionNames.MCPForDAAuth] = auth;
  inputs[QuestionNames.MCPForDAAuthMetadataUrl] = oauthMetadataUrl;
  const result = await runCommand(Stage.updateActionWithMCP, inputs);
  return result;
}
