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
import * as parser from "jsonc-parser";
import { getDefaultString, localize } from "../utils/localizeUtils";
import { ExtensionErrors } from "../error/error";

function sanitizeMCPName(name: string): string {
  // Replace special characters except "-" with "_", but if two special characters are adjacent,
  // only replace with one "_". Finally, substring to the first 13 characters.
  return name
    .replace(/[^a-zA-Z0-9-]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .substring(0, 13);
}

export async function updateActionWithMCP(args?: any[]): Promise<Result<any, FxError>> {
  const inputs = getSystemInputs();
  let mcpName = args && args.length > 0 ? args[0].serverName : undefined;
  let server = args && args.length > 0 ? args[0].serverConfig?.url : undefined;

  // Sanitize mcpName if it's provided as an argument
  if (mcpName) {
    mcpName = sanitizeMCPName(mcpName);
  }

  if (!mcpName && !server) {
    const mcpFile = path.join(inputs.projectPath!, ".vscode", "mcp.json");
    if (!fs.pathExistsSync(mcpFile)) {
      void vscode.window.showErrorMessage(localize("teamstoolkit.MCP.FileNotFound"));
      return err(
        new UserError(
          "da-mcp",
          ExtensionErrors.MCPFileNotFound,
          getDefaultString("teamstoolkit.MCP.FileNotFound"),
          localize("teamstoolkit.MCP.FileNotFound")
        )
      );
    }
    // const mcpContent = await fs.readJSON(mcpFile);
    const mcpOriginalContent = fs.readFileSync(mcpFile, "utf-8");
    const mcpContent = parser.parse(mcpOriginalContent);
    if (!mcpContent || !mcpContent.servers) {
      void vscode.window.showErrorMessage(localize("teamstoolkit.MCP.ContentInvalid"));
      return err(
        new UserError(
          "da-mcp",
          ExtensionErrors.MCPContentInvalid,
          getDefaultString("teamstoolkit.MCP.ContentInvalid"),
          localize("teamstoolkit.MCP.ContentInvalid")
        )
      );
    }

    // TODO: support multiple MCP servers
    const mcpNames = Object.keys(mcpContent.servers);
    if (mcpNames.length === 0) {
      void vscode.window.showErrorMessage(localize("teamstoolkit.MCP.ServerNotFound"));
      return err(
        new UserError(
          "da-mcp",
          ExtensionErrors.MCPServerNotFound,
          getDefaultString("teamstoolkit.MCP.ServerNotFound"),
          localize("teamstoolkit.MCP.ServerNotFound")
        )
      );
    }
    if (mcpNames.length === 1) {
      mcpName = sanitizeMCPName(mcpNames[0]);
      server = mcpContent.servers[mcpNames[0]].url;
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
        void vscode.window.showErrorMessage(
          result.error.message || localize("teamstoolkit.MCP.SelectServerFailed")
        );
        return err(result.error);
      }
      const originalMcpName = result.value.result as string;
      mcpName = originalMcpName.replace(/[^a-zA-Z0-9]/g, "").substring(0, 10);
      server = mcpContent.servers[originalMcpName].url;
    }
  } else if (!mcpName || !server) {
    void vscode.window.showErrorMessage(localize("teamstoolkit.MCP.NameOrServerUrlMissing"));
    return err(
      new UserError(
        "da-mcp",
        ExtensionErrors.MCPNameOrServerUrlMissing,
        getDefaultString("teamstoolkit.MCP.NameOrServerUrlMissing"),
        localize("teamstoolkit.MCP.NameOrServerUrlMissing")
      )
    );
  }

  inputs[QuestionNames.MCPForDAServerUrl] = server;
  inputs[QuestionNames.MCPForDAServerName] = mcpName;
  const allMcpTools = vscode.lm.tools;
  const tools = allMcpTools
    .filter((tool: vscode.LanguageModelToolInformation) =>
      tool.name.includes(`mcp_${mcpName as string}`)
    )
    .map((tool: vscode.LanguageModelToolInformation) => {
      const index = tool.name.indexOf(mcpName);
      const newName = tool.name.substring(index + (mcpName as string).length + 1);
      return {
        name: newName,
        description: tool.description,
        inputSchema: tool.inputSchema,
        tags: tool.tags,
      };
    });
  if (tools.length === 0) {
    void vscode.window.showErrorMessage(localize("teamstoolkit.MCP.ToolsNotFound"));
    // Return an error result
    return err(
      new UserError(
        "da-mcp",
        ExtensionErrors.MCPToolsNotFound,
        getDefaultString("teamstoolkit.MCP.ToolsNotFound"),
        localize("teamstoolkit.MCP.ToolsNotFound")
      )
    );
  }
  inputs[QuestionNames.MCPForDAAvailableTools] = tools;

  let auth: "OAuthPluginVault" | "NoneAuth" = "NoneAuth";
  let oauthMetadataUrl = undefined;
  try {
    await axios.get(server);
  } catch (error) {
    if (error.status == 401) {
      auth = "OAuthPluginVault";
      const errorDetails = error.response?.headers?.["www-authenticate"];
      if (errorDetails) {
        const match = errorDetails.match(/resource_metadata=\s*"([^"]+)"/);
        if (match) {
          oauthMetadataUrl = match[1];
        }
      }
    }
  }
  if (auth === "OAuthPluginVault" && !oauthMetadataUrl) {
    const originalURL = new URL(server);
    const wellKnownURL = `${originalURL.protocol}//${originalURL.host}/.well-known/oauth-authorization-server`;
    try {
      const response = await axios.get(wellKnownURL);
      if (response.status === 200) {
        inputs[QuestionNames.MCPForDAAuthWellKnownUrl] = wellKnownURL;
      }
    } finally {
    }
  }
  inputs[QuestionNames.MCPForDAAuth] = auth;
  inputs[QuestionNames.MCPForDAAuthMetadataUrl] = oauthMetadataUrl;
  const result = await runCommand(Stage.updateActionWithMCP, inputs);
  return result;
}
