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
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { TelemetryEvent } from "../telemetry/extTelemetryEvents";
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
import { getTriggerFromProperty } from "../utils/telemetryUtils";

function sanitizeMCPName(name: string): string {
  // Replace special characters except "-" with "_", but if two special characters are adjacent,
  // only replace with one "_". Finally, substring to the first 13 characters.
  return name
    .replace(/[^a-zA-Z0-9-]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .substring(0, 13);
}

function extractLocalServerIdentifier(serverConfig: any): string | undefined {
  return serverConfig?.type === "stdio" ? serverConfig.args?.[2] : undefined;
}

export async function updateActionWithMCP(args?: any[]): Promise<Result<any, FxError>> {
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.UpdateActionWithMCPStart,
    getTriggerFromProperty(args && args.length > 1 ? [args[1]] : undefined)
  );
  const inputs = getSystemInputs();
  let mcpName = args && args.length > 0 ? args[0].serverName : undefined;
  let server = args && args.length > 0 ? args[0].serverConfig?.url : undefined;
  let isLocalMCP = args && args.length > 0 && args[0].serverConfig?.type === "stdio";
  // For stdio type (local MCP), extract identifier from args array (e.g., args[2] contains the identifier)
  let localServerIdentifier =
    args && args.length > 0 ? extractLocalServerIdentifier(args[0].serverConfig) : undefined;

  // Sanitize mcpName if it's provided as an argument
  if (mcpName) {
    mcpName = sanitizeMCPName(mcpName);
  }

  if (!mcpName && !server && !isLocalMCP) {
    const projectPath = inputs.projectPath;
    if (!projectPath) {
      return err(
        new UserError(
          "da-mcp",
          ExtensionErrors.MCPFileNotFound,
          getDefaultString("teamstoolkit.MCP.FileNotFound"),
          localize("teamstoolkit.MCP.FileNotFound")
        )
      );
    }
    const mcpFile = path.join(projectPath, ".vscode", "mcp.json");
    if (!fs.pathExistsSync(mcpFile)) {
      void vscode.window.showErrorMessage(localize("teamstoolkit.MCP.FileNotFound"));
      const error = new UserError(
        "da-mcp",
        ExtensionErrors.MCPFileNotFound,
        getDefaultString("teamstoolkit.MCP.FileNotFound"),
        localize("teamstoolkit.MCP.FileNotFound")
      );
      ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.UpdateActionWithMCP, error);
      return err(error);
    }
    // const mcpContent = await fs.readJSON(mcpFile);
    const mcpOriginalContent = fs.readFileSync(mcpFile, "utf-8");
    const mcpContent = parser.parse(mcpOriginalContent);
    if (!mcpContent || !mcpContent.servers) {
      void vscode.window.showErrorMessage(localize("teamstoolkit.MCP.ContentInvalid"));
      const error = new UserError(
        "da-mcp",
        ExtensionErrors.MCPContentInvalid,
        getDefaultString("teamstoolkit.MCP.ContentInvalid"),
        localize("teamstoolkit.MCP.ContentInvalid")
      );
      ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.UpdateActionWithMCP, error);
      return err(error);
    }

    // TODO: support multiple MCP servers
    const mcpNames = Object.keys(mcpContent.servers);
    if (mcpNames.length === 0) {
      void vscode.window.showErrorMessage(localize("teamstoolkit.MCP.ServerNotFound"));
      const error = new UserError(
        "da-mcp",
        ExtensionErrors.MCPServerNotFound,
        getDefaultString("teamstoolkit.MCP.ServerNotFound"),
        localize("teamstoolkit.MCP.ServerNotFound")
      );
      ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.UpdateActionWithMCP, error);
      return err(error);
    }
    if (mcpNames.length === 1) {
      mcpName = sanitizeMCPName(mcpNames[0]);
      const serverConfig = mcpContent.servers[mcpNames[0]];
      server = serverConfig.url;
      isLocalMCP = serverConfig.type === "stdio";
      localServerIdentifier = extractLocalServerIdentifier(serverConfig);
    } else {
      const mcpNameSelection: SingleSelectConfig = {
        name: "mcpName",
        title: "Select MCP Server",
        options: mcpNames.map((name) => {
          const serverConfig = mcpContent.servers[name];
          const identifier = extractLocalServerIdentifier(serverConfig);
          const detail =
            serverConfig.type === "stdio"
              ? `${identifier as string}`
              : (serverConfig.url as string);
          return {
            id: name,
            label: name,
            detail: detail,
          };
        }),
      };
      const result = await VS_CODE_UI.selectOption(mcpNameSelection);
      if (result.isErr()) {
        void vscode.window.showErrorMessage(
          result.error.message || localize("teamstoolkit.MCP.SelectServerFailed")
        );
        ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.UpdateActionWithMCP, result.error);
        return err(result.error);
      }
      const originalMcpName = result.value.result as string;
      mcpName = originalMcpName.replace(/[^a-zA-Z0-9]/g, "").substring(0, 10);
      const serverConfig = mcpContent.servers[originalMcpName];
      server = serverConfig.url;
      isLocalMCP = serverConfig.type === "stdio";
      localServerIdentifier = extractLocalServerIdentifier(serverConfig);
    }
  }

  if (!mcpName || (!isLocalMCP && !server) || (isLocalMCP && !localServerIdentifier)) {
    void vscode.window.showErrorMessage(localize("teamstoolkit.MCP.NameOrServerUrlMissing"));
    const error = new UserError(
      "da-mcp",
      ExtensionErrors.MCPNameOrServerUrlMissing,
      getDefaultString("teamstoolkit.MCP.NameOrServerUrlMissing"),
      localize("teamstoolkit.MCP.NameOrServerUrlMissing")
    );
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.UpdateActionWithMCP, error);
    return err(error);
  }

  inputs[QuestionNames.MCPForDAServerUrl] = server;
  inputs[QuestionNames.MCPForDAServerName] = mcpName;
  if (isLocalMCP && localServerIdentifier) {
    inputs[QuestionNames.MCPLocalServerIdentifier] = localServerIdentifier;
  }

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
    const error = new UserError(
      "da-mcp",
      ExtensionErrors.MCPToolsNotFound,
      getDefaultString("teamstoolkit.MCP.ToolsNotFound"),
      localize("teamstoolkit.MCP.ToolsNotFound")
    );
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.UpdateActionWithMCP, error);
    return err(error);
  }
  inputs[QuestionNames.MCPForDAAvailableTools] = tools;

  let auth: "OAuthPluginVault" | "NoneAuth" = "NoneAuth";
  let oauthMetadataUrl = undefined;

  if (!isLocalMCP && server) {
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
  }

  inputs[QuestionNames.MCPForDAAuth] = auth;
  inputs[QuestionNames.MCPForDAAuthMetadataUrl] = oauthMetadataUrl;
  const result = await runCommand(Stage.updateActionWithMCP, inputs);
  if (result.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.UpdateActionWithMCP, result.error, {
      "auth-type": auth,
      "tool-number": tools.length.toString(),
      "mcp-type": isLocalMCP ? "local" : "remote",
    });
  } else {
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.UpdateActionWithMCP, {
      "auth-type": auth,
      "tool-number": tools.length.toString(),
      "mcp-type": isLocalMCP ? "local" : "remote",
    });
  }
  return result;
}
