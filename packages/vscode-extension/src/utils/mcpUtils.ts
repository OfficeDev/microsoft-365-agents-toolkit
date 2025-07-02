// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as util from "util";
import * as vscode from "vscode";
import * as path from "path";
import fs from "fs-extra";
import { context } from "../globalVariables";
import { localize } from "./localizeUtils";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { TelemetryEvent, TelemetryProperty } from "../telemetry/extTelemetryEvents";

/**
 * Setup MCP Server by checking for required files and prompting user to create them if missing
 */
export async function setupMCPServer(): Promise<void> {
  const workspaceFolders = vscode.workspace.workspaceFolders;
  if (!workspaceFolders || workspaceFolders.length === 0) {
    return;
  }

  const workspaceRoot = workspaceFolders[0].uri.fsPath;

  // Check which files are missing
  const copilotInstructionsPath = path.join(workspaceRoot, ".github", "copilot-instructions.md");
  const mcpConfigPath = path.join(workspaceRoot, ".vscode", "mcp.json");
  const missingCopilotInstructions = !fs.existsSync(copilotInstructionsPath);
  const missingMcpConfig = checkMCPConfigNeedsUpdate();

  if (!missingCopilotInstructions && !missingMcpConfig) {
    return;
  }

  // Create message based on which files are missing or need updates
  const filesToModify: string[] = [];
  if (missingCopilotInstructions) {
    filesToModify.push(".github/copilot-instructions.md");
  }
  if (missingMcpConfig) {
    filesToModify.push(".vscode/mcp.json");
  }

  const fileList =
    filesToModify.length === 1
      ? filesToModify[0]
      : filesToModify.length === 2
      ? `${filesToModify[0]} and ${filesToModify[1]}`
      : filesToModify.join(", ");

  const message = util.format(localize("teamstoolkit.mcpUtils.setupMcpServer.message"), fileList);

  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.PromptMCPServer, {
    [TelemetryProperty.MissingCopilotInstructions]: String(missingCopilotInstructions),
    [TelemetryProperty.MissingMCPConfig]: String(missingMcpConfig),
  });
  await vscode.window
    .showInformationMessage(
      message,
      localize("teamstoolkit.mcpUtils.setupMcpServer.confirm"),
      localize("teamstoolkit.mcpUtils.setupMcpServer.skip")
    )
    .then((selection) => {
      if (selection !== localize("teamstoolkit.mcpUtils.setupMcpServer.confirm")) {
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.UserSelection]: "skip",
        });
        return; // User chose to skip setup
      }

      if (missingCopilotInstructions) {
        createCopilotInstructionsFile(workspaceRoot, copilotInstructionsPath);
      }

      if (missingMcpConfig) {
        updateMCPConfigFile(workspaceRoot, mcpConfigPath);
      }
      const successMessage = localize("teamstoolkit.mcpUtils.setupMcpServer.successMessage");
      void vscode.window.showInformationMessage(successMessage);
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.PromptMCPServer, {
        [TelemetryProperty.UserSelection]: "confirm",
      });
    });
}

/**
 * Create .github/copilot-instructions.md file with default content
 */
function createCopilotInstructionsFile(
  workspaceRoot: string,
  copilotInstructionsPath: string
): void {
  const githubDir = path.join(workspaceRoot, ".github");

  // Create .github directory if it doesn't exist
  if (!fs.existsSync(githubDir)) {
    fs.mkdirSync(githubDir, { recursive: true });
  }

  // Get default content from template
  const templatePath = path.join(
    context?.extensionPath || "",
    "media/mcp",
    "copilot-instructions.md"
  );
  let defaultContent = "";

  try {
    if (fs.existsSync(templatePath)) {
      defaultContent = fs.readFileSync(templatePath, "utf8");
    }
  } catch (error) {
    console.warn("Failed to read copilot instructions template:", error);
  }

  // Create copilot-instructions.md file with default content
  fs.writeFileSync(copilotInstructionsPath, defaultContent, "utf8");
}

/**
 * Create .vscode/mcp.json file with m365agentstoolkit mcp server configuration
 */
function updateMCPConfigFile(workspaceRoot: string, mcpConfigPath: string): void {
  if (!fs.existsSync(mcpConfigPath)) {
    const vscodeDir = path.join(workspaceRoot, ".vscode");

    // Create .vscode directory if it doesn't exist
    if (!fs.existsSync(vscodeDir)) {
      fs.mkdirSync(vscodeDir, { recursive: true });
    }

    // Create mcp.json with m365agentstoolkit server configuration
    const mcpConfig = {
      servers: {
        m365agentstoolkit: {
          command: "npx",
          args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
        },
      },
    };

    fs.writeFileSync(mcpConfigPath, JSON.stringify(mcpConfig, null, 2), "utf8");
  } else {
    // If mcp.json already exists, check if it needs to be updated
    const configContent = fs.readFileSync(mcpConfigPath, "utf8");
    const config = JSON.parse(configContent);

    // Ensure servers object exists
    if (!config.servers || typeof config.servers !== "object") {
      config.servers = {};
    }

    // Add m365agentstoolkit server configuration
    config.servers["M365AgentsToolkit MCP Server"] = {
      command: "npx",
      args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
    };

    fs.writeFileSync(mcpConfigPath, JSON.stringify(config, null, 2), "utf8");
  }

  // const obj = {
  //   name: "M365 Agents Toolkit MCP Server",
  //   command: "npx",
  //   args: ["@microsoft/m365agentstoolkit-mcp@latest", "server", "start"],
  // };
  // const link = `vscode:mcp/install?${encodeURIComponent(JSON.stringify(obj))}`;
  // void vscode.commands.executeCommand("vscode.open", vscode.Uri.parse(link));
}

/**
 * Check if MCP config file needs to be updated with proper m365agentstoolkit server configuration
 */
function checkMCPConfigNeedsUpdate(): boolean {
  try {
    const userMCPSettings = vscode.workspace.getConfiguration("mcp");
    const userServers = userMCPSettings.get("servers");
    if (userServers && JSON.stringify(userServers).includes("@microsoft/m365agentstoolkit-mcp")) {
      return false; // User already has m365agentstoolkit server configured
    }

    return true;
  } catch (error) {
    return true; // If we can't parse it, assume it needs update
  }
}
