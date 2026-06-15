// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as util from "util";
import * as vscode from "vscode";
import * as path from "path";
import fs from "fs-extra";
import * as globalVariables from "../globalVariables";
import * as localizeUtils from "./localizeUtils";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { TelemetryEvent, TelemetryProperty } from "../telemetry/extTelemetryEvents";
import { FxError } from "@microsoft/teamsfx-api";

export const mcpUtilsDeps = {
  getWorkspaceUri: () => globalVariables.workspaceUri,
  getExtensionPath: () => globalVariables.context?.extensionPath || "",
  localize: (key: string, ...args: any[]) => localizeUtils.localize(key, ...args),
  showInformationMessage: (message: string, ...items: string[]) =>
    vscode.window.showInformationMessage(message, ...items),
  showErrorMessage: (message: string) => vscode.window.showErrorMessage(message),
  sendTelemetryEvent: (eventName: string, properties?: any) =>
    ExtTelemetry.sendTelemetryEvent(eventName as any, properties),
  sendTelemetryErrorEvent: (eventName: string, error: FxError, properties?: any) =>
    ExtTelemetry.sendTelemetryErrorEvent(eventName as any, error, properties),
  getMcpServers: () => vscode.workspace.getConfiguration("mcp").get("servers"),
  existsSync: (filePath: string) => fs.existsSync(filePath),
  mkdirSync: (dirPath: string, options?: Parameters<typeof fs.mkdirSync>[1]) =>
    fs.mkdirSync(dirPath, options),
  readFileSync: (filePath: string, encoding: BufferEncoding) => fs.readFileSync(filePath, encoding),
  writeFileSync: (file: any, data: any, encoding?: BufferEncoding) =>
    fs.writeFileSync(file, data, encoding),
  openSync: (filePath: string, flags: number | string, mode?: number) =>
    fs.openSync(filePath, flags, mode),
  closeSync: (fd: number) => fs.closeSync(fd),
};

/**
 * Setup MCP Server by checking for required files and prompting user to create them if missing
 */
export async function setupMCPServer(): Promise<void> {
  const workspaceUri = mcpUtilsDeps.getWorkspaceUri();
  if (!workspaceUri) {
    return; // No workspace opened
  }
  const workspaceRoot = workspaceUri.fsPath;

  // Check which files are missing
  const copilotInstructionsPath = path.join(workspaceRoot, ".github", "copilot-instructions.md");
  const mcpConfigPath = path.join(workspaceRoot, ".vscode", "mcp.json");
  const missingCopilotInstructions = !mcpUtilsDeps.existsSync(copilotInstructionsPath);
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

  const message = util.format(
    mcpUtilsDeps.localize("teamstoolkit.mcpUtils.setupMcpServer.message"),
    fileList
  );

  mcpUtilsDeps.sendTelemetryEvent(TelemetryEvent.PromptMCPServer, {
    [TelemetryProperty.MissingCopilotInstructions]: String(missingCopilotInstructions),
    [TelemetryProperty.MissingMCPConfig]: String(missingMcpConfig),
  });
  await mcpUtilsDeps
    .showInformationMessage(
      message,
      mcpUtilsDeps.localize("teamstoolkit.mcpUtils.setupMcpServer.confirm"),
      mcpUtilsDeps.localize("teamstoolkit.mcpUtils.setupMcpServer.skip")
    )
    .then((selection) => {
      if (selection !== mcpUtilsDeps.localize("teamstoolkit.mcpUtils.setupMcpServer.confirm")) {
        mcpUtilsDeps.sendTelemetryEvent(TelemetryEvent.PromptMCPServer, {
          [TelemetryProperty.UserSelection]: "skip",
        });
        return; // User chose to skip setup
      }

      try {
        if (missingCopilotInstructions) {
          createCopilotInstructionsFile(workspaceRoot, copilotInstructionsPath);
        }

        if (missingMcpConfig) {
          updateMCPConfigFile(workspaceRoot, mcpConfigPath);
        }
      } catch (error) {
        const errorMessage = util.format(
          mcpUtilsDeps.localize("teamstoolkit.mcpUtils.setupMcpServer.errorMessage"),
          (error as Error).toString()
        );
        void mcpUtilsDeps.showErrorMessage(errorMessage);
        mcpUtilsDeps.sendTelemetryErrorEvent(TelemetryEvent.PromptMCPServer, error as FxError, {
          [TelemetryProperty.UserSelection]: "confirm",
        });
        return; // Exit if there was an error creating files
      }

      const successMessage = mcpUtilsDeps.localize(
        "teamstoolkit.mcpUtils.setupMcpServer.successMessage"
      );
      void mcpUtilsDeps.showInformationMessage(successMessage);
      mcpUtilsDeps.sendTelemetryEvent(TelemetryEvent.PromptMCPServer, {
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
  if (!mcpUtilsDeps.existsSync(githubDir)) {
    mcpUtilsDeps.mkdirSync(githubDir, { recursive: true });
  }

  // Get default content from template
  const templatePath = path.join(
    mcpUtilsDeps.getExtensionPath(),
    "media/mcp",
    "copilot-instructions.md"
  );

  const defaultContent = mcpUtilsDeps.readFileSync(templatePath, "utf8");
  // Create copilot-instructions.md file with default content
  const fd = mcpUtilsDeps.openSync(
    copilotInstructionsPath,
    fs.constants.O_CREAT | fs.constants.O_EXCL | fs.constants.O_RDWR,
    0o600
  );
  mcpUtilsDeps.writeFileSync(fd, defaultContent, "utf8");
  mcpUtilsDeps.closeSync(fd);
}

/**
 * Create .vscode/mcp.json file with m365agentstoolkit mcp server configuration
 */
function updateMCPConfigFile(workspaceRoot: string, mcpConfigPath: string): void {
  if (!mcpUtilsDeps.existsSync(mcpConfigPath)) {
    const vscodeDir = path.join(workspaceRoot, ".vscode");

    // Create .vscode directory if it doesn't exist
    if (!mcpUtilsDeps.existsSync(vscodeDir)) {
      mcpUtilsDeps.mkdirSync(vscodeDir, { recursive: true });
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

    mcpUtilsDeps.writeFileSync(mcpConfigPath, JSON.stringify(mcpConfig, null, 2), "utf8");
  } else {
    // If mcp.json already exists, check if it needs to be updated
    const configContent = mcpUtilsDeps.readFileSync(mcpConfigPath, "utf8");
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

    const fd = mcpUtilsDeps.openSync(mcpConfigPath, fs.constants.O_RDWR, 0o600);
    mcpUtilsDeps.writeFileSync(fd, JSON.stringify(config, null, 2), "utf8");
    mcpUtilsDeps.closeSync(fd);
  }
}

/**
 * Check if MCP config file needs to be updated with proper m365agentstoolkit server configuration
 */
function checkMCPConfigNeedsUpdate(): boolean {
  try {
    const userServers = mcpUtilsDeps.getMcpServers();
    if (userServers && JSON.stringify(userServers).includes("@microsoft/m365agentstoolkit-mcp")) {
      return false; // User already has m365agentstoolkit server configured
    }

    return true;
  } catch (error) {
    return true; // If we can't parse it, assume it needs update
  }
}
