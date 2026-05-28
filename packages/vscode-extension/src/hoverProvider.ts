// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";
import { environmentNameManager, envUtil } from "@microsoft/teamsfx-core";
import { environmentVariableRegex } from "./constants";
import { getSystemInputs } from "./utils/systemEnvUtils";
import { DotenvParseOutput } from "dotenv";

export class ManifestTemplateHoverProvider implements vscode.HoverProvider {
  public async provideHover(
    document: vscode.TextDocument,
    position: vscode.Position,
    token: vscode.CancellationToken
  ): Promise<vscode.Hover | undefined> {
    const line = document.lineAt(position.line);

    const matches = environmentVariableRegex.exec(line.text);
    const regex = environmentVariableRegex;

    if (matches !== null && regex !== undefined) {
      const key = matches[0].replace(/{/g, "").replace(/}/g, "").replace(/\$/g, "");
      const indexOf = line.text.indexOf(matches[0]);
      const position = new vscode.Position(line.lineNumber, indexOf);
      const range = document.getWordRangeAtPosition(
        new vscode.Position(position.line, indexOf),
        new RegExp(regex)
      );
      const spfxLocal = document.fileName.endsWith("manifest.local.json");
      const message = await this.generateHoverMessageV3(key, spfxLocal);
      const hover = new vscode.Hover(message, range);
      return hover;
    }

    return undefined;
  }

  private async generateHoverMessageV3(
    key: string,
    displayLocalValue: boolean
  ): Promise<vscode.MarkdownString> {
    // Get environment variables
    const inputs = getSystemInputs();
    const envNamesRes = await envUtil.listEnv(inputs.projectPath!);
    if (envNamesRes.isErr()) {
      return new vscode.MarkdownString();
    }
    const envNames: string[] = envNamesRes.value;

    const envInfos: { [key: string]: DotenvParseOutput } = {};
    for (const envName of envNames) {
      const envRes = await envUtil.readEnv(inputs.projectPath!, envName, false);
      if (envRes.isOk()) {
        envInfos[envName] = envRes.value;
      }
    }

    // Generate hover message
    let message = "";
    let args;
    if (displayLocalValue) {
      const envName = environmentNameManager.getLocalEnvName();
      const envInfo = envInfos[envName];
      const value = envInfo ? envInfo[key] : undefined;
      if (value) {
        message = `**${envName}**: ${value} \n\n`;
      } else {
        message += `**${envName}** Trigger debug to see placeholder value \n\n`;
      }
      args = [{ type: "env", env: envName }];
    } else {
      for (const envName of envNames) {
        const envInfo = envInfos[envName];
        const value = envInfo[key];
        if (value) {
          message += `**${envName}**: ${value} \n\n`;
        } else {
          if (envName === environmentNameManager.getLocalEnvName()) {
            message += `**${envName}** Trigger debug to see placeholder value \n\n`;
          } else {
            const commandUri = vscode.Uri.parse("command:fx-extension.provision");
            message += `**${envName}**: [Trigger Teams: Provision in the cloud command to see placeholder value](${commandUri.toString()}) \n\n`;
          }
        }
      }
      args = [{ type: "env" }];
    }

    const commandUri = vscode.Uri.parse(
      `command:fx-extension.openConfigState?${encodeURIComponent(JSON.stringify(args))}`
    );
    message += `[✏️Edit env file](${commandUri.toString()})`;

    const markdown = new vscode.MarkdownString(message);
    markdown.isTrusted = true;
    return markdown;
  }
}
