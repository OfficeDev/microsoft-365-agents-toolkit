// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import * as vscode from "vscode";
import { ok, FxError } from "@microsoft/teamsfx-api";
import { Correlator } from "@microsoft/teamsfx-core";
import { WebviewPanel } from "../controls/webviewPanel";
import { TreatmentVariableValue } from "../exp/treatmentVariables";
import { isTeamsFxProject, isOfficeAddInProject } from "../globalVariables";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  InProductGuideInteraction,
  TelemetryEvent,
  TelemetryProperty,
  TelemetryTriggerFrom,
} from "../telemetry/extTelemetryEvents";
import { localize } from "../utils/localizeUtils";
import { getTriggerFromProperty, isTriggerFromWalkThrough } from "../utils/telemetryUtils";
import { createNewProjectHandler } from "./lifecycleHandlers";
import { PanelType } from "../controls/PanelType";

export async function openReadMeHandler(...args: unknown[]) {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ClickOpenReadMe, getTriggerFromProperty(args));
  if (!isTeamsFxProject && !isOfficeAddInProject) {
    const createProject = {
      title: localize("teamstoolkit.handlers.createProjectTitle"),
      run: async (): Promise<void> => {
        await Correlator.run(
          async () => await createNewProjectHandler(TelemetryTriggerFrom.Notification)
        );
      },
    };

    const openFolder = {
      title: localize("teamstoolkit.handlers.openFolderTitle"),
      run: async (): Promise<void> => {
        await vscode.commands.executeCommand("vscode.openFolder");
      },
    };

    void vscode.window
      .showInformationMessage(
        localize("teamstoolkit.handlers.createProjectNotification"),
        createProject,
        openFolder
      )
      .then((selection) => {
        selection?.run();
      });
  } else if (vscode.workspace.workspaceFolders && vscode.workspace.workspaceFolders.length > 0) {
    const workspaceFolder = vscode.workspace.workspaceFolders[0];
    const workspacePath: string = workspaceFolder.uri.fsPath;
    // show README.md or src/README.md(SPFx) in workspace root folder
    const rootReadmePath = `${workspacePath}/README.md`;
    const uri = (await fs.pathExists(rootReadmePath))
      ? vscode.Uri.file(rootReadmePath)
      : vscode.Uri.file(`${workspacePath}/src/README.md`);

    if (TreatmentVariableValue.inProductDoc) {
      const content = await fs.readFile(uri.fsPath, "utf8");
      if (content.includes("## Get Started with the Notification bot")) {
        // A notification bot project.
        if (content.includes("express")) {
          // Express server notification bot.
          ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InteractWithInProductDoc, {
            [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.Auto,
            [TelemetryProperty.Interaction]: InProductGuideInteraction.Open,
            [TelemetryProperty.Identifier]: PanelType.ExpressServerNotificationBotReadme,
          });
          WebviewPanel.createOrShow(PanelType.ExpressServerNotificationBotReadme);
        } else {
          ExtTelemetry.sendTelemetryEvent(TelemetryEvent.InteractWithInProductDoc, {
            [TelemetryProperty.TriggerFrom]: TelemetryTriggerFrom.Auto,
            [TelemetryProperty.Interaction]: InProductGuideInteraction.Open,
            [TelemetryProperty.Identifier]: PanelType.FunctionBasedNotificationBotReadme,
          });
          WebviewPanel.createOrShow(PanelType.FunctionBasedNotificationBotReadme);
        }
      }
    }

    // Always open README.md in current panel instead of side-by-side.
    await vscode.workspace.openTextDocument(uri);
    const PreviewMarkdownCommand = "markdown.showPreview";
    await vscode.commands.executeCommand(PreviewMarkdownCommand, uri);
  }
  return ok<unknown, FxError>(null);
}

export async function openSampleReadmeHandler(args?: any) {
  if (vscode.workspace.workspaceFolders && vscode.workspace.workspaceFolders.length > 0) {
    const workspaceFolder = vscode.workspace.workspaceFolders[0];
    const workspacePath: string = workspaceFolder.uri.fsPath;
    const uri = vscode.Uri.file(`${workspacePath}/README.md`);
    await vscode.workspace.openTextDocument(uri);
    if (isTriggerFromWalkThrough(args as unknown[])) {
      const PreviewMarkdownCommand = "markdown.showPreviewToSide";
      await vscode.commands.executeCommand(PreviewMarkdownCommand, uri);
    } else {
      const PreviewMarkdownCommand = "markdown.showPreview";
      await vscode.commands.executeCommand(PreviewMarkdownCommand, uri);
    }
  }
}
