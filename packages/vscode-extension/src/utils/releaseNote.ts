// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";
import * as versionUtil from "./versionUtil";
import { PrereleaseState, SyncedState, UserState } from "../constants";
import * as util from "util";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { TelemetryEvent } from "../telemetry/extTelemetryEvents";
import * as folder from "../folder";
import { localize } from "./localizeUtils";

export class ReleaseNote {
  private context: vscode.ExtensionContext;

  constructor(context: vscode.ExtensionContext) {
    this.context = context;
  }

  public async show() {
    const extensionId = versionUtil.getExtensionId();
    const teamsToolkit = vscode.extensions.getExtension(extensionId);
    const teamsToolkitVersion = teamsToolkit?.packageJSON.version;
    if (teamsToolkitVersion === undefined) {
      return;
    }
    const isPrerelease = versionUtil.isPrereleaseVersion(teamsToolkitVersion);
    if (isPrerelease) {
      const currentPrereleaseVersion = this.context.globalState.get<string>(
        PrereleaseState.Version
      );
      if (
        currentPrereleaseVersion === undefined ||
        teamsToolkitVersion != currentPrereleaseVersion
      ) {
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ShowWhatIsNewNotification);
        await this.context.globalState.update(PrereleaseState.Version, teamsToolkitVersion);
        await this.showChangelog("PRERELEASE.md");
      }
    } else {
      const currentStableVersion = this.context.globalState.get<string>(SyncedState.Version);
      await this.context.globalState.update(SyncedState.Version, teamsToolkitVersion);
      if (
        currentStableVersion !== undefined &&
        versionUtil.compare(teamsToolkitVersion, currentStableVersion) === 1
      ) {
        // it is existinig user
        await this.context.globalState.update(UserState.IsExisting, "yes");
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ShowWhatIsNewNotification);

        const changelog = {
          title: localize("teamstoolkit.upgrade.changelog"),
          run: async (): Promise<void> => {
            await this.showChangelog("CHANGELOG.md");
          },
        };

        void vscode.window
          .showInformationMessage(
            util.format(localize("teamstoolkit.upgrade.banner"), teamsToolkitVersion),
            changelog
          )
          .then(async (selection) => {
            if (selection?.title === localize("teamstoolkit.upgrade.changelog")) {
              ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ShowWhatIsNewContext);
              await selection.run();
            }
          });
      }
    }
  }

  private async showChangelog(filename: string) {
    const uri = vscode.Uri.file(`${folder.getResourceFolder()}/${filename}`);
    await vscode.workspace.openTextDocument(uri);
    const PreviewMarkdownCommand = "markdown.showPreview";
    await vscode.commands.executeCommand(PreviewMarkdownCommand, uri);
  }
}
