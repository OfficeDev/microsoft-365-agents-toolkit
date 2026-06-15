// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";

import {
  featureFlagManager,
  FeatureFlags,
  getSideloadingStatus,
  isSandboxedEnabled,
} from "@microsoft/teamsfx-core";
import {
  checkSandboxCallback,
  checkSideloadingCallback,
} from "../../handlers/accounts/checkAccessCallback";
import { TelemetryTriggerFrom } from "../../telemetry/extTelemetryEvents";
import { localize } from "../../utils/localizeUtils";
import { DynamicNode } from "../dynamicNode";
import { errorIcon, infoIcon, passIcon } from "./common";
import M365TokenInstance from "../../commonlib/m365Login";

enum ContextValues {
  Normal = "checkSideloading",
  ShowInfo = "checkSideloading-info",
}

export const sideloadingNodeDeps = {
  getSideloadingStatus: (token: string) => getSideloadingStatus(token),
  getBooleanValue: (flag: Parameters<typeof featureFlagManager.getBooleanValue>[0]) =>
    featureFlagManager.getBooleanValue(flag),
  isSandboxedEnabled: () => isSandboxedEnabled(M365TokenInstance),
  checkSandboxCallback: () => checkSandboxCallback(),
  checkSideloadingCallback: () => checkSideloadingCallback(),
  localize: (key: string, ...args: any[]) => localize(key, ...args),
};

export class SideloadingNode extends DynamicNode {
  constructor(
    private eventEmitter: vscode.EventEmitter<DynamicNode | undefined | void>,
    public token: string
  ) {
    super("", vscode.TreeItemCollapsibleState.None);
    this.contextValue = ContextValues.Normal;
  }

  public override getChildren(): vscode.ProviderResult<DynamicNode[]> {
    return null;
  }

  public override async getTreeItem(): Promise<vscode.TreeItem> {
    let isSideloadingAllowed: boolean | undefined;
    if (this.token != "") {
      isSideloadingAllowed = await sideloadingNodeDeps.getSideloadingStatus(this.token);
      if (isSideloadingAllowed === false) {
        if (sideloadingNodeDeps.getBooleanValue(FeatureFlags.SandBoxedTeam)) {
          // Suggest users to use sandboxed containers for local testing
          const isSandboxedAllowed = await sideloadingNodeDeps.isSandboxedEnabled();
          if (isSandboxedAllowed) {
            await sideloadingNodeDeps.checkSandboxCallback();
          } else {
            await sideloadingNodeDeps.checkSideloadingCallback();
          }
        } else {
          await sideloadingNodeDeps.checkSideloadingCallback();
        }
      }
    }
    if (isSideloadingAllowed === undefined) {
      this.label = sideloadingNodeDeps.localize(
        "teamstoolkit.accountTree.sideloadingStatusUnknown"
      );
      this.iconPath = infoIcon;
      this.tooltip = sideloadingNodeDeps.localize(
        "teamstoolkit.accountTree.sideloadingStatusUnknownTooltip"
      );
      this.contextValue = ContextValues.Normal;
      this.command = undefined;
    } else if (isSideloadingAllowed) {
      this.label = sideloadingNodeDeps.localize("teamstoolkit.accountTree.sideloadingPass");
      this.iconPath = passIcon;
      this.tooltip = sideloadingNodeDeps.localize(
        "teamstoolkit.accountTree.sideloadingPassTooltip"
      );
      this.contextValue = ContextValues.Normal;
      this.command = undefined;
    } else {
      this.label = sideloadingNodeDeps.localize("teamstoolkit.accountTree.sideloadingWarning");
      this.iconPath = errorIcon;
      this.tooltip = sideloadingNodeDeps.localize(
        "teamstoolkit.accountTree.sideloadingWarningTooltip"
      );
      this.contextValue = ContextValues.ShowInfo;
      this.command = {
        title: this.label,
        command: "fx-extension.checkSideloading",
        arguments: [TelemetryTriggerFrom.TreeView, this],
      };
    }
    return this;
  }
}
