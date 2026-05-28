// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";
import { MosServiceScope, PackageService } from "@microsoft/teamsfx-core";
import M365TokenInstance from "../../commonlib/m365Login";
import { TelemetryTriggerFrom } from "../../telemetry/extTelemetryEvents";
import { localize } from "../../utils/localizeUtils";
import { DynamicNode } from "../dynamicNode";
import { infoIcon, passIcon, warningIcon } from "./common";

enum ContextValues {
  Normal = "checkCopilot",
  ShowInfo = "checkCopilot-info",
}

const copilotCheckServiceScope = process.env.SIDELOADING_SERVICE_SCOPE ?? MosServiceScope;

export class CopilotNode extends DynamicNode {
  constructor(
    private eventEmitter: vscode.EventEmitter<DynamicNode | undefined | void>,
    public token: string
  ) {
    super("", vscode.TreeItemCollapsibleState.None);
    this.contextValue = ContextValues.Normal;
  }

  private async checkCopilot(): Promise<boolean | undefined> {
    try {
      const m365TokenStatus = await M365TokenInstance.getAccessToken({
        scopes: [copilotCheckServiceScope],
        showDialog: false,
      });
      if (m365TokenStatus.isOk()) {
        const m365TokenResult = m365TokenStatus.value;
        if (m365TokenResult !== undefined && m365TokenResult !== "") {
          return await PackageService.GetSharedInstance().getCopilotStatus(m365TokenResult, true);
        }
      }
    } catch (error) {
      return undefined;
    }

    return undefined;
  }

  public override getChildren(): vscode.ProviderResult<DynamicNode[]> {
    return null;
  }

  public override async getTreeItem(): Promise<vscode.TreeItem> {
    let isCopilotAllowed: boolean | undefined;
    if (this.token != "") {
      isCopilotAllowed = await this.checkCopilot();
      if (isCopilotAllowed === false) {
        // Don't popup the warning since it's noisy
        // await checkCopilotCallback();
      }
    }
    if (isCopilotAllowed === undefined) {
      this.label = localize("teamstoolkit.accountTree.copilotStatusUnknown");
      this.iconPath = infoIcon;
      this.tooltip = localize("teamstoolkit.accountTree.copilotStatusUnknownTooltip");
      this.contextValue = ContextValues.Normal;
      this.command = undefined;
    } else if (isCopilotAllowed) {
      this.label = localize("teamstoolkit.accountTree.copilotPass");
      this.iconPath = passIcon;
      this.tooltip = localize("teamstoolkit.accountTree.copilotPassTooltip");
      this.contextValue = ContextValues.Normal;
      this.command = undefined;
    } else {
      this.label = localize("teamstoolkit.accountTree.copilotWarning");
      this.iconPath = warningIcon;
      this.tooltip = localize("teamstoolkit.accountTree.copilotWarningTooltip");
      this.contextValue = ContextValues.ShowInfo;
      this.command = {
        title: this.label,
        command: "fx-extension.checkCopilotCallback",
        arguments: [TelemetryTriggerFrom.TreeView, this],
      };
    }
    return this;
  }
}
