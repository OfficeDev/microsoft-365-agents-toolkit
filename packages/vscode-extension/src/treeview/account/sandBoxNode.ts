// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";

import { localize } from "../../utils/localizeUtils";
import { DynamicNode } from "../dynamicNode";
import { passIcon, warningIcon } from "./common";
import M365TokenInstance from "../../commonlib/m365Login";
import { GraphClient } from "@microsoft/teamsfx-core";
import { isSandboxedEnabled } from "../../debug/common/debugConstants";

enum ContextValues {
  Normal = "checkSideloading",
  ShowInfo = "checkSideloading-info",
}

export class SandboxNode extends DynamicNode {
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
    const isSandboxedAllowed = await isSandboxedEnabled(M365TokenInstance);
    if (isSandboxedAllowed) {
      this.contextValue = ContextValues.ShowInfo;
      this.label = localize("teamstoolkit.accountTree.sandboxedTeamEnabled");
      this.iconPath = passIcon;
      this.tooltip = localize("teamstoolkit.accountTree.sandboxedTeamEnabled.tooltip");
      this.contextValue = ContextValues.Normal;
      this.command = undefined;
    } else {
      this.label = localize("teamstoolkit.accountTree.sandboxedTeamDisabled");
      this.iconPath = warningIcon;
      this.tooltip = localize("teamstoolkit.accountTree.sandboxedTeamDisabled.tooltip");
      this.contextValue = ContextValues.Normal;
      this.command = undefined;
    }
    return this;
  }
}
