// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as util from "util";
import * as vscode from "vscode";
import { SubscriptionInfo } from "@microsoft/teamsfx-api";
import { AppStudioScopes, environmentNameManager } from "@microsoft/teamsfx-core";
import { M365Login } from "../commonlib/m365Login";
import azureAccountManager from "../commonlib/azureLogin";
import { signedIn } from "../commonlib/common/constant";
import { isSPFxProject } from "../globalVariables";
import {
  getM365TenantFromEnv,
  getProvisionSucceedFromEnv,
  getResourceGroupNameFromEnv,
  getSubscriptionInfoFromEnv,
} from "../utils/envTreeUtils";
import { localize } from "../utils/localizeUtils";
import { DynamicNode } from "./dynamicNode";

enum EnvInfo {
  Local = "local",
  TestTool = "testtool",
  LocalForExistingApp = "local-existing-app",
  RemoteEnv = "environment",
  ProvisionedRemoteEnv = "environment-provisioned",
}

interface accountStatus {
  isM365AccountLogin: boolean;

  // azure account is optional for SPFx and existing tab app
  isAzureAccountLogin?: boolean;

  warnings: string[];
}

const provisionedIcon = new vscode.ThemeIcon("folder-active");
const nonProvisionedIcon = new vscode.ThemeIcon("symbol-folder");
const warningIcon = new vscode.ThemeIcon("warning");
const subscriptionIcon = new vscode.ThemeIcon("key");
const resourceGroupIcon = new vscode.ThemeIcon("symbol-method");

export class EnvironmentNode extends DynamicNode {
  private _children: DynamicNode[] | undefined;

  constructor(public identifier: string) {
    super(identifier, vscode.TreeItemCollapsibleState.None);
  }

  public async getChildren(): Promise<DynamicNode[] | undefined | null> {
    if (this._children !== undefined) {
      return this._children;
    }

    const children: DynamicNode[] = [];
    if (environmentNameManager.isRemoteEnvironment(this.identifier)) {
      // check account status
      const accountStatus = await this.checkAccountForEnvironment(this.identifier);
      if (!accountStatus.isM365AccountLogin || accountStatus.isAzureAccountLogin === false) {
        const warningNode = new WarningNode(this.identifier, accountStatus);
        children.push(warningNode);
      }
      // show subscription
      const subscriptionInfo = await getSubscriptionInfoFromEnv(this.identifier);
      if (subscriptionInfo) {
        const subscriptionNode = new SubscriptionNode(this.identifier, subscriptionInfo);
        children.push(subscriptionNode);
      }
    }

    this._children = children;
    return children;
  }

  public async getTreeItem(): Promise<vscode.TreeItem> {
    const envInfo = await this.getCurrentEnvInfo(this.identifier);
    this.iconPath = envInfo === EnvInfo.ProvisionedRemoteEnv ? provisionedIcon : nonProvisionedIcon;

    const children = await this.getChildren();
    this.collapsibleState =
      children && children.length > 0
        ? vscode.TreeItemCollapsibleState.Expanded
        : vscode.TreeItemCollapsibleState.None;

    this.description = envInfo === EnvInfo.ProvisionedRemoteEnv ? "(Provisioned)" : "";
    this.contextValue = envInfo;
    return this;
  }

  private async checkAccountForEnvironment(env: string): Promise<accountStatus> {
    let isM365AccountLogin = true;
    const warnings: string[] = [];

    // Check M365 account status
    const loginStatusRes = await M365Login.getInstance().getStatus({ scopes: AppStudioScopes });
    const loginStatus = loginStatusRes.isOk() ? loginStatusRes.value : undefined;
    if (loginStatus && loginStatus.status == signedIn) {
      // Signed account doesn't match
      const m365TenantId = await getM365TenantFromEnv(env);
      if (m365TenantId && loginStatus.accountInfo?.tid !== m365TenantId) {
        isM365AccountLogin = false;
        warnings.push(localize("teamstoolkit.commandsTreeViewProvider.m365AccountNotMatch"));
      }
    } else {
      // Not signed in
      isM365AccountLogin = false;
      warnings.push(localize("teamstoolkit.commandsTreeViewProvider.m365AccountNotSignedIn"));
    }

    // Check Azure account status
    if (isSPFxProject) {
      return {
        isM365AccountLogin,
        warnings,
      };
    }

    let isAzureAccountLogin = true;
    if (azureAccountManager.getAccountInfo() !== undefined) {
      const subscriptionInfo = await getSubscriptionInfoFromEnv(env);
      const provisionedSubId = subscriptionInfo?.subscriptionId;

      if (provisionedSubId) {
        const subscriptions: SubscriptionInfo[] = await azureAccountManager.listAllSubscriptions();
        const targetSub = subscriptions.find(
          (sub) => sub.subscriptionId === subscriptionInfo?.subscriptionId
        );
        if (targetSub === undefined) {
          isAzureAccountLogin = false;
          warnings.push(
            util.format(
              localize("teamstoolkit.commandsTreeViewProvider.azureAccountNotMatch"),
              subscriptionInfo?.subscriptionName ?? subscriptionInfo?.subscriptionId
            )
          );
        }
      }
    } else {
      isAzureAccountLogin = false;
      warnings.push(localize("teamstoolkit.commandsTreeViewProvider.azureAccountNotSignedIn"));
    }

    return {
      isM365AccountLogin,
      isAzureAccountLogin,
      warnings,
    };
  }

  // Get the environment info for the given environment name.
  private async getCurrentEnvInfo(envName: string): Promise<EnvInfo> {
    if (envName === environmentNameManager.getLocalEnvName()) {
      return EnvInfo.Local;
    } else if (envName === environmentNameManager.getTestToolEnvName()) {
      return EnvInfo.TestTool;
    } else {
      const provisionSucceeded = await getProvisionSucceedFromEnv(envName);
      return provisionSucceeded ? EnvInfo.ProvisionedRemoteEnv : EnvInfo.RemoteEnv;
    }
  }
}

class WarningNode extends DynamicNode {
  constructor(public identifier: string, accountStatus: accountStatus) {
    super(identifier, vscode.TreeItemCollapsibleState.None);
    if (accountStatus.isAzureAccountLogin === false && !accountStatus.isM365AccountLogin) {
      this.label = localize("teamstoolkit.envTree.missingAzureAndM365Account");
    } else if (!accountStatus.isM365AccountLogin) {
      this.label = localize("teamstoolkit.envTree.missingM365Account");
    } else if (accountStatus.isAzureAccountLogin === false) {
      this.label = localize("teamstoolkit.envTree.missingAzureAccount");
    }

    this.iconPath = warningIcon;
    this.tooltip = this.formatWarningMessages(accountStatus.warnings);
  }

  public override getChildren(): vscode.ProviderResult<DynamicNode[]> {
    return null;
  }

  public override getTreeItem(): vscode.TreeItem | Promise<vscode.TreeItem> {
    return this;
  }

  private formatWarningMessages(warnings: string[]): string {
    let warningMessage = "";
    if (warnings.length > 1) {
      const formattedWarnings = warnings.map((warning) => "> ".concat(warning));
      warningMessage = formattedWarnings.join("\n");
    } else {
      warningMessage = warnings[0];
    }

    return warningMessage;
  }
}

class SubscriptionNode extends DynamicNode {
  private resourceGroupNode?: ResourceGroupNode;
  constructor(public identifier: string, private subscriptionInfo: SubscriptionInfo) {
    super(identifier, vscode.TreeItemCollapsibleState.None);
    this.contextValue = "openSubscriptionInPortal";
    this.iconPath = subscriptionIcon;
  }

  public override getChildren(): vscode.ProviderResult<DynamicNode[]> {
    if (this.resourceGroupNode) {
      return [this.resourceGroupNode];
    }
    return null;
  }

  public override async getTreeItem(): Promise<vscode.TreeItem> {
    this.tooltip = this.subscriptionInfo.subscriptionName
      ? util.format(
          localize("teamstoolkit.envTree.subscriptionTooltip"),
          this.identifier,
          this.subscriptionInfo.subscriptionName,
          this.subscriptionInfo.subscriptionId
        )
      : util.format(
          localize("teamstoolkit.envTree.subscriptionTooltipWithoutName"),
          this.identifier,
          this.subscriptionInfo.subscriptionId
        );
    this.label = this.subscriptionInfo.subscriptionName ?? this.subscriptionInfo.subscriptionId;
    this.description = this.subscriptionInfo.subscriptionId;

    const resourceGroupName = await getResourceGroupNameFromEnv(this.identifier);
    if (resourceGroupName) {
      this.resourceGroupNode = new ResourceGroupNode(this.identifier, resourceGroupName);
      this.collapsibleState = vscode.TreeItemCollapsibleState.Expanded;
    } else {
      this.resourceGroupNode = undefined;
      this.collapsibleState = vscode.TreeItemCollapsibleState.None;
    }

    return this;
  }
}

class ResourceGroupNode extends DynamicNode {
  constructor(public identifier: string, private resourceGroup: string) {
    super(resourceGroup, vscode.TreeItemCollapsibleState.None);
    this.contextValue = "openResourceGroupInPortal";
    this.iconPath = resourceGroupIcon;
  }

  public override getChildren(): vscode.ProviderResult<DynamicNode[]> {
    return null;
  }

  public override getTreeItem(): vscode.TreeItem | Promise<vscode.TreeItem> {
    return this;
  }
}
