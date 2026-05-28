// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";

import { TokenProvider } from "@microsoft/teamsfx-api";
import { DynamicNode } from "../dynamicNode";
import envTreeProviderInstance from "../environmentTreeViewProvider";
import { AzureAccountNode } from "./azureNode";
import { M365AccountNode } from "./m365Node";
import { AppStudioScopes } from "@microsoft/teamsfx-core";
import { isSPFxProject } from "../../globalVariables";

class AccountTreeViewProvider implements vscode.TreeDataProvider<DynamicNode> {
  private static instance: AccountTreeViewProvider;
  private _onDidChangeTreeData: vscode.EventEmitter<DynamicNode | undefined | void> =
    new vscode.EventEmitter<DynamicNode | undefined | void>();
  readonly onDidChangeTreeData: vscode.Event<DynamicNode | undefined | void> =
    this._onDidChangeTreeData.event;

  public m365AccountNode = new M365AccountNode(this._onDidChangeTreeData);
  public azureAccountNode = new AzureAccountNode(this._onDidChangeTreeData);

  private constructor() {}

  public static getInstance() {
    if (!AccountTreeViewProvider.instance) {
      AccountTreeViewProvider.instance = new AccountTreeViewProvider();
    }
    return AccountTreeViewProvider.instance;
  }

  public subscribeToStatusChanges(tokenProvider: TokenProvider) {
    void tokenProvider.m365TokenProvider?.setStatusChangeMap(
      "tree-view",
      { scopes: AppStudioScopes },
      (status, token, accountInfo) =>
        m365AccountStatusChangeHandler("appStudio", status, token, accountInfo)
    );
    void tokenProvider.azureAccountProvider?.setStatusChangeMap(
      "tree-view",
      azureAccountStatusChangeHandler
    );
  }

  public getTreeItem(element: DynamicNode): vscode.TreeItem | Promise<vscode.TreeItem> {
    return element.getTreeItem();
  }

  public getChildren(element?: DynamicNode): vscode.ProviderResult<DynamicNode[]> {
    if (!element) {
      return this.getAccountNodes();
    }
    return element.getChildren();
  }

  private getAccountNodes(): DynamicNode[] {
    if (isSPFxProject) {
      return [this.m365AccountNode];
    } else {
      return [this.m365AccountNode, this.azureAccountNode];
    }
  }
}

async function m365AccountStatusChangeHandler(
  source: string,
  status: string,
  token?: string | undefined,
  accountInfo?: Record<string, unknown> | undefined
) {
  const instance = AccountTreeViewProvider.getInstance();
  if (status === "SignedIn") {
    if (accountInfo) {
      await instance.m365AccountNode.setSignedIn(
        (accountInfo.upn as string) ?? (accountInfo.unique_name as string) ?? "",
        (accountInfo.tid as string) ?? ""
      );
      if (token && source === "appStudio") {
        instance.m365AccountNode.updateChecks(token, true, true);
      }
    }
  } else if (status === "SigningIn") {
    instance.m365AccountNode.setSigningIn();
  } else if (status === "SignedOut") {
    instance.m365AccountNode.setSignedOut();
  } else if (status == "Switching") {
    instance.m365AccountNode.setSwitching();
  }
  await envTreeProviderInstance.reloadEnvironments();
  return Promise.resolve();
}

async function azureAccountStatusChangeHandler(
  status: string,
  token?: string | undefined,
  accountInfo?: Record<string, unknown> | undefined
) {
  const instance = AccountTreeViewProvider.getInstance();
  if (status === "SignedIn") {
    const username = (accountInfo?.email as string) || (accountInfo?.upn as string);
    if (username) {
      await instance.azureAccountNode.setSignedIn(
        token as string,
        accountInfo?.tid as string,
        username
      );
      await envTreeProviderInstance.reloadEnvironments();
    }
  } else if (status === "SigningIn") {
    instance.azureAccountNode.setSigningIn();
  } else if (status === "SignedOut") {
    instance.azureAccountNode.setSignedOut();
    await envTreeProviderInstance.reloadEnvironments();
  }
  return Promise.resolve();
}

export default AccountTreeViewProvider.getInstance();
