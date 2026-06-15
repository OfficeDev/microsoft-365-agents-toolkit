import * as chai from "chai";
import * as vscode from "vscode";
import { vi } from "vitest";
import { mockValue } from "../mocks/vitestMockUtils";

import { FxError, LoginStatus, ok, Result, SubscriptionInfo } from "@microsoft/teamsfx-api";
import { FeatureFlags, GraphScopes, featureFlagManager } from "@microsoft/teamsfx-core";

import { M365Login } from "../../src/commonlib/m365Login";
import * as globalVariables from "../../src/globalVariables";
import { warningIcon } from "../../src/treeview/account/common";
import { DynamicNode } from "../../src/treeview/dynamicNode";
import { EnvironmentNode, environmentTreeItemDeps } from "../../src/treeview/environmentTreeItem";
import * as commonUtils from "../../src/utils/commonUtils";
import * as localizeUtils from "../../src/utils/localizeUtils";
import * as envTreeUtils from "../../src/utils/envTreeUtils";

describe("EnvironmentNode", () => {
  it("getTreeItem for local", async () => {
    const environmentNode = new EnvironmentNode("local");
    vi.spyOn(environmentNode, "getChildren").mockReturnValue(Promise.resolve([]));

    const treeItem = await environmentNode.getTreeItem();

    chai.assert.deepEqual(treeItem.iconPath, new vscode.ThemeIcon("symbol-folder"));
    chai.assert.equal(treeItem.collapsibleState, vscode.TreeItemCollapsibleState.None);
    chai.assert.equal(treeItem.contextValue, "local");
  });

  it("getTreeItem for local", async () => {
    const environmentNode = new EnvironmentNode("testtool");
    vi.spyOn(environmentNode, "getChildren").mockReturnValue(Promise.resolve([]));

    const treeItem = await environmentNode.getTreeItem();

    chai.assert.deepEqual(treeItem.iconPath, new vscode.ThemeIcon("symbol-folder"));
    chai.assert.equal(treeItem.collapsibleState, vscode.TreeItemCollapsibleState.None);
    chai.assert.equal(treeItem.contextValue, "testtool");
  });

  it("getChildren returns warning for SPFx project", async () => {
    const environmentNode = new EnvironmentNode("test");
    vi.spyOn(environmentTreeItemDeps, "isRemoteEnvironment").mockReturnValue(true);
    vi.spyOn(environmentTreeItemDeps, "M365LoginGetStatus").mockReturnValue(
      Promise.resolve<Result<LoginStatus, FxError>>(
        ok({
          status: "SignedIn",
          accountInfo: {
            tid: "test",
          },
        })
      )
    );
    vi.spyOn(environmentTreeItemDeps, "getM365TenantFromEnv").mockReturnValue(
      Promise.resolve("m365TenantId")
    );
    vi.spyOn(environmentTreeItemDeps, "isSPFxProject").mockReturnValue(true);
    vi.spyOn(environmentTreeItemDeps, "getSubscriptionInfoFromEnv").mockReturnValue(
      Promise.resolve<SubscriptionInfo | undefined>({
        subscriptionName: "subscriptionName",
        subscriptionId: "subscriptionId",
        tenantId: "tenantId",
      })
    );
    vi.spyOn(environmentTreeItemDeps, "localize").mockImplementation((key: string) => {
      if (key === "teamstoolkit.commandsTreeViewProvider.m365AccountNotMatch") {
        return "test string";
      }
      return "";
    });

    const children = await environmentNode.getChildren();

    chai.assert.equal(children?.length, 2);
    const warningNode = (await (children as DynamicNode[])[0].getTreeItem()) as DynamicNode;
    chai.assert.deepEqual(warningNode.iconPath, warningIcon);
    chai.assert.equal(warningNode.tooltip, "test string");
    chai.assert.equal(warningNode.getChildren(), null);
    chai.assert.equal(warningNode.getTreeItem(), warningNode);
  });

  it("getChildren returns subscription", async () => {
    const environmentNode = new EnvironmentNode("test");
    vi.spyOn(environmentTreeItemDeps, "isRemoteEnvironment").mockReturnValue(true);
    vi.spyOn(environmentTreeItemDeps, "M365LoginGetStatus").mockReturnValue(
      Promise.resolve<Result<LoginStatus, FxError>>(
        ok({
          status: "SignedIn",
          accountInfo: {
            tid: "test",
          },
        })
      )
    );
    vi.spyOn(environmentTreeItemDeps, "getM365TenantFromEnv").mockReturnValue(
      Promise.resolve("test")
    );
    vi.spyOn(environmentTreeItemDeps, "isSPFxProject").mockReturnValue(true);
    vi.spyOn(environmentTreeItemDeps, "getSubscriptionInfoFromEnv").mockReturnValue(
      Promise.resolve<SubscriptionInfo | undefined>({
        subscriptionName: "subscriptionName",
        subscriptionId: "subscriptionId",
        tenantId: "tenantId",
      })
    );
    vi.spyOn(environmentTreeItemDeps, "localize").mockImplementation((key: string) => {
      if (key === "teamstoolkit.envTree.subscriptionTooltip") {
        return "'%s' environment is provisioned in Azure subscription '%s' (ID: %s)";
      }
      return "";
    });
    vi.spyOn(environmentTreeItemDeps, "getResourceGroupNameFromEnv").mockReturnValue(
      Promise.resolve("resource group")
    );

    const children = await environmentNode.getChildren();

    chai.assert.equal(children?.length, 1);
    const subscriptionNode = (await (children as DynamicNode[])[0].getTreeItem()) as DynamicNode;
    chai.assert.deepEqual(subscriptionNode.iconPath, new vscode.ThemeIcon("key"));
    chai.assert.equal(subscriptionNode.label, "subscriptionName");
    chai.assert.equal(
      subscriptionNode.tooltip,
      "'test' environment is provisioned in Azure subscription 'subscriptionName' (ID: subscriptionId)"
    );
    chai.assert.equal(subscriptionNode.description, "subscriptionId");
    const subscriptionNodeTreeItem = await subscriptionNode.getTreeItem();
    chai.assert.equal(subscriptionNodeTreeItem, subscriptionNode);

    const subscriptionNodeChildren = await subscriptionNode.getChildren();
    const resourceGroupNode = (subscriptionNodeChildren as DynamicNode[])[0];
    chai.assert.equal(resourceGroupNode.getTreeItem(), resourceGroupNode);
    chai.assert.isNull(resourceGroupNode.getChildren());
  });

  it("checkAccountForEnvironment uses Graph scopes in sovereign high", async () => {
    vi.spyOn(featureFlagManager, "getStringValue").mockReturnValue("GCC H");
    const environmentNode = new EnvironmentNode("test");
    const getStatusStub = vi
      .spyOn(M365Login.getInstance(), "getStatus")
      .mockResolvedValue(ok({ status: "SignedOut", accountInfo: {} } as any));
    mockValue(globalVariables, "isSPFxProject", true);

    await environmentNode.getChildren();

    chai.assert.isTrue(getStatusStub.calledOnceWithExactly({ scopes: GraphScopes }));
  });
});
