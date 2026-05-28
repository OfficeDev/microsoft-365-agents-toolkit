import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";

import { FxError, LoginStatus, ok, Result, SubscriptionInfo } from "@microsoft/teamsfx-api";

import { M365Login } from "../../src/commonlib/m365Login";
import * as globalVariables from "../../src/globalVariables";
import { warningIcon } from "../../src/treeview/account/common";
import { DynamicNode } from "../../src/treeview/dynamicNode";
import { EnvironmentNode } from "../../src/treeview/environmentTreeItem";
import * as commonUtils from "../../src/utils/commonUtils";
import * as localizeUtils from "../../src/utils/localizeUtils";
import * as envTreeUtils from "../../src/utils/envTreeUtils";

describe("EnvironmentNode", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("getTreeItem for local", async () => {
    const environmentNode = new EnvironmentNode("local");
    sandbox.stub(environmentNode, "getChildren").returns(Promise.resolve([]));

    const treeItem = await environmentNode.getTreeItem();

    chai.assert.deepEqual(treeItem.iconPath, new vscode.ThemeIcon("symbol-folder"));
    chai.assert.equal(treeItem.collapsibleState, vscode.TreeItemCollapsibleState.None);
    chai.assert.equal(treeItem.contextValue, "local");
  });

  it("getTreeItem for local", async () => {
    const environmentNode = new EnvironmentNode("testtool");
    sandbox.stub(environmentNode, "getChildren").returns(Promise.resolve([]));

    const treeItem = await environmentNode.getTreeItem();

    chai.assert.deepEqual(treeItem.iconPath, new vscode.ThemeIcon("symbol-folder"));
    chai.assert.equal(treeItem.collapsibleState, vscode.TreeItemCollapsibleState.None);
    chai.assert.equal(treeItem.contextValue, "testtool");
  });

  it("getChildren returns warning for SPFx project", async () => {
    const environmentNode = new EnvironmentNode("test");
    sandbox.stub(M365Login.getInstance(), "getStatus").returns(
      Promise.resolve<Result<LoginStatus, FxError>>(
        ok({
          status: "SignedIn",
          accountInfo: {
            tid: "test",
          },
        })
      )
    );
    sandbox.stub(envTreeUtils, "getM365TenantFromEnv").returns(Promise.resolve("m365TenantId"));
    sandbox.stub(globalVariables, "isSPFxProject").value(true);
    // eslint-disable-next-line no-secrets/no-secrets
    sandbox.stub(envTreeUtils, "getSubscriptionInfoFromEnv").returns(
      Promise.resolve<SubscriptionInfo | undefined>({
        subscriptionName: "subscriptionName",
        subscriptionId: "subscriptionId",
        tenantId: "tenantId",
      })
    );
    sandbox.stub(localizeUtils, "localize").callsFake((key: string, _defValue?: string) => {
      // eslint-disable-next-line no-secrets/no-secrets
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
    sandbox.stub(M365Login.getInstance(), "getStatus").returns(
      Promise.resolve<Result<LoginStatus, FxError>>(
        ok({
          status: "SignedIn",
          accountInfo: {
            tid: "test",
          },
        })
      )
    );
    sandbox.stub(envTreeUtils, "getM365TenantFromEnv").returns(Promise.resolve("test"));
    sandbox.stub(globalVariables, "isSPFxProject").value(true);
    // eslint-disable-next-line no-secrets/no-secrets
    sandbox.stub(envTreeUtils, "getSubscriptionInfoFromEnv").returns(
      Promise.resolve<SubscriptionInfo | undefined>({
        subscriptionName: "subscriptionName",
        subscriptionId: "subscriptionId",
        tenantId: "tenantId",
      })
    );
    sandbox.stub(localizeUtils, "localize").callsFake((key: string, _defValue?: string) => {
      if (key === "teamstoolkit.envTree.subscriptionTooltip") {
        return "'%s' environment is provisioned in Azure subscription '%s' (ID: %s)";
      }
      return "";
    });
    sandbox
      .stub(envTreeUtils, "getResourceGroupNameFromEnv")
      .returns(Promise.resolve("resource group"));

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
});
