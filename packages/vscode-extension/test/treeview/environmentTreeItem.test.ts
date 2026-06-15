import * as chai from "chai";
import * as vscode from "vscode";
import { vi } from "vitest";
import { mockValue } from "../mocks/vitestMockUtils";

import { FxError, LoginStatus, ok, Result, SubscriptionInfo } from "@microsoft/teamsfx-api";
import {
  FeatureFlags,
  GraphScopes,
  environmentNameManager,
  featureFlagManager,
} from "@microsoft/teamsfx-core";

import { M365Login } from "../../src/commonlib/m365Login";
import azureAccountManager from "../../src/commonlib/azureLogin";
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

  it("getChildren returns cached children", async () => {
    const environmentNode = new EnvironmentNode("test");
    vi.spyOn(environmentTreeItemDeps, "isRemoteEnvironment").mockReturnValue(false);

    const children1 = await environmentNode.getChildren();
    const children2 = await environmentNode.getChildren();

    chai.expect(children1).to.equal(children2);
  });

  it("getChildren adds warning when Azure account is not signed in", async () => {
    const environmentNode = new EnvironmentNode("test");
    vi.spyOn(environmentTreeItemDeps, "isRemoteEnvironment").mockReturnValue(true);
    vi.spyOn(environmentTreeItemDeps, "M365LoginGetStatus").mockResolvedValue(
      ok({
        status: "SignedIn",
        accountInfo: {
          tid: "tenant-id",
        },
      } as LoginStatus)
    );
    vi.spyOn(environmentTreeItemDeps, "getM365TenantFromEnv").mockResolvedValue("tenant-id");
    vi.spyOn(environmentTreeItemDeps, "isSPFxProject").mockReturnValue(false);
    vi.spyOn(environmentTreeItemDeps, "azureAccountManagerGetAccountInfo").mockReturnValue(
      undefined
    );
    vi.spyOn(environmentTreeItemDeps, "getSubscriptionInfoFromEnv").mockResolvedValue(undefined);
    vi.spyOn(environmentTreeItemDeps, "localize").mockImplementation((key: string) => key);

    const children = await environmentNode.getChildren();
    const warningNode = (await (children as DynamicNode[])[0].getTreeItem()) as DynamicNode;

    chai.expect(warningNode).to.not.be.undefined;
    chai
      .expect(String(warningNode.tooltip))
      .to.include("teamstoolkit.commandsTreeViewProvider.azureAccountNotSignedIn");
  });

  describe("environmentTreeItemDeps delegation", () => {
    it("delegates to environmentNameManager", () => {
      vi.spyOn(environmentNameManager, "isRemoteEnvironment").mockReturnValue(true);
      vi.spyOn(environmentNameManager, "getLocalEnvName").mockReturnValue("local");
      vi.spyOn(environmentNameManager, "getTestToolEnvName").mockReturnValue("testtool");

      chai.expect(environmentTreeItemDeps.isRemoteEnvironment("dev")).to.be.true;
      chai.expect(environmentTreeItemDeps.getLocalEnvName()).to.equal("local");
      chai.expect(environmentTreeItemDeps.getTestToolEnvName()).to.equal("testtool");
    });

    it("delegates to env tree utils", async () => {
      vi.spyOn(envTreeUtils, "getSubscriptionInfoFromEnv").mockResolvedValue(undefined);
      vi.spyOn(envTreeUtils, "getM365TenantFromEnv").mockResolvedValue("tenant");
      vi.spyOn(envTreeUtils, "getProvisionSucceedFromEnv").mockResolvedValue(true);
      vi.spyOn(envTreeUtils, "getResourceGroupNameFromEnv").mockResolvedValue("rg");

      const subscriptionInfo = await environmentTreeItemDeps.getSubscriptionInfoFromEnv("dev");
      const tenant = await environmentTreeItemDeps.getM365TenantFromEnv("dev");
      const provisioned = await environmentTreeItemDeps.getProvisionSucceedFromEnv("dev");
      const resourceGroup = await environmentTreeItemDeps.getResourceGroupNameFromEnv("dev");

      chai.expect(subscriptionInfo).to.be.undefined;
      chai.expect(tenant).to.equal("tenant");
      chai.expect(provisioned).to.be.true;
      chai.expect(resourceGroup).to.equal("rg");
    });

    it("delegates to localize, azure account and M365 login", async () => {
      vi.spyOn(localizeUtils, "localize").mockReturnValue("localized");
      vi.spyOn(azureAccountManager, "getAccountInfo").mockReturnValue(undefined);
      vi.spyOn(azureAccountManager, "listAllSubscriptions").mockResolvedValue([]);
      vi.spyOn(M365Login.getInstance(), "getStatus").mockResolvedValue(
        ok({ status: "SignedOut", accountInfo: {} } as LoginStatus)
      );

      const localized = environmentTreeItemDeps.localize("key");
      const account = environmentTreeItemDeps.azureAccountManagerGetAccountInfo();
      const subscriptions = await environmentTreeItemDeps.azureAccountManagerListAllSubscriptions();
      const status = await environmentTreeItemDeps.M365LoginGetStatus([]);

      chai.expect(localized).to.equal("localized");
      chai.expect(account).to.be.undefined;
      chai.expect(subscriptions).to.deep.equal([]);
      chai.expect(status.isOk()).to.be.true;
    });
  });
});
