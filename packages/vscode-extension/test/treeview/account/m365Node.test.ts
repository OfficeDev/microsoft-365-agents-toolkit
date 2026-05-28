import { featureFlagManager } from "@microsoft/teamsfx-core";
import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import { AccountItemStatus, loadingIcon, m365Icon } from "../../../src/treeview/account/common";
import { M365AccountNode } from "../../../src/treeview/account/m365Node";
import { DynamicNode } from "../../../src/treeview/dynamicNode";
import * as tool from "@microsoft/teamsfx-core/build/common/tools";
import * as globalVariables from "../../../src/globalVariables";
import { MockTools } from "../../mocks/mockTools";
import { ok } from "@microsoft/teamsfx-api";

describe("m365Node", () => {
  const sandbox = sinon.createSandbox();
  const eventEmitter = new vscode.EventEmitter<DynamicNode | undefined | void>();

  afterEach(() => {
    sandbox.restore();
  });

  it("setSignedIn", async () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    const m365Node = new M365AccountNode(eventEmitter);
    await m365Node.setSignedIn("test upn", "");
    const treeItem = await m365Node.getTreeItem();

    chai.assert.equal(treeItem.iconPath, m365Icon);
    chai.assert.equal(treeItem.collapsibleState, vscode.TreeItemCollapsibleState.None);
    chai.assert.equal(treeItem.label, "test upn");
    chai.assert.equal(treeItem.contextValue, "signedinM365");
    chai.assert.equal(treeItem.command, undefined);
  });

  it("setSignedIn - multitenant", async () => {
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    sandbox.stub(globalVariables, "tools").value(new MockTools());
    sandbox
      .stub(globalVariables.tools.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(ok("test-token"));
    sandbox.stub(tool, "listAllTenants").resolves([
      {
        tenantId: "0022fd51-06f5-4557-8a34-69be98de6e20",
        displayName: "MSFT",
      },
      {
        tenantId: "313ef12c-d7cb-4f01-af90-1b113db5aa9a",
        displayName: "Cisco",
      },
    ]);

    const m365Node = new M365AccountNode(eventEmitter);
    await m365Node.setSignedIn("test upn", "0022fd51-06f5-4557-8a34-69be98de6e20");
    const treeItem = await m365Node.getTreeItem();

    chai.assert.equal(treeItem.iconPath, m365Icon);
    chai.assert.equal(treeItem.collapsibleState, vscode.TreeItemCollapsibleState.None);
    chai.assert.equal(treeItem.label, "test upn (MSFT)");
    chai.assert.equal(treeItem.contextValue, "signedinM365");
    chai.assert.equal(treeItem.command, undefined);
  });

  it("setSigningIn", async () => {
    const m365Node = new M365AccountNode(eventEmitter);
    m365Node.setSigningIn();
    const treeItem = await m365Node.getTreeItem();

    chai.assert.equal(treeItem.iconPath, loadingIcon);
    chai.assert.equal(treeItem.collapsibleState, vscode.TreeItemCollapsibleState.None);
    chai.assert.equal(treeItem.contextValue, "");
  });

  it("setSignedOut", async () => {
    const m365Node = new M365AccountNode(eventEmitter);
    m365Node.status = AccountItemStatus.SignedIn;
    await m365Node.setSignedOut();
    const treeItem = await m365Node.getTreeItem();

    chai.assert.equal(treeItem.iconPath, m365Icon);
    chai.assert.equal(treeItem.collapsibleState, vscode.TreeItemCollapsibleState.None);
    chai.assert.equal(treeItem.contextValue, "signinM365");
  });

  it("setSwitching", async () => {
    const m365Node = new M365AccountNode(eventEmitter);
    m365Node.setSwitching();
    const treeItem = await m365Node.getTreeItem();

    chai.assert.equal(treeItem.iconPath, loadingIcon);
    chai.assert.equal(treeItem.collapsibleState, vscode.TreeItemCollapsibleState.None);
    chai.assert.equal(treeItem.contextValue, "");
  });

  it("getChildren", () => {
    const m365Node = new M365AccountNode(eventEmitter);
    chai.assert.isDefined(m365Node.getChildren());
  });

  it("updateChecks", () => {
    const m365Node = new M365AccountNode(eventEmitter);
    m365Node.updateChecks("test token", false, false);
    chai.assert.isDefined(m365Node.getChildren());
    chai.assert.equal(2, (m365Node.getChildren() as any).length);
    m365Node.updateChecks("test token", true, false);
    chai.assert.isDefined(m365Node.getChildren());
    chai.assert.equal(2, (m365Node.getChildren() as any).length);
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    const m365NodeWithCopilot = new M365AccountNode(eventEmitter);
    m365NodeWithCopilot.updateChecks("test token", false, true);
    chai.assert.isDefined(m365NodeWithCopilot.getChildren());
    chai.assert.equal(2, (m365NodeWithCopilot.getChildren() as any).length);
    m365NodeWithCopilot.updateChecks("test token", true, true);
    chai.assert.isDefined(m365NodeWithCopilot.getChildren());
    chai.assert.equal(2, (m365NodeWithCopilot.getChildren() as any).length);
  });
});
