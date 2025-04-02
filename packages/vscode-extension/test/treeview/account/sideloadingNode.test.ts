import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as tools from "@microsoft/teamsfx-core/build/common/tools";
import { errorIcon, infoIcon, passIcon } from "../../../src/treeview/account/common";
import { SideloadingNode } from "../../../src/treeview/account/sideloadingNode";
import { DynamicNode } from "../../../src/treeview/dynamicNode";
import * as checkAccessCallback from "../../../src/handlers/accounts/checkAccessCallback";
import { featureFlagManager, GraphClient } from "@microsoft/teamsfx-core";
import { SANDBOX_SENSITIVITY_LABEL } from "@microsoft/teamsfx-core/src/common/tools";

describe("sideloadingNode", () => {
  const sandbox = sinon.createSandbox();
  const eventEmitter = new vscode.EventEmitter<DynamicNode | undefined | void>();

  afterEach(() => {
    sandbox.restore();
  });

  it("getTreeItem with empty string", async () => {
    const sideloadingNode = new SideloadingNode(eventEmitter, "");
    const treeItem = await sideloadingNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, infoIcon);
  });

  it("getTreeItem with invalid token", async () => {
    sandbox.stub(tools, "getSideloadingStatus").returns(Promise.resolve(false));
    sandbox.stub(checkAccessCallback, "checkSideloadingCallback");
    const sideloadingNode = new SideloadingNode(eventEmitter, "token");
    const treeItem = await sideloadingNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, errorIcon);
  });

  it("getTreeItem with valid token", async () => {
    sandbox.stub(tools, "getSideloadingStatus").returns(Promise.resolve(true));
    const sideloadingNode = new SideloadingNode(eventEmitter, "token");
    const treeItem = await sideloadingNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, passIcon);
  });

  it("getChildren", () => {
    const sideloadingNode = new SideloadingNode(eventEmitter, "token");
    chai.assert.isNull(sideloadingNode.getChildren());
  });

  it("Check sandbox permission", async () => {
    sandbox.stub(tools, "getSideloadingStatus").returns(Promise.resolve(false));
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    sandbox.stub(GraphClient.prototype, "GetTeamsAppSettingsAsync").resolves({
      sandboxingConfiguration: {
        isSideloadingEnabled: false,
        sensitivityLabelUsedToIdentifySandboxedContainers: SANDBOX_SENSITIVITY_LABEL,
      },
    });
    sandbox.stub(checkAccessCallback, "checkSandboxCallback");
    const sideloadingNode = new SideloadingNode(eventEmitter, "token");
    const treeItem = await sideloadingNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, errorIcon);
  });

  it("Check sandbox permission - disabled", async () => {
    sandbox.stub(tools, "getSideloadingStatus").returns(Promise.resolve(false));
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    sandbox.stub(GraphClient.prototype, "GetTeamsAppSettingsAsync").resolves({
      sandboxingConfiguration: {
        isSideloadingEnabled: false,
        sensitivityLabelUsedToIdentifySandboxedContainers: "",
      },
    });
    sandbox.stub(checkAccessCallback, "checkSideloadingCallback");
    const sideloadingNode = new SideloadingNode(eventEmitter, "token");
    const treeItem = await sideloadingNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, errorIcon);
  });
});
