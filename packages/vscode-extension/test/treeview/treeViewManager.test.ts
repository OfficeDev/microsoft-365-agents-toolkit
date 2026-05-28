import { TeamsAppManifest, ok } from "@microsoft/teamsfx-api";
import { featureFlagManager, manifestUtils } from "@microsoft/teamsfx-core";
import * as featureFlags from "@microsoft/teamsfx-core/build/common/featureFlags";
import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as globalVariables from "../../src/globalVariables";
import { CommandsTreeViewProvider } from "../../src/treeview/commandsTreeViewProvider";
import treeViewManager from "../../src/treeview/treeViewManager";

describe("TreeViewManager", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("registerTreeViews", () => {
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);
    chai.assert.isDefined(treeViewManager.getTreeView("teamsfx-accounts"));

    const lifecycleTreeView = treeViewManager.getTreeView("teamsfx-lifecycle");
    chai.assert.isDefined(lifecycleTreeView);
    chai.assert.equal((lifecycleTreeView as any).commands.length, 3);
    chai.assert.equal((lifecycleTreeView as any).commands[0].commandId, "fx-extension.provision");
  });

  it("Development Treeview", () => {
    sandbox.stub(globalVariables, "context").value({ extensionPath: "" });
    sandbox.stub(globalVariables, "isSPFxProject").value(false);
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    const developmentTreeview = treeViewManager.getTreeView("teamsfx-development");
    chai.assert.isDefined(developmentTreeview);
    chai.assert.equal((developmentTreeview as any).commands.length, 4);
  });

  it("Development Treeview when ChatParticipant is enabled", () => {
    sandbox.stub(globalVariables, "context").value({ extensionPath: "" });
    sandbox.stub(globalVariables, "isSPFxProject").value(false);
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    const developmentTreeview = treeViewManager.getTreeView("teamsfx-development");
    chai.assert.isDefined(developmentTreeview);
    chai.assert.equal((developmentTreeview as any).commands.length, 5);
  });

  it("setRunningCommand", () => {
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);
    const command = (treeViewManager as any).commandMap.get("fx-extension.create");
    const setStatusStub = sandbox.stub(command, "setStatus");
    treeViewManager.setRunningCommand("fx-extension.create", ["fx-extension.openSamples"]);

    chai.assert.equal(setStatusStub.callCount, 1);

    treeViewManager.restoreRunningCommand(["fx-extension.openSamples"]);
    chai.assert.equal(setStatusStub.callCount, 2);
  });

  it("updateDevelopmentTreeView", () => {
    sandbox.stub(globalVariables, "isSPFxProject").value(false);
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    const developmentTreeviewProvider = treeViewManager.getTreeView(
      "teamsfx-development"
    ) as CommandsTreeViewProvider;

    const commands = developmentTreeviewProvider.getCommands();
    chai.assert.equal(commands.length, 4);

    sandbox.stub(globalVariables, "isSPFxProject").value(true);
    treeViewManager.updateDevelopmentTreeView();

    chai.assert.equal(commands.length, 5);
  });

  it("updateTreeViewsByContent if remove project related commands", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value("");
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
    sandbox.stub(manifestUtils, "readAppManifest").resolves(ok({} as TeamsAppManifest));
    sandbox.stub(manifestUtils, "getCapabilities").returns(["tab"]);
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    const developmentTreeviewProvider = treeViewManager.getTreeView(
      "teamsfx-development"
    ) as CommandsTreeViewProvider;

    const utilityTreeviewProvider = treeViewManager.getTreeView(
      "teamsfx-utility"
    ) as CommandsTreeViewProvider;

    await treeViewManager.updateTreeViewsByContent(true);
    const developmentCommands = developmentTreeviewProvider.getCommands();
    const utilityCommands = utilityTreeviewProvider.getCommands();
    chai.assert.equal(developmentCommands.length, 3);
    chai.assert.equal(utilityCommands.length, 3);
  });

  it("updateTreeViewsByContent if remove project related commands when ChatParticipant is enabled", async () => {
    sandbox.stub(globalVariables, "workspaceUri").value("");
    sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
    sandbox.stub(manifestUtils, "readAppManifest").resolves(ok({} as TeamsAppManifest));
    sandbox.stub(manifestUtils, "getCapabilities").returns(["tab"]);
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    const developmentTreeviewProvider = treeViewManager.getTreeView(
      "teamsfx-development"
    ) as CommandsTreeViewProvider;

    const developmentCommands = developmentTreeviewProvider.getCommands();
    const utilityTreeviewProvider = treeViewManager.getTreeView(
      "teamsfx-utility"
    ) as CommandsTreeViewProvider;
    const utilityCommands = utilityTreeviewProvider.getCommands();
    chai.assert.equal(developmentCommands.length, 5);
    chai.assert.equal(utilityCommands.length, 3);

    await treeViewManager.updateTreeViewsByContent(true);
    chai.assert.equal(developmentCommands.length, 4);
    chai.assert.equal(utilityCommands.length, 3);
  });
});
