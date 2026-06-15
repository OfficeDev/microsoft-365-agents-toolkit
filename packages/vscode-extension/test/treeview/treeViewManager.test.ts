import { TeamsAppManifest, ok } from "@microsoft/teamsfx-api";
import { featureFlagManager, manifestUtils } from "@microsoft/teamsfx-core";
import * as chai from "chai";
import * as vscode from "vscode";
import * as globalVariables from "../../src/globalVariables";
import { CommandsTreeViewProvider } from "../../src/treeview/commandsTreeViewProvider";
import treeViewManager, { treeViewManagerDeps } from "../../src/treeview/treeViewManager";
import * as commonUtils from "../../src/utils/commonUtils";
import { vi } from "vitest";
import { mockValue } from "../mocks/vitestMockUtils";

describe("TreeViewManager", () => {
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
    mockValue(globalVariables, "context", { extensionPath: "" });
    mockValue(globalVariables, "isSPFxProject", false);
    vi.spyOn(featureFlagManager, "getBooleanValue").mockReturnValue(false);
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    const developmentTreeview = treeViewManager.getTreeView("teamsfx-development");
    chai.assert.isDefined(developmentTreeview);
    chai.assert.equal((developmentTreeview as any).commands.length, 4);
  });

  it("Development Treeview when HideGitHubCopilotPreviewTag is enabled", () => {
    mockValue(globalVariables, "context", { extensionPath: "" });
    mockValue(globalVariables, "isSPFxProject", false);
    vi.spyOn(featureFlagManager, "getBooleanValue").mockReturnValue(true);
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    const developmentTreeview = treeViewManager.getTreeView("teamsfx-development");
    chai.assert.isDefined(developmentTreeview);
    chai.assert.equal((developmentTreeview as any).commands.length, 5);
  });

  it("Development Treeview when enable extend MetaOS to DA", () => {
    mockValue(globalVariables, "isMetaOSAddinProject", true);
    mockValue(globalVariables, "isDeclarativeCopilotApp", false);
    vi.spyOn(featureFlagManager, "getBooleanValue").mockReturnValue(true);

    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    const developmentTreeview = treeViewManager.getTreeView("teamsfx-development");
    chai.assert.isDefined(developmentTreeview);
    chai.assert.equal((developmentTreeview as any).commands.length, 6);
  });

  it("setRunningCommand", () => {
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);
    const command = (treeViewManager as any).commandMap.get("fx-extension.create");
    const setStatusStub = vi.spyOn(command, "setStatus");
    treeViewManager.setRunningCommand("fx-extension.create", ["fx-extension.openSamples"]);

    chai.assert.equal(setStatusStub.callCount, 1);

    treeViewManager.restoreRunningCommand(["fx-extension.openSamples"]);
    chai.assert.equal(setStatusStub.callCount, 2);
  });

  it("updateDevelopmentTreeView", () => {
    mockValue(globalVariables, "isSPFxProject", false);
    vi.spyOn(featureFlagManager, "getBooleanValue").mockReturnValue(false);
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    const developmentTreeviewProvider = treeViewManager.getTreeView(
      "teamsfx-development"
    ) as CommandsTreeViewProvider;

    const commands = developmentTreeviewProvider.getCommands();
    chai.assert.equal(commands.length, 4);

    mockValue(globalVariables, "isSPFxProject", true);
    treeViewManager.updateDevelopmentTreeView();

    chai.assert.equal(commands.length, 5);
  });

  it("updateTreeViewsByContent if remove project related commands", async () => {
    mockValue(globalVariables, "workspaceUri", "");
    vi.spyOn(featureFlagManager, "getBooleanValue").mockReturnValue(false);
    vi.spyOn(manifestUtils, "readAppManifest").mockResolvedValue(ok({} as TeamsAppManifest));
    vi.spyOn(manifestUtils, "getCapabilities").mockReturnValue(["tab"]);
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

  it("updateTreeViewsByContent if remove project related commands when HideGitHubCopilotPreviewTag is enabled", async () => {
    mockValue(globalVariables, "workspaceUri", "");
    vi.spyOn(featureFlagManager, "getBooleanValue").mockReturnValue(true);
    vi.spyOn(manifestUtils, "readAppManifest").mockResolvedValue(ok({} as TeamsAppManifest));
    vi.spyOn(manifestUtils, "getCapabilities").mockReturnValue(["tab"]);
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

  it("updateTreeViewsByContent when adaptiveCardInWorkspace is enabled", async () => {
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    const developmentTreeviewProvider = treeViewManager.getTreeView(
      "teamsfx-development"
    ) as CommandsTreeViewProvider;

    const commands = developmentTreeviewProvider.getCommands();
    chai.assert.equal(commands.length, 4);

    vi.spyOn(treeViewManagerDeps, "hasAdaptiveCardInWorkspace").mockReturnValue(
      Promise.resolve(true)
    );
    await treeViewManager.updateTreeViewsByContent();

    chai.assert.equal(commands.length, 5);
  });

  it("Development Treeview when Add knowledge is enabled", () => {
    vi.spyOn(featureFlagManager, "getBooleanValue").mockReturnValue(true);
    mockValue(globalVariables, "isDeclarativeCopilotApp", true);
    treeViewManager.registerTreeViews({
      subscriptions: [],
    } as unknown as vscode.ExtensionContext);

    const developmentTreeview = treeViewManager.getTreeView("teamsfx-development");
    chai.assert.isDefined(developmentTreeview);
    chai.assert.equal((developmentTreeview as any).commands.length, 9);
  });
});
