import * as chai from "chai";
import * as vscode from "vscode";
import * as tools from "@microsoft/teamsfx-core/build/common/tools";
import { errorIcon, infoIcon, passIcon } from "../../../src/treeview/account/common";
import { SideloadingNode } from "../../../src/treeview/account/sideloadingNode";
import { sideloadingNodeDeps } from "../../../src/treeview/account/sideloadingNode";
import { DynamicNode } from "../../../src/treeview/dynamicNode";
import * as checkAccessCallback from "../../../src/handlers/accounts/checkAccessCallback";
import { featureFlagManager, GraphClient } from "@microsoft/teamsfx-core";
import { vi } from "vitest";
import { mockValue } from "../../mocks/vitestMockUtils";

describe("sideloadingNode", () => {
  const eventEmitter = new vscode.EventEmitter<DynamicNode | undefined | void>();

  it("getTreeItem with empty string", async () => {
    const sideloadingNode = new SideloadingNode(eventEmitter, "");
    const treeItem = await sideloadingNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, infoIcon);
  });

  it("getTreeItem with invalid token", async () => {
    vi.spyOn(sideloadingNodeDeps, "getSideloadingStatus").mockReturnValue(Promise.resolve(false));
    mockValue(
      sideloadingNodeDeps,
      "checkSideloadingCallback",
      vi.fn().mockResolvedValue(undefined) as never
    );
    const sideloadingNode = new SideloadingNode(eventEmitter, "token");
    const treeItem = await sideloadingNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, errorIcon);
  });

  it("getTreeItem with valid token", async () => {
    vi.spyOn(sideloadingNodeDeps, "getSideloadingStatus").mockReturnValue(Promise.resolve(true));
    const sideloadingNode = new SideloadingNode(eventEmitter, "token");
    const treeItem = await sideloadingNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, passIcon);
  });

  it("getChildren", () => {
    const sideloadingNode = new SideloadingNode(eventEmitter, "token");
    chai.assert.isNull(sideloadingNode.getChildren());
  });

  it("Check sandbox permission", async () => {
    vi.spyOn(sideloadingNodeDeps, "getSideloadingStatus").mockReturnValue(Promise.resolve(false));
    vi.spyOn(sideloadingNodeDeps, "getBooleanValue").mockReturnValue(true);
    vi.spyOn(GraphClient.prototype, "GetTeamsAppSettingsAsync").mockResolvedValue({
      sandboxingConfiguration: {
        isSideloadingEnabled: false,
        sensitivityLabelUsedToIdentifySandboxedContainers: "0fcfd0ff-1cda-407e-bc2b-a350307bd1d5",
      },
    });
    mockValue(
      sideloadingNodeDeps,
      "checkSandboxCallback",
      vi.fn().mockResolvedValue(undefined) as never
    );
    const sideloadingNode = new SideloadingNode(eventEmitter, "token");
    const treeItem = await sideloadingNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, errorIcon);
  });

  it("Check sandbox permission - disabled", async () => {
    vi.spyOn(sideloadingNodeDeps, "getSideloadingStatus").mockReturnValue(Promise.resolve(false));
    vi.spyOn(sideloadingNodeDeps, "getBooleanValue").mockReturnValue(true);
    vi.spyOn(GraphClient.prototype, "GetTeamsAppSettingsAsync").mockResolvedValue({
      sandboxingConfiguration: {
        isSideloadingEnabled: false,
        sensitivityLabelUsedToIdentifySandboxedContainers: "",
      },
    });
    mockValue(
      sideloadingNodeDeps,
      "checkSideloadingCallback",
      vi.fn().mockResolvedValue(undefined) as never
    );
    const sideloadingNode = new SideloadingNode(eventEmitter, "token");
    const treeItem = await sideloadingNode.getTreeItem();

    chai.assert.equal(treeItem.iconPath, errorIcon);
  });
});
