import * as chai from "chai";
import * as vscode from "vscode";
import { vi } from "vitest";

import { CommandStatus, TreeViewCommand } from "../../src/treeview/treeViewCommand";
import * as localizeUtils from "../../src/utils/localizeUtils";

describe("TreeViewCommand", () => {
  it("setStatus", async () => {
    vi.spyOn(localizeUtils, "localize").mockImplementation((key: string) => {
      if (key === "teamstoolkit.commandsTreeViewProvider.key.running") {
        return "test running";
      } else if (key === "teamstoolkit.commandsTreeViewProvider.key.blockTooltip") {
        return "blocked tooltip";
      }
      return "";
    });

    const command = new TreeViewCommand("label", "tooltip", "command", "key");

    command.setStatus(CommandStatus.Ready);
    chai.assert.equal(command.label, "label");
    chai.assert.equal(command.tooltip, "tooltip");

    command.setStatus(CommandStatus.Running);
    chai.assert.equal(command.label, "test running");
    chai.assert.deepEqual(command.iconPath, new vscode.ThemeIcon("loading~spin"));

    command.setStatus(CommandStatus.Blocked, command.getBlockingTooltip());
    chai.assert.equal(command.tooltip, "blocked tooltip");
  });
});
