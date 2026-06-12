import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";

import {
  CommandStatus,
  TreeViewCommand,
  treeViewCommandDeps,
} from "../../src/treeview/treeViewCommand";

describe("TreeViewCommand", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("setStatus", async () => {
    sandbox.stub(treeViewCommandDeps, "localize").callsFake((key: string) => {
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
