import * as chai from "chai";
import * as sinon from "sinon";

import { TreeViewCommand } from "../../src/treeview/treeViewCommand";
import { CommandsTreeViewProvider } from "../../src/treeview/commandsTreeViewProvider";

describe("CommandsTreeViewProvider", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("getCommands", async () => {
    const provider = new CommandsTreeViewProvider([new TreeViewCommand("test", "test")]);
    const commands = provider.getCommands();

    chai.assert.equal(commands.length, 1);
    chai.assert.equal(commands[0].label, "test");
  });

  it("getTreeItem", async () => {
    const provider = new CommandsTreeViewProvider([]);
    const command = provider.getTreeItem(new TreeViewCommand("test", "test"));

    chai.assert.equal(command.label, "test");
  });

  it("getChildren", async () => {
    const provider = new CommandsTreeViewProvider([new TreeViewCommand("test", "test")]);
    const commands = await provider.getChildren();

    chai.assert.equal(commands.length, 1);
    chai.assert.equal(commands[0].label, "test");
  });
});
