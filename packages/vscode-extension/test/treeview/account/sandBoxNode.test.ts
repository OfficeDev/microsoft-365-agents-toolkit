import * as vscode from "vscode";
import { expect } from "chai";
import { createSandbox, SinonSandbox, SinonStub } from "sinon";
import { SandboxNode } from "../../../src/treeview/account/sandBoxNode";
import { passIcon, warningIcon } from "../../../src/treeview/account/common";
import { MockedM365Provider } from "@microsoft/teamsfx-core/tests/core/utils";
import { GraphClient } from "@microsoft/teamsfx-core";
import { localize } from "../../../src/utils/localizeUtils";

describe("SandboxNode", () => {
  let sandbox: SinonSandbox;
  let eventEmitterStub: vscode.EventEmitter<any>;
  let sandboxNode: SandboxNode;
  const mockTokenProvider = new MockedM365Provider();

  beforeEach(() => {
    sandbox = createSandbox();

    const mockEventEmitter = { fire: sandbox.stub() } as any;
    sandboxNode = new SandboxNode(mockEventEmitter, "fake-token");
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("should show sandbox enabled info when sandbox is enabled", async () => {
    sandbox.stub(GraphClient.prototype, "GetTeamsAppSettingsAsync").resolves({
      sandboxingConfiguration: {
        isSideloadingEnabled: false,
        sensitivityLabelUsedToIdentifySandboxedContainers: "0fcfd0ff-1cda-407e-bc2b-a350307bd1d5",
      },
    });

    await sandboxNode.getTreeItem();

    expect(sandboxNode.label).to.equal(localize("teamstoolkit.accountTree.sandboxedTeamEnabled"));
    expect(sandboxNode.iconPath).to.equal(passIcon);
  });

  it("should show sandbox disabled info when sandbox is not enabled", async () => {
    sandbox.stub(GraphClient.prototype, "GetTeamsAppSettingsAsync").resolves({
      sandboxingConfiguration: {
        isSideloadingEnabled: false,
        sensitivityLabelUsedToIdentifySandboxedContainers: "",
      },
    });

    const mockEventEmitter = { fire: sandbox.stub() } as any;
    sandboxNode = new SandboxNode(mockEventEmitter, "fake-token");
    await sandboxNode.getTreeItem();

    expect(sandboxNode.label).to.equal(localize("teamstoolkit.accountTree.sandboxedTeamDisabled"));
    expect(sandboxNode.iconPath).to.equal(warningIcon);
  });

  it("should return null for getChildren", () => {
    const result = sandboxNode.getChildren();
    expect(result).to.be.null;
  });
});
