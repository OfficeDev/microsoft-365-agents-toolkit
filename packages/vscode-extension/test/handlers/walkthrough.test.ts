import * as vscode from "vscode";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import {
  createProjectFromWalkthroughHandler,
  openBuildIntelligentAppsWalkthroughHandler,
  walkthroughDeps,
} from "../../src/handlers/walkthrough";
import * as sinon from "sinon";
import { expect } from "chai";
import { Inputs, ok } from "@microsoft/teamsfx-api";

describe("walkthrough", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("create proejct from walkthrough", async () => {
    const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const inputs = {} as Inputs;
    const systemInputsStub = sandbox.stub(walkthroughDeps, "getSystemInputs").callsFake(() => {
      return inputs;
    });

    const runCommandStub = sandbox.stub(walkthroughDeps, "runCommand").resolves(ok(null));

    await createProjectFromWalkthroughHandler([
      "walkthrough",
      { "project-type": "custom-copilot-type", capabilities: "cutsom-copilot-agent" },
    ]);

    sandbox.assert.calledOnce(sendTelemetryEventStub);
    sandbox.assert.calledOnce(systemInputsStub);
    sandbox.assert.calledOnce(runCommandStub);

    expect(Object.keys(inputs)).lengthOf(2);
  });

  it("build intelligent apps", async () => {
    const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const executeCommands = sandbox.stub(vscode.commands, "executeCommand");

    await openBuildIntelligentAppsWalkthroughHandler();
    sandbox.assert.calledOnce(sendTelemetryEventStub);
    sandbox.assert.calledOnceWithExactly(
      executeCommands,
      "workbench.action.openWalkthrough",
      "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps"
    );
  });
});
