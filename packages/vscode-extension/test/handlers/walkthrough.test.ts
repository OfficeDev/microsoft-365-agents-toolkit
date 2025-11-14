import { Inputs, ok } from "@microsoft/teamsfx-api";
import { expect } from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as handlers from "../../src/handlers/sharedOpts";
import {
  createProjectFromWalkthroughHandler,
  openBuildIntelligentAppsWalkthroughHandler,
} from "../../src/handlers/walkthrough";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import * as environmentUtils from "../../src/utils/systemEnvUtils";

describe("walkthrough", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("create proejct from walkthrough", async () => {
    const sendTelemetryEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const inputs = {} as Inputs;
    const systemInputsStub = sandbox.stub(environmentUtils, "getSystemInputs").callsFake(() => {
      return inputs;
    });

    const runCommandStub = sandbox.stub(handlers, "runCommand").resolves(ok(null));

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
    sandbox.assert.calledOnce(executeCommands);
    const args = executeCommands.getCall(0).args;
    expect(args[0]).equals("workbench.action.openWalkthrough");
    expect([
      "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentApps",
      "TeamsDevApp.ms-teams-vscode-extension#buildIntelligentAppsWithChat",
    ]).includes(args[1]);
  });
});
