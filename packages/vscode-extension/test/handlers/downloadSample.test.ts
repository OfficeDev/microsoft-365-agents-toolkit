import { err, Inputs, Platform, Stage, SystemError } from "@microsoft/teamsfx-api";
import * as projectSettingsHelper from "@microsoft/teamsfx-core/build/src/common/projectSettingsHelper";
import * as chai from "chai";
import * as sinon from "sinon";
import * as vscode from "vscode";
import * as globalVariables from "../../src/globalVariables";
import { downloadSample, downloadSampleApp } from "../../src/handlers/downloadSample";
import { ExtTelemetry } from "../../src/telemetry/extTelemetry";
import { TelemetryTriggerFrom } from "../../src/telemetry/extTelemetryEvents";
import { MockCore } from "../mocks/mockCore";

describe("downloadSampleApp", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {});

  afterEach(() => {
    sandbox.restore();
  });

  it("happy path", async () => {
    sandbox.stub(globalVariables, "checkIsSPFx").returns(false);
    sandbox.stub(vscode.commands, "executeCommand");
    sandbox.stub(globalVariables, "core").value(new MockCore());
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const errorEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    const createProject = sandbox.spy(globalVariables.core, "createSampleProject");

    await downloadSampleApp(TelemetryTriggerFrom.CopilotChat, "test");

    chai.assert.isTrue(createProject.calledOnce);
    chai.assert.isTrue(errorEventStub.notCalled);
  });

  it("has error", async () => {
    sandbox.stub(globalVariables, "checkIsSPFx").returns(false);
    sandbox.stub(vscode.commands, "executeCommand");
    sandbox.stub(globalVariables, "core").value(new MockCore());
    sandbox.stub(ExtTelemetry, "sendTelemetryEvent");
    const errorEventStub = sandbox.stub(ExtTelemetry, "sendTelemetryErrorEvent");
    sandbox.stub(projectSettingsHelper, "isValidOfficeAddInProject").returns(false);
    sandbox
      .stub(globalVariables.core, "createSampleProject")
      .rejects(err(new Error("Cannot get user login information")));

    await downloadSampleApp(TelemetryTriggerFrom.CopilotChat, "test");

    chai.assert.isTrue(errorEventStub.calledOnce);
  });
});

describe("DownloadSample", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("downloadSample", async () => {
    const inputs: Inputs = {
      scratch: "no",
      platform: Platform.VSCode,
    };
    sandbox.stub(globalVariables, "core").value(new MockCore());
    const createProject = sandbox.spy(globalVariables.core, "createSampleProject");

    await downloadSample(inputs);

    inputs.stage = Stage.create;
    chai.assert.isTrue(createProject.calledOnceWith(inputs));
  });

  it("downloadSample - error", async () => {
    const inputs: Inputs = {
      scratch: "no",
      platform: Platform.VSCode,
    };
    sandbox.stub(globalVariables, "core").value(new MockCore());
    const showErrorMessageStub = sandbox.stub(vscode.window, "showErrorMessage");
    const createProject = sandbox
      .stub(globalVariables.core, "createSampleProject")
      .rejects(err(new Error("Cannot get user login information")));

    await downloadSample(inputs);

    inputs.stage = Stage.create;
    chai.assert.isTrue(createProject.calledOnceWith(inputs));
    chai.assert.isTrue(showErrorMessageStub.calledOnce);
  });

  it("downloadSample - LoginFailureError", async () => {
    const inputs: Inputs = {
      scratch: "no",
      platform: Platform.VSCode,
    };
    sandbox.stub(globalVariables, "core").value(new MockCore());
    const showErrorMessageStub = sandbox.stub(vscode.window, "showErrorMessage");
    const createProject = sandbox
      .stub(globalVariables.core, "createProject")
      .resolves(err(new SystemError("test", "test", "Cannot get user login information")));

    await downloadSample(inputs);
  });
});
