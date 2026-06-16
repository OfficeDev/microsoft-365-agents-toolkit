// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IProgressHandler } from "@microsoft/teamsfx-api";
import { LocalEnvManager } from "@microsoft/teamsfx-core";
import * as sinon from "sinon";
import { createTaskStopCb } from "../../../../src/cmds/preview/commonUtils";
import cliLogger from "../../../../src/commonlib/log";
import cliTelemetry from "../../../../src/telemetry/cliTelemetry";
import { expect } from "../../utils";

describe("commonUtils createTaskStopCb npm install", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("npm install failure path", async () => {
    const getNpmInstallLogInfoStub = sandbox.stub(
      LocalEnvManager.prototype,
      "getNpmInstallLogInfo"
    );
    getNpmInstallLogInfoStub.resolves({
      cwd: "c:/tmp/app",
      exitCode: 1,
      timestamp: new Date(),
      nodeVersion: "18.0.0",
      npmVersion: "9.0.0",
      errorMessage: "install failed",
    });
    sandbox.stub(cliTelemetry, "sendTelemetryErrorEvent").callsFake(() => {});
    sandbox.stub(cliLogger, "necessaryLog").callsFake(() => {});

    const progressHandler = sandbox.createStubInstance(MockProgressHandler);
    const taskStopCallback = createTaskStopCb(progressHandler, { k: "v" });
    await taskStopCallback("npm install", false, {
      command: "command",
      success: false,
      stdout: [],
      stderr: [],
      exitCode: 1,
    });
    expect(progressHandler.end.calledOnce).to.be.true;
    expect(getNpmInstallLogInfoStub.calledOnce).to.be.true;
  });
});

class MockProgressHandler implements IProgressHandler {
  start(detail?: string): Promise<void> {
    return Promise.resolve();
  }
  next(detail?: string): Promise<void> {
    return Promise.resolve();
  }
  end(success: boolean): Promise<void> {
    return Promise.resolve();
  }
}
