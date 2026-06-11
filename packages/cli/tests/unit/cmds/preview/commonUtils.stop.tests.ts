// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IProgressHandler } from "@microsoft/teamsfx-api";
import * as sinon from "sinon";
import { createTaskStopCb } from "../../../../src/cmds/preview/commonUtils";
import { expect } from "../../utils";

describe("commonUtils createTaskStopCb", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("happy path", async () => {
    const progressHandler = sandbox.createStubInstance(MockProgressHandler);
    const taskStopCallback = createTaskStopCb(progressHandler);
    await taskStopCallback("stop", true, {
      command: "command",
      success: true,
      stdout: [],
      stderr: [],
      exitCode: null,
    });
    expect(progressHandler.end.calledOnce).to.be.true;
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
