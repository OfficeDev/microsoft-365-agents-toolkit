// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IProgressHandler } from "@microsoft/teamsfx-api";
import * as sinon from "sinon";
import { createTaskStartCb } from "../../../../src/cmds/preview/commonUtils";
import { expect } from "../../utils";

describe("commonUtils createTaskStartCb", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("happy path", async () => {
    const progressHandler = sandbox.createStubInstance(MockProgressHandler);
    const taskStartCallback = createTaskStartCb(progressHandler, "start message");
    await taskStartCallback("start", true);
    expect(progressHandler.start.calledOnce).to.be.true;
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
