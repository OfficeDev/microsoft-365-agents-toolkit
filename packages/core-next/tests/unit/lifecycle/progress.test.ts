/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import * as sinon from "sinon";
import { createProgressAdapter, silentProgress } from "../../../src/lifecycle/progress";

function createMockUi() {
  const handler = {
    start: sinon.stub().resolves(),
    next: sinon.stub().resolves(),
    end: sinon.stub().resolves(),
  };

  const createProgressBar = sinon.stub().returns(handler);

  return {
    ui: {
      selectOption: sinon.stub(),
      selectOptions: sinon.stub(),
      inputText: sinon.stub(),
      selectFile: sinon.stub(),
      selectFiles: sinon.stub(),
      selectFolder: sinon.stub(),
      openUrl: sinon.stub(),
      showMessage: sinon.stub(),
      createProgressBar,
      confirm: sinon.stub(),
    } as any,
    createProgressBar,
    handler,
  };
}

describe("createProgressAdapter", () => {
  it("should create progress bar on onStart", async () => {
    const { ui, createProgressBar, handler } = createMockUi();
    const progress = createProgressAdapter(ui);

    await progress.onStart("provision", 3);

    expect(createProgressBar.calledOnce).to.be.true;
    expect(createProgressBar.firstCall.args[0]).to.equal("Running provision");
    expect(createProgressBar.firstCall.args[1]).to.equal(3);
    expect(handler.start.calledOnce).to.be.true;
  });

  it("should use custom title when provided", async () => {
    const { ui, createProgressBar } = createMockUi();
    const progress = createProgressAdapter(ui, "Deploying...");

    await progress.onStart("deploy", 2);

    expect(createProgressBar.firstCall.args[0]).to.equal("Deploying...");
  });

  it("should call next() on onStepStart", async () => {
    const { ui, handler } = createMockUi();
    const progress = createProgressAdapter(ui);

    await progress.onStart("provision", 3);
    await progress.onStepStart(0, "teamsApp/create");

    expect(handler.next.calledOnce).to.be.true;
    expect(handler.next.firstCall.args[0]).to.equal("teamsApp/create");
  });

  it("should call end() on onEnd", async () => {
    const { ui, handler } = createMockUi();
    const progress = createProgressAdapter(ui);

    await progress.onStart("provision", 1);
    await progress.onEnd(true);

    expect(handler.end.calledOnce).to.be.true;
    expect(handler.end.firstCall.args[0]).to.equal(true);
  });

  it("should pass false to end() on failure", async () => {
    const { ui, handler } = createMockUi();
    const progress = createProgressAdapter(ui);

    await progress.onStart("provision", 1);
    await progress.onEnd(false);

    expect(handler.end.firstCall.args[0]).to.equal(false);
  });
});

describe("silentProgress", () => {
  it("should be callable without errors", async () => {
    await silentProgress.onStart("provision", 3);
    await silentProgress.onStepStart(0, "step");
    await silentProgress.onStepComplete(0, "step", 100);
    await silentProgress.onEnd(true);
    // No error = success
  });
});
