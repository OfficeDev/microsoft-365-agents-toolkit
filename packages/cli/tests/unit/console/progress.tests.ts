// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import sinon from "sinon";

import Progress from "../../../src/console/progress";
import ScreenManager, { Row } from "../../../src/console/screen";
import * as Utils from "../../../src/utils";
import { expect } from "../utils";

describe("Progress", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    Progress["instances"] = [];
    Progress["rows"] = [];
    Progress["finishedRows"] = [];
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("static add", () => {
    sandbox.stub(ScreenManager, "addProgress").returns(new Row(() => "Test static add"));
    const instance = new Progress("Test static add", 3);
    Progress["add"](instance);
    expect(Progress["instances"]).deep.equals([instance]);
  });

  it("static finish", () => {
    const updateStub = sandbox.stub(Row.prototype, "update");
    const romoveCBStub = sandbox.stub(Row.prototype, "removeCB");
    const freezeStub = sandbox.stub(Row.prototype, "freeze");
    const instance = new Progress("Test static finish", 3);
    const row = new Row(() => "Test static finish");
    Progress["instances"] = [instance];
    Progress["rows"] = [row];
    Progress["finish"](instance);
    sinon.assert.calledOnce(updateStub);
    sinon.assert.calledOnce(romoveCBStub);
    sinon.assert.calledOnce(freezeStub);
    expect(Progress["instances"]).deep.equals([]);
    expect(Progress["rows"]).deep.equals([]);
    expect(Progress["finishedRows"]).deep.equals([]);
  });

  it("static finish hide", () => {
    const updateStub = sandbox.stub(Row.prototype, "update");
    const romoveCBStub = sandbox.stub(Row.prototype, "removeCB");
    const freezeStub = sandbox.stub(Row.prototype, "freeze");
    const instance = new Progress("Test static finish", 3);
    const row = new Row(() => "Test static finish");
    Progress["instances"] = [instance];
    Progress["rows"] = [row];
    Progress["finish"](instance, true);
    sinon.assert.calledOnce(updateStub);
    sinon.assert.calledOnce(romoveCBStub);
    // sinon.assert.calledOnce(freezeStub);
    expect(Progress["instances"]).deep.equals([]);
    expect(Progress["rows"]).deep.equals([]);
    expect(Progress["finishedRows"]).deep.equals([]);
  });

  it("static end", () => {
    const endStub = sandbox.stub(Progress.prototype, "end");
    const instance = new Progress("Test static end", 3);
    Progress["instances"] = [instance];
    Progress["end"](true);
    sinon.assert.calledOnce(endStub);
  });

  it("start", async () => {
    const addStub = sandbox.stub<any, any>(Progress, "add");
    const instance = new Progress("Test start", 3);
    await instance.start();
    expect(instance["status"]).equals("running");
    expect(instance["detail"]).equals(undefined);
    expect(instance["currentStep"]).equals(0);
    sinon.assert.calledOnce(addStub);
  });

  it("end", async () => {
    const finishStub = sandbox.stub<any, any>(Progress, "finish");
    const instance = new Progress("Test finish", 3);
    Progress["instances"] = [instance];
    await instance.end(true);
    expect(instance["status"]).equals("done");
    expect(instance["currentPercentage"]).equals(100);
    sinon.assert.calledOnce(finishStub);
  });

  it("next", async () => {
    const instance = new Progress("Test next", 3);
    instance["currentStep"] = 3;
    await instance.next("step 1");
    expect(instance["currentStep"]).equals(4);
    expect(instance["totalSteps"]).equals(4);
  });
  it("text", async () => {
    const instance = new Progress("Test next", 3);
    instance["currentStep"] = 3;
    await instance.text("step 1");
    expect(instance["currentStep"]).equals(3);
  });
  it("updatePercentage", () => {
    const instance = new Progress("Test next", 3);
    instance["currentPercentage"] = 0;
    instance["currentStep"] = 1;
    instance["updatePercentage"]();
    expect(instance["currentPercentage"]).gt(0).lte(100);
    instance["currentStep"] = 2;
    instance["updatePercentage"]();
    expect(instance["currentPercentage"]).gt(5).lte(100);
  });

  it("wholeMessage", () => {
    sandbox.stub(Utils, "getColorizedString").callsFake((messages) => {
      return messages.map((m) => m.content).join("");
    });
    sandbox.stub<any, any>(Progress.prototype, "updatePercentage");
    const instance = new Progress("Test next", 3);
    instance["status"] = "running";
    expect(instance.wholeMessage()).not.contains("Failed");
    expect(instance.wholeMessage()).not.contains("Done");
    instance["status"] = "done";
    expect(instance.wholeMessage()).contains("Done");
    instance["status"] = "error";
    expect(instance.wholeMessage()).contains("Failed");
  });
});
