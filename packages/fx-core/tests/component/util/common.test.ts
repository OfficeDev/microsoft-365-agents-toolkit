// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Siglud <siglud@gmail.com>
 */
import { LogProvider, UserError, err } from "@microsoft/teamsfx-api";
import { errorHandle } from "../../../src/component/utils/common";
import { BaseComponentInnerError } from "../../../src/component/error/componentError";
import * as sinon from "sinon";
import * as chai from "chai";

describe("errorHandle", () => {
  let logProvider: LogProvider;
  let errorHandler: sinon.SinonStub;

  beforeEach(() => {
    logProvider = {
      debug: (_message) => {},
    } as LogProvider;
    errorHandler = sinon.stub();
  });

  afterEach(() => {
    sinon.restore();
  });

  it("should call errorHandler if provided", async () => {
    const error = new Error("test error");
    await errorHandle(error, "testSource", logProvider, errorHandler);
    chai.expect(errorHandler.calledOnce).to.be.true;
  });

  it("should log error detail if error is BaseComponentInnerError", async () => {
    const error = new BaseComponentInnerError("source", "UserError", "test error");
    error.detail = "test detail";
    const logSpy = sinon.spy(logProvider, "debug");
    const result = await errorHandle(error, "testSource", logProvider);
    chai.expect(logSpy.calledWith(`Error occurred: ${error.detail}`)).to.be.true;
  });

  it("should return error as is if error is UserError or SystemError", async () => {
    const error = new UserError("test error", "testSource", "testMessage");
    const result = await errorHandle(error, "testSource", logProvider);
    chai.expect(result).to.deep.equal(err(error));
  });

  it("should return error as SystemError if error is not BaseComponentInnerError, UserError, or SystemError", async () => {
    const error = new Error("test error");
    const result = await errorHandle(error, "testSource", logProvider);
    chai.expect(result.isErr()).to.be.true;
    chai.expect((result.mapErr((e) => e.source) as any).error).to.equal("testSource");
  });
});
