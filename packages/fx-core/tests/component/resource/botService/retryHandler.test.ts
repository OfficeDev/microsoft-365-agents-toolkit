// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as chai from "chai";
import { RetryHandler } from "../../../../src/component/resource/botService/retryHandler";
import { Messages } from "./messages";
import * as sinon from "sinon";
import { Retry } from "../../../../src/component/resource/botService/constants";

describe("Test retry handler", () => {
  const maxTry = 3;
  const sandbox = sinon.createSandbox();
  beforeEach(() => {
    sandbox.stub(Retry, "BACKOFF_TIME_MS").value(0);
    sandbox.stub(Retry, "RETRY_TIMES").value(maxTry);
  });
  afterEach(() => {
    sandbox.restore();
  });
  it("fn resolve", async () => {
    const res = await RetryHandler.Retry(async () => true);
    chai.assert.isTrue(res);
  });

  it("fn reject and retry", async () => {
    const errorMessage = "fn rejects";
    let count = 0;
    try {
      await RetryHandler.Retry(async () => {
        ++count;
        throw new Error(errorMessage);
      });
    } catch (e) {
      chai.assert.equal(e.message, errorMessage);
      chai.assert.equal(count, maxTry);
      return;
    }
    chai.assert.fail(Messages.ShouldNotReachHere);
  });
});
