/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { AxiosResponse } from "axios";
import { sendWithRetry } from "../../../src/http/retry";

describe("sendWithRetry", () => {
  it("does not retry 4xx client errors (throws immediately)", async () => {
    let callCount = 0;
    const error400 = Object.assign(new Error("Bad Request"), {
      response: { status: 400, data: {} },
    });

    try {
      await sendWithRetry(() => {
        callCount++;
        throw error400;
      });
      expect.fail("should have thrown");
    } catch (e) {
      expect(e).to.equal(error400);
    }
    expect(callCount).to.equal(1); // no retries
  });

  it("retries 429 Too Many Requests", async () => {
    let callCount = 0;
    const error429 = Object.assign(new Error("Too Many Requests"), {
      response: { status: 429 },
    });

    try {
      await sendWithRetry(
        () => {
          callCount++;
          throw error429;
        },
        1 // 1 retry = 2 total attempts
      );
      expect.fail("should have thrown");
    } catch {
      // expected
    }
    expect(callCount).to.equal(2); // initial + 1 retry
  });

  it("retries 412 Precondition Failed (transient Graph API condition)", async () => {
    let callCount = 0;
    const error412 = Object.assign(new Error("Precondition Failed"), {
      response: { status: 412 },
    });

    try {
      await sendWithRetry(() => {
        callCount++;
        throw error412;
      }, 1);
      expect.fail("should have thrown");
    } catch {
      // expected
    }
    expect(callCount).to.equal(2); // initial + 1 retry
  });

  it("retries network errors without response", async () => {
    let callCount = 0;
    try {
      await sendWithRetry(() => {
        callCount++;
        throw new Error("ECONNRESET");
      }, 1);
      expect.fail("should have thrown");
    } catch {
      // expected
    }
    expect(callCount).to.equal(2);
  });

  it("returns immediately on 2xx success", async () => {
    let callCount = 0;
    const result = await sendWithRetry(() => {
      callCount++;
      return Promise.resolve({ status: 200, data: "ok" } as AxiosResponse);
    });
    expect(callCount).to.equal(1);
    expect(result.data).to.equal("ok");
  });
});
