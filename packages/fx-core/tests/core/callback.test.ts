// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CoreCallbackEvent, FxError, UserError } from "@microsoft/teamsfx-api";
import { CallbackRegistry } from "../../src/core/callback";
import { expect } from "chai";

describe("Core event callback tests", async () => {
  it("basic operation should work well", async () => {
    const event = CoreCallbackEvent.lock;
    expect(CallbackRegistry.has(event)).to.be.false;

    let e = false;
    let d = [];

    const cb = (name: string, fe?: FxError, fd?: string[]) => {
      if (fe) {
        e = true;
      }
      if (fd) {
        d = fd;
      }
    };
    CallbackRegistry.set(event, cb);
    expect(CallbackRegistry.has(event)).to.be.true;
    expect(CallbackRegistry.has(CoreCallbackEvent.unlock)).to.be.false;

    const funcs = CallbackRegistry.get(CoreCallbackEvent.lock);
    expect(funcs.length).eql(1);

    funcs[0]("", new UserError({}), ["1", "2"]);
    expect(e).is.true;
    expect(d.length).eql(2);

    CallbackRegistry.refresh();
    expect(CallbackRegistry.has(CoreCallbackEvent.lock)).to.be.false;
    expect(CallbackRegistry.has(CoreCallbackEvent.unlock)).to.be.false;
  });
});
