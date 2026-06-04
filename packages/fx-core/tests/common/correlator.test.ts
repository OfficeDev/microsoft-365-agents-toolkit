// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import { Correlator } from "../../src/common/correlator";

chai.use(chaiAsPromised);

describe("Correlator", () => {
  const func = () => {};
  it("setId", async () => {
    const setedId = Correlator.setId();
    const getId = Correlator.getId();
    chai.assert.equal(setedId, getId);
  });

  it("run when id is set", () => {
    const setedId = Correlator.setId();
    Correlator.run(func);
    const getId = Correlator.getId();
    chai.assert.equal(setedId, getId);
  });

  it("run when id is not set", () => {
    Correlator.run(func);
    const getId = Correlator.getId();
    chai.assert.isDefined(getId);
  });
});
