/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { DriverRegistry, driverRegistry } from "../../../../src/drivers/registry";
import { builtinDrivers, registerBuiltinDrivers } from "../../../../src/drivers/builtin";

describe("registerBuiltinDrivers", () => {
  it("should export the expected number of built-in drivers", () => {
    expect(builtinDrivers).to.have.lengthOf(22);
  });

  it("should have correct driver IDs", () => {
    const ids = builtinDrivers.map((d) => d.id);
    expect(ids).to.include("file/createOrUpdateEnvironmentFile");
    expect(ids).to.include("file/createOrUpdateJsonFile");
    expect(ids).to.include("script");
  });

  it("should register all drivers into a fresh registry without errors", () => {
    // Use a fresh registry to avoid conflicts with other tests
    const registry = new DriverRegistry();
    for (const driver of builtinDrivers) {
      registry.register(driver);
    }
    expect(registry.size).to.equal(22);
    expect(registry.has("file/createOrUpdateEnvironmentFile")).to.be.true;
    expect(registry.has("file/createOrUpdateJsonFile")).to.be.true;
    expect(registry.has("script")).to.be.true;
    expect(registry.has("teamsApp/create")).to.be.true;
    expect(registry.has("teamsApp/configure")).to.be.true;
    expect(registry.has("teamsApp/publishAppPackage")).to.be.true;
    expect(registry.has("aadApp/create")).to.be.true;
    expect(registry.has("aadApp/update")).to.be.true;
    expect(registry.has("botAadApp/create")).to.be.true;
    expect(registry.has("botFramework/create")).to.be.true;
  });

  it("should be idempotent (safe to call twice)", () => {
    // registerBuiltinDrivers uses has() check before register()
    // so calling it twice should not throw
    registerBuiltinDrivers();
    registerBuiltinDrivers();
    expect(driverRegistry.has("file/createOrUpdateEnvironmentFile")).to.be.true;
  });
});
