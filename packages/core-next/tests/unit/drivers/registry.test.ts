/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { DriverRegistry } from "../../../src/drivers/registry";

describe("DriverRegistry", () => {
  it("should register and retrieve a driver", () => {
    const registry = new DriverRegistry();
    const descriptor = {
      id: "test-driver",
      name: "Test Driver",
      description: "A test driver",
      execute: async () => {},
    };
    registry.register(descriptor as any);
    expect(registry.has("test-driver")).to.be.true;
    expect(registry.get("test-driver")?.name).to.equal("Test Driver");
    expect(registry.size).to.equal(1);
  });

  it("should throw on duplicate registration", () => {
    const registry = new DriverRegistry();
    const descriptor = { id: "dup", name: "Dup", description: "", execute: async () => {} };
    registry.register(descriptor as any);
    expect(() => registry.register(descriptor as any)).to.throw(/already registered/);
  });

  it("should return undefined for unregistered driver", () => {
    const registry = new DriverRegistry();
    expect(registry.get("nonexistent")).to.be.undefined;
  });
});
