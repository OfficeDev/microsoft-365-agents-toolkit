/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { TemplateRegistry } from "../../../src/templates/registry";

describe("TemplateRegistry", () => {
  it("should register and retrieve a template", () => {
    const registry = new TemplateRegistry();
    const descriptor = {
      id: "test-template",
      name: "Test Template",
      category: "declarative-agent" as const,
      languages: ["typescript" as const],
      description: "A test template",
      templateName: "test-template",
      scaffold: async () => {},
    };
    registry.register(descriptor as any);
    expect(registry.has("test-template")).to.be.true;
    expect(registry.get("test-template")?.name).to.equal("Test Template");
    expect(registry.size).to.equal(1);
  });

  it("should throw on duplicate registration", () => {
    const registry = new TemplateRegistry();
    const descriptor = {
      id: "dup",
      name: "Dup",
      category: "declarative-agent" as const,
      languages: [],
      description: "",
      templateName: "dup",
      scaffold: async () => {},
    };
    registry.register(descriptor as any);
    expect(() => registry.register(descriptor as any)).to.throw(/already registered/);
  });

  it("should return undefined for unregistered template", () => {
    const registry = new TemplateRegistry();
    expect(registry.get("nonexistent")).to.be.undefined;
  });

  it("should list all templates", () => {
    const registry = new TemplateRegistry();
    registry.register({
      id: "a",
      name: "A",
      category: "declarative-agent",
      languages: [],
      description: "",
      templateName: "a",
      scaffold: async () => {},
    } as any);
    registry.register({
      id: "b",
      name: "B",
      category: "declarative-agent",
      languages: [],
      description: "",
      templateName: "b",
      scaffold: async () => {},
    } as any);
    expect(registry.list()).to.have.length(2);
  });
});
