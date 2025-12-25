import { assert } from "chai";
import "mocha";
import fs from "fs-extra";
import { createSandbox } from "sinon";
import { PluginManifestWrapper, RuntimeType } from "../../src/wrappers/PluginManifestWrapper";

describe("PluginManifestWrapper", () => {
  const sandbox = createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("create", () => {
    it("should create a new plugin manifest with required fields", () => {
      const plugin = PluginManifestWrapper.create({
        schemaVersion: "v2.4",
        nameForHuman: "Test Plugin",
        descriptionForHuman: "A test plugin",
      });

      assert.equal(plugin.schemaVersion, "v2.4");
      assert.equal(plugin.nameForHuman, "Test Plugin");
      assert.equal(plugin.descriptionForHuman, "A test plugin");
      assert.equal(plugin.namespace, "Test_Plugin");
      assert.isFalse(plugin.isDirty);
    });

    it("should use custom namespace if provided", () => {
      const plugin = PluginManifestWrapper.create({
        schemaVersion: "v2.4",
        nameForHuman: "Test Plugin",
        descriptionForHuman: "A test plugin",
        namespace: "custom_ns",
      });

      assert.equal(plugin.namespace, "custom_ns");
    });
  });

  describe("fromJSON", () => {
    it("should create wrapper from JSON string", () => {
      const json = JSON.stringify({
        schema_version: "v2.4",
        name_for_human: "JSON Plugin",
        description_for_human: "From JSON",
        namespace: "json_plugin",
      });

      const plugin = PluginManifestWrapper.fromJSON(json);

      assert.equal(plugin.schemaVersion, "v2.4");
      assert.equal(plugin.nameForHuman, "JSON Plugin");
      assert.isUndefined(plugin.filePath);
    });
  });

  describe("fluent setters", () => {
    it("should set properties and mark dirty", () => {
      const plugin = PluginManifestWrapper.create({
        schemaVersion: "v2.4",
        nameForHuman: "Test",
        descriptionForHuman: "Test",
      });

      assert.isFalse(plugin.isDirty);

      plugin.setNameForHuman("New Name");
      assert.equal(plugin.nameForHuman, "New Name");
      assert.isTrue(plugin.isDirty);
    });

    it("should support method chaining", () => {
      const plugin = PluginManifestWrapper.create({
        schemaVersion: "v2.4",
        nameForHuman: "Test",
        descriptionForHuman: "Test",
      });

      const result = plugin
        .setNameForHuman("Chained Name")
        .setDescriptionForHuman("Chained Description")
        .setDescriptionForModel("Model description")
        .setNamespace("chained_ns");

      assert.strictEqual(result, plugin);
      assert.equal(plugin.nameForHuman, "Chained Name");
      assert.equal(plugin.descriptionForHuman, "Chained Description");
      assert.equal(plugin.namespace, "chained_ns");
    });
  });

  describe("function operations", () => {
    it("should add functions", () => {
      const plugin = PluginManifestWrapper.create({
        schemaVersion: "v2.4",
        nameForHuman: "Test",
        descriptionForHuman: "Test",
      });

      plugin.addFunction("getUser", "Get user by ID");

      assert.isTrue(plugin.hasFunction("getUser"));
      assert.equal(plugin.functions.length, 1);
      assert.equal(plugin.getFunction("getUser")?.description, "Get user by ID");
    });

    it("should not add duplicate functions", () => {
      const plugin = PluginManifestWrapper.create({
        schemaVersion: "v2.4",
        nameForHuman: "Test",
        descriptionForHuman: "Test",
      });

      plugin.addFunction("getUser", "First").addFunction("getUser", "Second");

      assert.equal(plugin.functions.length, 1);
      assert.equal(plugin.getFunction("getUser")?.description, "First");
    });

    it("should remove functions", () => {
      const plugin = PluginManifestWrapper.create({
        schemaVersion: "v2.4",
        nameForHuman: "Test",
        descriptionForHuman: "Test",
      });

      plugin.addFunction("func1").addFunction("func2").removeFunction("func1");

      assert.isFalse(plugin.hasFunction("func1"));
      assert.isTrue(plugin.hasFunction("func2"));
    });
  });

  describe("runtime operations", () => {
    it("should add OpenAPI runtime", () => {
      const plugin = PluginManifestWrapper.create({
        schemaVersion: "v2.4",
        nameForHuman: "Test",
        descriptionForHuman: "Test",
      });

      plugin.addOpenApiRuntime("openapi.yaml", "None", ["*"]);

      assert.equal(plugin.runtimes.length, 1);
      assert.equal(plugin.runtimes[0].type, RuntimeType.OpenApi);
    });

    it("should get API spec paths", () => {
      const plugin = PluginManifestWrapper.create({
        schemaVersion: "v2.4",
        nameForHuman: "Test",
        descriptionForHuman: "Test",
      });

      plugin.addOpenApiRuntime("spec1.yaml").addOpenApiRuntime("spec2.yaml");

      const paths = plugin.getApiSpecPaths();
      assert.deepEqual(paths, ["spec1.yaml", "spec2.yaml"]);
    });

    it("should remove runtime by spec URL", () => {
      const plugin = PluginManifestWrapper.create({
        schemaVersion: "v2.4",
        nameForHuman: "Test",
        descriptionForHuman: "Test",
      });

      plugin.addOpenApiRuntime("keep.yaml").addOpenApiRuntime("remove.yaml");
      plugin.removeRuntimeBySpecUrl("remove.yaml");

      assert.deepEqual(plugin.getApiSpecPaths(), ["keep.yaml"]);
    });
  });

  describe("clone", () => {
    it("should create independent copy", () => {
      const original = PluginManifestWrapper.create({
        schemaVersion: "v2.4",
        nameForHuman: "Original",
        descriptionForHuman: "Test",
      });

      const cloned = original.clone();
      cloned.setNameForHuman("Cloned");

      assert.equal(original.nameForHuman, "Original");
      assert.equal(cloned.nameForHuman, "Cloned");
    });
  });

  describe("save", () => {
    it("should throw if no file path", async () => {
      const plugin = PluginManifestWrapper.create({
        schemaVersion: "v2.4",
        nameForHuman: "Test",
        descriptionForHuman: "Test",
      });

      try {
        await plugin.save();
        assert.fail("Should have thrown");
      } catch (e: unknown) {
        assert.include((e as Error).message, "No file path");
      }
    });

    it("should save to specified path", async () => {
      const writeStub = sandbox.stub(fs, "writeFile").resolves();
      const plugin = PluginManifestWrapper.create({
        schemaVersion: "v2.4",
        nameForHuman: "Test",
        descriptionForHuman: "Test",
      });

      plugin.setNameForHuman("Modified");
      await plugin.save("/path/to/plugin.json");

      assert.isTrue(writeStub.calledOnce);
      assert.equal(plugin.filePath, "/path/to/plugin.json");
      assert.isFalse(plugin.isDirty);
    });
  });
});
