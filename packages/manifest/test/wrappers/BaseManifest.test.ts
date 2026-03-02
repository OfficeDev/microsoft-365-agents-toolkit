// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import "mocha";
import { BaseManifest } from "../../src/wrappers/BaseManifest";

// Create a concrete implementation for testing
class TestManifest extends BaseManifest<{ name: string; value: number }> {
  constructor(data: { name: string; value: number }, filePath?: string) {
    super(data, filePath);
  }

  static create(data: { name: string; value: number }, filePath?: string): TestManifest {
    return new TestManifest(data, filePath);
  }

  static async readFromFile(filePath: string): Promise<TestManifest> {
    const data = await BaseManifest.readJsonFile<{ name: string; value: number }>(filePath);
    return new TestManifest(data, filePath);
  }

  static readFromFileSync(filePath: string): TestManifest {
    const data = BaseManifest.readJsonFileSync<{ name: string; value: number }>(filePath);
    return new TestManifest(data, filePath);
  }

  validate(): Promise<string[]> {
    return Promise.resolve([]);
  }

  toJSON(): string {
    return JSON.stringify(this._data, null, 2);
  }

  clone(): TestManifest {
    return new TestManifest(JSON.parse(this.toJSON()), this._filePath);
  }

  // Expose protected methods for testing
  public testMarkDirty(): void {
    this.markDirty();
  }
}

describe("BaseManifest", () => {
  describe("constructor", () => {
    it("should initialize with data", () => {
      const manifest = TestManifest.create({ name: "test", value: 42 });

      assert.deepEqual(manifest.data, { name: "test", value: 42 });
      assert.isUndefined(manifest.filePath);
      assert.isFalse(manifest.isDirty);
    });

    it("should initialize with data and file path", () => {
      const manifest = TestManifest.create({ name: "test", value: 42 }, "/path/to/file.json");

      assert.equal(manifest.filePath, "/path/to/file.json");
    });
  });

  describe("data getter", () => {
    it("should return readonly data", () => {
      const manifest = TestManifest.create({ name: "test", value: 42 });

      const data = manifest.data;
      assert.equal(data.name, "test");
      assert.equal(data.value, 42);
    });
  });

  describe("filePath getter", () => {
    it("should return undefined when no file path", () => {
      const manifest = TestManifest.create({ name: "test", value: 42 });
      assert.isUndefined(manifest.filePath);
    });

    it("should return file path when set", () => {
      const manifest = TestManifest.create({ name: "test", value: 42 }, "/my/path.json");
      assert.equal(manifest.filePath, "/my/path.json");
    });
  });

  describe("isDirty getter", () => {
    it("should return false initially", () => {
      const manifest = TestManifest.create({ name: "test", value: 42 });
      assert.isFalse(manifest.isDirty);
    });

    it("should return true after markDirty is called", () => {
      const manifest = TestManifest.create({ name: "test", value: 42 });
      manifest.testMarkDirty();
      assert.isTrue(manifest.isDirty);
    });
  });

  describe("markDirty", () => {
    it("should set isDirty to true", () => {
      const manifest = TestManifest.create({ name: "test", value: 42 });
      assert.isFalse(manifest.isDirty);

      manifest.testMarkDirty();
      assert.isTrue(manifest.isDirty);
    });
  });

  describe("save", () => {
    it("should throw error when no file path available", async () => {
      const manifest = TestManifest.create({ name: "test", value: 42 });

      try {
        await manifest.save();
        assert.fail("Should have thrown an error");
      } catch (e: unknown) {
        assert.include((e as Error).message, "No file path specified");
      }
    });
  });

  describe("toJSON", () => {
    it("should convert to JSON string", () => {
      const manifest = TestManifest.create({ name: "test", value: 42 });
      const json = manifest.toJSON();

      const parsed = JSON.parse(json);
      assert.equal(parsed.name, "test");
      assert.equal(parsed.value, 42);
    });
  });

  describe("clone", () => {
    it("should create independent copy", () => {
      const original = TestManifest.create({ name: "original", value: 1 });
      const cloned = original.clone();

      // Modify cloned data
      (cloned as unknown as { _data: { name: string } })._data.name = "cloned";

      assert.equal(original.data.name, "original");
      assert.equal(cloned.data.name, "cloned");
    });
  });
});
