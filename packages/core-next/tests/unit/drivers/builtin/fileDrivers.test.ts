/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import { createOrUpdateEnvironmentFileDriver } from "../../../../src/drivers/builtin/file/createOrUpdateEnvironmentFile";
import { createOrUpdateJsonFileDriver } from "../../../../src/drivers/builtin/file/createOrUpdateJsonFile";
import { createMockContext } from "../../testHelper";

describe("file/createOrUpdateEnvironmentFile driver", () => {
  let tmpDir: string;

  afterEach(async () => {
    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  async function setup(): Promise<string> {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-test-envfile-"));
    return tmpDir;
  }

  it("should create a new .env file", async () => {
    const dir = await setup();
    const ctx = createMockContext({ projectPath: dir });

    const result = await createOrUpdateEnvironmentFileDriver.executeFn(ctx, {
      target: ".env",
      envs: { API_KEY: "abc123", ENDPOINT: "https://example.com" },
    });

    expect(result.isOk()).to.be.true;
    const content = await fs.readFile(path.join(dir, ".env"), "utf-8");
    expect(content).to.include("API_KEY=abc123");
    expect(content).to.include("ENDPOINT=https://example.com");

    if (result.isOk()) {
      expect(result.value.outputs).to.have.property("API_KEY", "abc123");
      expect(result.value.outputs).to.have.property("ENDPOINT", "https://example.com");
    }
  });

  it("should merge with existing .env file", async () => {
    const dir = await setup();
    const envPath = path.join(dir, ".env");
    await fs.writeFile(envPath, "EXISTING=value1\nKEEP=me\n", "utf-8");

    const ctx = createMockContext({ projectPath: dir });
    const result = await createOrUpdateEnvironmentFileDriver.executeFn(ctx, {
      target: ".env",
      envs: { EXISTING: "updated", NEW_VAR: "hello" },
    });

    expect(result.isOk()).to.be.true;
    const content = await fs.readFile(envPath, "utf-8");
    expect(content).to.include("EXISTING=updated");
    expect(content).to.include("KEEP=me");
    expect(content).to.include("NEW_VAR=hello");
  });

  it("should create nested directories for target", async () => {
    const dir = await setup();
    const ctx = createMockContext({ projectPath: dir });

    const result = await createOrUpdateEnvironmentFileDriver.executeFn(ctx, {
      target: "sub/dir/.env.local",
      envs: { VAR: "val" },
    });

    expect(result.isOk()).to.be.true;
    const content = await fs.readFile(path.join(dir, "sub", "dir", ".env.local"), "utf-8");
    expect(content).to.include("VAR=val");
  });

  it("should fail when projectPath is undefined", async () => {
    const ctx = createMockContext({ projectPath: undefined });

    const result = await createOrUpdateEnvironmentFileDriver.executeFn(ctx, {
      target: ".env",
      envs: { A: "1" },
    });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("ProjectPathRequired");
    }
  });

  it("should reject invalid config (missing envs)", async () => {
    const ctx = createMockContext();
    const result = await createOrUpdateEnvironmentFileDriver.executeFn(ctx, {
      target: ".env",
    });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("InvalidDriverInput");
    }
  });

  it("should reject invalid config (empty target)", async () => {
    const ctx = createMockContext();
    const result = await createOrUpdateEnvironmentFileDriver.executeFn(ctx, {
      target: "",
      envs: { A: "1" },
    });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("InvalidDriverInput");
    }
  });
});

describe("file/createOrUpdateJsonFile driver", () => {
  let tmpDir: string;

  afterEach(async () => {
    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  async function setup(): Promise<string> {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-test-jsonfile-"));
    return tmpDir;
  }

  it("should create a new JSON file", async () => {
    const dir = await setup();
    const ctx = createMockContext({ projectPath: dir });

    const result = await createOrUpdateJsonFileDriver.executeFn(ctx, {
      target: "config.json",
      content: { key: "value", nested: { a: 1 } },
    });

    expect(result.isOk()).to.be.true;
    const raw = await fs.readFile(path.join(dir, "config.json"), "utf-8");
    const parsed = JSON.parse(raw);
    expect(parsed.key).to.equal("value");
    expect(parsed.nested.a).to.equal(1);
  });

  it("should deep-merge with existing JSON file", async () => {
    const dir = await setup();
    const filePath = path.join(dir, "settings.json");
    await fs.writeFile(
      filePath,
      JSON.stringify({ existing: true, nested: { keep: 1, replace: "old" } }),
      "utf-8"
    );

    const ctx = createMockContext({ projectPath: dir });
    const result = await createOrUpdateJsonFileDriver.executeFn(ctx, {
      target: "settings.json",
      content: { nested: { replace: "new", added: 2 }, extra: "yes" },
    });

    expect(result.isOk()).to.be.true;
    const raw = await fs.readFile(filePath, "utf-8");
    const parsed = JSON.parse(raw);
    expect(parsed.existing).to.be.true;
    expect(parsed.nested.keep).to.equal(1);
    expect(parsed.nested.replace).to.equal("new");
    expect(parsed.nested.added).to.equal(2);
    expect(parsed.extra).to.equal("yes");
  });

  it("should support 'appsettings' alias", async () => {
    const dir = await setup();
    const ctx = createMockContext({ projectPath: dir });

    const result = await createOrUpdateJsonFileDriver.executeFn(ctx, {
      target: "appsettings.json",
      appsettings: { ConnectionStrings: { Default: "Server=localhost" } },
    });

    expect(result.isOk()).to.be.true;
    const raw = await fs.readFile(path.join(dir, "appsettings.json"), "utf-8");
    const parsed = JSON.parse(raw);
    expect(parsed.ConnectionStrings.Default).to.equal("Server=localhost");
  });

  it("should fail when projectPath is undefined", async () => {
    const ctx = createMockContext({ projectPath: undefined });

    const result = await createOrUpdateJsonFileDriver.executeFn(ctx, {
      target: "f.json",
      content: { a: 1 },
    });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("ProjectPathRequired");
    }
  });

  it("should reject config with neither content nor appsettings", async () => {
    const ctx = createMockContext();
    const result = await createOrUpdateJsonFileDriver.executeFn(ctx, {
      target: "f.json",
    });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("InvalidDriverInput");
    }
  });
});
