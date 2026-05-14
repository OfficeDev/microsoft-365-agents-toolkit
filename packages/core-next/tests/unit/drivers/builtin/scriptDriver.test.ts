/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it, afterEach } from "mocha";
import * as fs from "node:fs/promises";
import * as path from "node:path";
import * as os from "node:os";
import {
  scriptDriver,
  parseOutputDirectives,
  defaultShell,
} from "../../../../src/drivers/builtin/script/run";
import { createMockContext } from "../../testHelper";

describe("script driver", () => {
  let tmpDir: string;

  afterEach(async () => {
    if (tmpDir) {
      await fs.rm(tmpDir, { recursive: true, force: true });
    }
  });

  async function setup(): Promise<string> {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "atk-test-script-"));
    return tmpDir;
  }

  it("should execute a simple echo command", async () => {
    const dir = await setup();
    const ctx = createMockContext({ projectPath: dir });

    const command = os.platform() === "win32" ? "echo hello" : "echo hello";
    const result = await scriptDriver.executeFn(ctx, { run: command });

    expect(result.isOk()).to.be.true;
  });

  it("should capture ::set-output directives as outputs", async () => {
    const dir = await setup();
    const ctx = createMockContext({ projectPath: dir });

    const command =
      os.platform() === "win32"
        ? "echo ::set-output MY_VAR=myvalue"
        : 'echo "::set-output MY_VAR=myvalue"';
    const result = await scriptDriver.executeFn(ctx, { run: command });

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.outputs).to.have.property("MY_VAR", "myvalue");
    }
  });

  it("should capture ::set-teamsfx-env directives", async () => {
    const dir = await setup();
    const ctx = createMockContext({ projectPath: dir });

    const command =
      os.platform() === "win32"
        ? "echo ::set-teamsfx-env API_URL=https://example.com"
        : 'echo "::set-teamsfx-env API_URL=https://example.com"';
    const result = await scriptDriver.executeFn(ctx, { run: command });

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.outputs).to.have.property("API_URL", "https://example.com");
    }
  });

  it("should redirect output to a file", async () => {
    const dir = await setup();
    const ctx = createMockContext({ projectPath: dir });

    const command = os.platform() === "win32" ? "echo redirected" : "echo redirected";
    const result = await scriptDriver.executeFn(ctx, {
      run: command,
      redirectTo: "output.log",
    });

    expect(result.isOk()).to.be.true;
    const logContent = await fs.readFile(path.join(dir, "output.log"), "utf-8");
    expect(logContent).to.include("redirected");
  });

  it("should resolve working directory relative to projectPath", async () => {
    const dir = await setup();
    const subDir = path.join(dir, "subdir");
    await fs.mkdir(subDir, { recursive: true });

    const ctx = createMockContext({ projectPath: dir });
    const command = os.platform() === "win32" ? "echo %CD%" : "pwd";
    const result = await scriptDriver.executeFn(ctx, {
      run: command,
      workingDirectory: "subdir",
    });

    expect(result.isOk()).to.be.true;
  });

  it("should return error for failing command", async () => {
    const dir = await setup();
    const ctx = createMockContext({ projectPath: dir });

    const result = await scriptDriver.executeFn(ctx, {
      run: "exit 1",
      shell: os.platform() === "win32" ? "cmd.exe" : "/bin/sh",
    });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("ScriptExecutionError");
    }
  });

  it("should reject invalid config (empty run)", async () => {
    const ctx = createMockContext();
    const result = await scriptDriver.executeFn(ctx, { run: "" });

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.code).to.equal("InvalidDriverInput");
    }
  });
});

describe("parseOutputDirectives", () => {
  it("should parse ::set-output directives", () => {
    const stdout = [
      "some output",
      "::set-output KEY1=value1",
      "::set-output KEY2=value2",
      "other line",
    ].join("\n");

    const result = parseOutputDirectives(stdout);
    expect(result).to.deep.equal({ KEY1: "value1", KEY2: "value2" });
  });

  it("should parse ::set-teamsfx-env directives", () => {
    const stdout = "::set-teamsfx-env ENDPOINT=https://api.example.com\n";
    const result = parseOutputDirectives(stdout);
    expect(result).to.deep.equal({ ENDPOINT: "https://api.example.com" });
  });

  it("should handle quoted values", () => {
    const stdout = '::set-output NAME="hello world"\n';
    const result = parseOutputDirectives(stdout);
    expect(result).to.deep.equal({ NAME: "hello world" });
  });

  it("should return empty object for no directives", () => {
    const result = parseOutputDirectives("just normal output\nno directives here");
    expect(result).to.deep.equal({});
  });

  it("should handle mixed directive types", () => {
    const stdout = ["::set-output A=1", "::set-teamsfx-env B=2"].join("\n");
    const result = parseOutputDirectives(stdout);
    expect(result).to.deep.equal({ A: "1", B: "2" });
  });
});

describe("defaultShell", () => {
  it("should return a string", () => {
    const shell = defaultShell();
    expect(shell).to.be.a("string");
    expect(shell.length).to.be.greaterThan(0);
  });
});
