// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, IProgressHandler, ok } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import child_process from "child_process";
import fs from "fs-extra";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import os from "os";
import * as sinon from "sinon";
import * as tools from "../../../../src/common/utils";
import {
  convertScriptErrorToFxError,
  defaultShell,
  executeCommand,
  getStderrHandler,
  parseSetOutputCommand,
  scriptDriver,
} from "../../../../src/component/driver/script/scriptDriver";
import * as charsetUtils from "../../../../src/component/utils/charsetUtils";
import { DefaultEncoding, getSystemEncoding } from "../../../../src/component/utils/charsetUtils";
import { ScriptExecutionError, ScriptTimeoutError } from "../../../../src/error/script";
import {
  MockLogProvider,
  MockUserInteraction,
  MockedAzureAccountProvider,
} from "../../../core/utils";
import { TestLogProvider } from "../../util/logProviderMock";
import { UserCancelError } from "../../../../src/error";

describe("Script Driver test", () => {
  const sandbox = sinon.createSandbox();
  const ui = new MockUserInteraction();
  beforeEach(() => {
    sandbox.stub(tools, "waitSeconds").resolves();
  });
  afterEach(async () => {
    sandbox.restore();
  });
  it("ui not provided - execute success: set-output and append to file", async () => {
    const appendFileSyncStub = sandbox.stub(fs, "appendFileSync");
    const args = {
      workingDirectory: "./",
      run: `echo '::set-output MY_KEY=MY_VALUE'`,
      redirectTo: "./log",
    };
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
      ui: ui,
      progressBar: {
        start: async (detail?: string): Promise<void> => {},
        next: async (detail?: string): Promise<void> => {},
        end: async (): Promise<void> => {},
      } as IProgressHandler,
      projectPath: "./",
    } as any;
    sandbox.stub(ui, "runCommand").value(undefined);
    const res = await scriptDriver.execute(args, context);
    assert.isTrue(res.result.isOk());
    if (res.result.isOk()) {
      const output = res.result.value;
      assert.equal(output.get("MY_KEY"), "MY_VALUE");
    }
    sinon.assert.called(appendFileSyncStub);
  });
  it("ui not provided - execute success: set-output and not append to file", async () => {
    const appendFileSyncStub = sandbox.stub(fs, "appendFileSync");
    const args = {
      workingDirectory: "./",
      run: `echo '::set-output MY_KEY=MY_VALUE'`,
    };
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
      ui: ui,
      progressBar: {
        start: async (detail?: string): Promise<void> => {},
        next: async (detail?: string): Promise<void> => {},
        end: async (): Promise<void> => {},
      } as IProgressHandler,
      projectPath: "./",
    } as any;
    sandbox.stub(ui, "runCommand").value(undefined);
    const res = await scriptDriver.execute(args, context);
    assert.isTrue(res.result.isOk());
    if (res.result.isOk()) {
      const output = res.result.value;
      assert.equal(output.get("MY_KEY"), "MY_VALUE");
    }
    sinon.assert.notCalled(appendFileSyncStub);
  });
  it("ui not provided - execute failed: child_process.exec return error", async () => {
    const error = new Error("test error");
    sandbox.stub(charsetUtils, "getSystemEncoding").resolves("utf-8");
    sandbox.stub(child_process, "exec").yields(error);
    const args = {
      workingDirectory: "./",
      run: "echo '::set-output MY_KEY=MY_VALUE'",
    };
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
      ui: ui,
      projectPath: "./",
    } as any;
    sandbox.stub(ui, "runCommand").value(undefined);
    const res = await scriptDriver.execute(args, context);
    assert.isTrue(res.result.isErr());
  });
  it("ui provided - execute - success", async () => {
    const args = {
      workingDirectory: "./",
      run: `echo '::set-output MY_KEY=MY_VALUE'`,
    };
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
      ui: ui,
      progressBar: {
        start: async (detail?: string): Promise<void> => {},
        next: async (detail?: string): Promise<void> => {},
        end: async (): Promise<void> => {},
      } as IProgressHandler,
      projectPath: "./",
    } as any;
    sandbox.stub(ui, "runCommand").resolves(ok(""));
    const res = await scriptDriver.execute(args, context);
    assert.isTrue(res.result.isOk());
    if (res.result.isOk()) {
      const output = res.result.value;
      assert.equal(output.get("MY_KEY"), "MY_VALUE");
    }
  });
  it("ui provided - execute - success no env output", async () => {
    const args = {
      workingDirectory: "./",
      run: `echo 'abc'`,
    };
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
      ui: ui,
      progressBar: {
        start: async (detail?: string): Promise<void> => {},
        next: async (detail?: string): Promise<void> => {},
        end: async (): Promise<void> => {},
      } as IProgressHandler,
      projectPath: "./",
    } as any;
    sandbox.stub(ui, "runCommand").resolves(ok(""));
    const res = await scriptDriver.execute(args, context);
    assert.isTrue(res.result.isOk());
    if (res.result.isOk()) {
      const output = res.result.value;
      assert.deepEqual(output, new Map());
    }
  });
  it("ui provided - execute - runCommand Error", async () => {
    const args = {
      workingDirectory: "./",
      run: `echo '::set-output MY_KEY=MY_VALUE'`,
    };
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
      ui: ui,
      progressBar: {
        start: async (detail?: string): Promise<void> => {},
        next: async (detail?: string): Promise<void> => {},
        end: async (): Promise<void> => {},
      } as IProgressHandler,
      projectPath: "./",
    } as any;
    sandbox.stub(ui, "runCommand").resolves(err(new UserCancelError()));
    const res = await scriptDriver.execute(args, context);
    assert.isTrue(res.result.isErr());
  });
  it("convertScriptErrorToFxError ScriptTimeoutError", async () => {
    const error = { killed: true } as child_process.ExecException;
    const res = convertScriptErrorToFxError(error, "test");
    assert.isTrue(res instanceof ScriptTimeoutError);
  });
  it("convertScriptErrorToFxError ScriptExecutionError", async () => {
    const error = { killed: false, message: "command failed" } as child_process.ExecException;
    const res = convertScriptErrorToFxError(error, "test");
    assert.isTrue(res instanceof ScriptExecutionError);
  });
});
describe("executeCommand", () => {
  const ui = new MockUserInteraction();
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("dotnet command", async () => {
    sandbox.stub(charsetUtils, "getSystemEncoding").resolves("utf-8");
    const stub = sandbox.stub(child_process, "exec").returns({} as any);
    stub.yields(null);
    sandbox.stub(ui, "runCommand").value(undefined);
    await executeCommand(
      "dotnet test && echo '::set-output MY_KEY=MY_VALUE'",
      "./",
      new TestLogProvider(),
      ui
    );
    assert.isTrue(stub.calledOnce);
  });
  // it("call ui.runCommand", async () => {
  //   const ui = new MockUserInteraction();
  //   const spyRunCommand = sandbox.spy(ui, "runCommand");
  //   const stub = sandbox.stub(child_process, "exec").returns({} as any);
  //   await executeCommand("abc", "./", new TestLogProvider(), ui);
  //   assert.isTrue(spyRunCommand.calledOnce);
  //   assert.isFalse(stub.calledOnce);
  // });
  it("call ui.runCommand error", async () => {
    sandbox.stub(ui, "runCommand").resolves(err(new UserCancelError()));
    sandbox.stub(child_process, "exec").returns({} as any);
    const res = await executeCommand("abc", "./", new TestLogProvider(), ui);
    assert.isTrue(res.isErr());
  });
  it("call ui.runCommand with output", async () => {
    sandbox.stub(ui, "runCommand").resolves(ok(""));
    sandbox.stub(child_process, "exec").returns({} as any);
    const res = await executeCommand(
      "echo '::set-teamsfx-env MY_KEY=MY_VALUE'",
      "./",
      new TestLogProvider(),
      ui
    );
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.deepEqual(res.value[1], { MY_KEY: "MY_VALUE" });
    }
  });
});
describe("getSystemEncoding", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("should return a string", async () => {
    const result = await getSystemEncoding();
    assert.isTrue(typeof result === "string");
  });
  it("should return default encoding on other platform", async () => {
    sandbox.stub(os, "platform").returns("netbsd");
    const result = await getSystemEncoding();
    assert.equal(result, "utf-8");
  });

  it("should return gb2312 on win32 platform", async () => {
    sandbox.stub(os, "platform").returns("win32");
    sandbox.stub(child_process, "exec").callsArgWith(2, null, "Active code page: 936");
    const result = await getSystemEncoding();
    assert.equal(result, "gb2312");
  });

  it("should return utf-8 on linux platform", async () => {
    sandbox.stub(os, "platform").returns("linux");
    sandbox.stub(child_process, "exec").callsArgWith(2, null, "UTF-8");
    const result = await getSystemEncoding();
    assert.equal(result, "utf-8");
  });

  it("should return utf-8 on darwin platform", async () => {
    sandbox.stub(os, "platform").returns("darwin");
    sandbox.stub(child_process, "exec").callsArgWith(2, null, "zh_CN.UTF-8");
    const result = await getSystemEncoding();
    assert.equal(result, "utf-8");
  });

  it("should return default encoding when Error happens on win32 platform", async () => {
    sandbox.stub(os, "platform").returns("win32");
    const error = new Error("test error");
    sandbox.stub(child_process, "exec").callsArgWith(2, error, "");
    const result = await getSystemEncoding();
    assert.equal(result, DefaultEncoding);
  });

  it("should return default encoding when Error happens on linux platform", async () => {
    sandbox.stub(os, "platform").returns("linux");
    const error = new Error("test error");
    sandbox.stub(child_process, "exec").callsArgWith(2, error, "");
    const result = await getSystemEncoding();
    assert.equal(result, DefaultEncoding);
  });

  it("should return default encoding when Error happens on darwin platform", async () => {
    sandbox.stub(os, "platform").returns("darwin");
    const error = new Error("test error");
    sandbox.stub(child_process, "exec").callsArgWith(2, error, "");
    const result = await getSystemEncoding();
    assert.equal(result, DefaultEncoding);
  });
  it("should return utf8 for azure cli", async () => {
    const result = await getSystemEncoding("@azure/static-web-apps-cli");
    assert.equal(result, "utf8");
  });
});

describe("parseSetOutputCommand", () => {
  it("parse one key value pair", async () => {
    const res = parseSetOutputCommand('echo "::set-teamsfx-env TAB_DOMAIN=localhost:53000"');
    assert.deepEqual(res, { TAB_DOMAIN: "localhost:53000" });
  });
  it("parse two key value pairs", async () => {
    const res = parseSetOutputCommand(
      'echo "::set-teamsfx-env TAB_DOMAIN=localhost:53000"; echo "::set-teamsfx-env TAB_ENDPOINT=https://localhost:53000";'
    );
    assert.deepEqual(res, {
      TAB_DOMAIN: "localhost:53000",
      TAB_ENDPOINT: "https://localhost:53000",
    });
  });
  it("parse value that contains space", async () => {
    const res = parseSetOutputCommand(
      `Write-Host ::set-teamsfx-env Test0="multi word variable"
        Write-Host ::set-teamsfx-env Test1=' multi word variable'
        Write-Host ::set-teamsfx-env Test2=multi+word+variable`
    );
    assert.deepEqual(res, {
      Test0: "multi word variable",
      Test1: " multi word variable",
      Test2: "multi+word+variable",
    });
  });
});

describe("getStderrHandler", () => {
  const sandbox = sinon.createSandbox();
  beforeEach(() => {});
  afterEach(async () => {
    sandbox.restore();
  });
  it("happy path", async () => {
    const logProvider = new MockLogProvider();
    const systemEncoding = "utf-8";
    const stderrStrings: string[] = [];
    const handler = getStderrHandler(
      logProvider,
      systemEncoding,
      stderrStrings,
      async (data: string) => {}
    );
    await handler(Buffer.from("test"));
    assert.deepEqual(stderrStrings, ["test"]);
  });
});

describe("defaultShell", () => {
  const sandbox = sinon.createSandbox();
  let restoreEnv: RestoreFn = () => {};
  afterEach(() => {
    sandbox.restore();
    restoreEnv();
  });
  it("SHELL", async () => {
    restoreEnv = mockedEnv({ SHELL: "/bin/bash" });
    const result = await defaultShell();
    assert.equal(result, "/bin/bash");
  });
  it("darwin - /bin/zsh", async () => {
    // Ensure SHELL does not override platform-specific logic
    restoreEnv = mockedEnv({ SHELL: undefined });
    sandbox.stub(process, "platform").value("darwin");
    sandbox.stub(fs, "pathExists").resolves(true);
    const result = await defaultShell();
    assert.equal(result, "/bin/zsh");
  });
  it("darwin - /bin/bash", async () => {
    // Ensure SHELL does not override platform-specific logic
    restoreEnv = mockedEnv({ SHELL: undefined });
    sandbox.stub(process, "platform").value("darwin");
    sandbox.stub(fs, "pathExists").onFirstCall().resolves(false).onSecondCall().resolves(true);
    const result = await defaultShell();
    assert.equal(result, "/bin/bash");
  });
  it("darwin - undefined", async () => {
    // Ensure SHELL does not override platform-specific logic
    restoreEnv = mockedEnv({ SHELL: undefined });
    sandbox.stub(process, "platform").value("darwin");
    sandbox.stub(fs, "pathExists").resolves(false);
    const result = await defaultShell();
    assert.isUndefined(result);
  });

  it("win32 - ComSpec", async () => {
    sandbox.stub(process, "platform").value("win32");
    // Ensure SHELL does not override platform-specific logic
    restoreEnv = mockedEnv({ SHELL: undefined, ComSpec: "cmd.exe" });
    const result = await defaultShell();
    assert.equal(result, "cmd.exe");
  });
  it("win32 - cmd.exe", async () => {
    sandbox.stub(process, "platform").value("win32");
    // Ensure SHELL does not override platform-specific logic
    restoreEnv = mockedEnv({ SHELL: undefined, ComSpec: undefined });
    const result = await defaultShell();
    assert.equal(result, "cmd.exe");
  });

  it("other OS - /bin/sh", async () => {
    // Ensure SHELL does not override platform-specific logic
    restoreEnv = mockedEnv({ SHELL: undefined });
    sandbox.stub(process, "platform").value("other");
    sandbox.stub(fs, "pathExists").resolves(true);
    const result = await defaultShell();
    assert.equal(result, "/bin/sh");
  });

  it("other OS - undefined", async () => {
    // Ensure SHELL does not override platform-specific logic
    restoreEnv = mockedEnv({ SHELL: undefined });
    sandbox.stub(process, "platform").value("other");
    sandbox.stub(fs, "pathExists").resolves(false);
    const result = await defaultShell();
    assert.isUndefined(result);
  });
});
