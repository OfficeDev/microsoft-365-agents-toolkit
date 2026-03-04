// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { EventEmitter } from "events";
import sinon from "sinon";
import cp from "child_process";
import { FxError, LogLevel, SystemError } from "@microsoft/teamsfx-api";
import { expect } from "../../utils";
import { Task, TaskResult } from "../../../../src/cmds/preview/task";
import { ServiceLogWriter } from "../../../../src/cmds/preview/serviceLogWriter";
import { CLILogProvider } from "../../../../src/commonlib/log";

function createMockChildProcess(exitCode: number | null = 0): cp.ChildProcess {
  const proc = new EventEmitter() as cp.ChildProcess;
  const mockStdout = new EventEmitter();
  const mockStderr = new EventEmitter();
  (proc as any).stdout = mockStdout;
  (proc as any).stderr = mockStderr;
  (proc as any).exitCode = exitCode;
  (proc as any).pid = 12345;
  return proc;
}

function createFxError(message: string): FxError {
  return new SystemError({
    source: "test",
    name: "TestError",
    message,
  });
}

// Save reference before sinon.useFakeTimers can replace it
const realSetImmediate = setImmediate;

/**
 * Flush microtask queue to allow async setup (await calls) inside
 * wait()/waitFor() to complete before we emit mock events.
 */
function flush(): Promise<void> {
  return new Promise((resolve) => realSetImmediate(resolve));
}

describe("Task", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("wait", () => {
    it("should resolve ok when task exits with code 0", async () => {
      const mockProc = createMockChildProcess(0);
      sandbox.stub(cp, "spawn").returns(mockProc);

      const task = new Task("test-task", false, "echo", ["hello"], { shell: true });
      const startCallback = sandbox.stub().resolves();
      const stopCallback = sandbox.stub().resolves(null);

      const promise = task.wait(startCallback, stopCallback);
      await flush();

      mockProc.stdout!.emit("data", Buffer.from("hello\n"));
      mockProc.emit("exit");

      const result = await promise;
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.success).to.be.true;
        expect(result.value.command).to.equal("echo");
        expect(result.value.stdout).to.deep.equal(["hello\n"]);
        expect(result.value.stderr).to.deep.equal([]);
        expect(result.value.exitCode).to.equal(0);
      }
      expect(startCallback.calledOnce).to.be.true;
      expect(startCallback.calledWith("test-task", false)).to.be.true;
      expect(stopCallback.calledOnce).to.be.true;
    });

    it("should resolve ok with success=false when task exits with non-zero code", async () => {
      const mockProc = createMockChildProcess(1);
      sandbox.stub(cp, "spawn").returns(mockProc);

      const task = new Task("test-task", false, "fail-cmd", undefined, { shell: true });
      const startCallback = sandbox.stub().resolves();
      const stopCallback = sandbox.stub().resolves(null);

      const promise = task.wait(startCallback, stopCallback);
      await flush();

      mockProc.stderr!.emit("data", Buffer.from("error output\n"));
      mockProc.emit("exit");

      const result = await promise;
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.success).to.be.false;
        expect(result.value.exitCode).to.equal(1);
        expect(result.value.stderr).to.deep.equal(["error output\n"]);
      }
    });

    it("should resolve err when stopCallback returns an error", async () => {
      const mockProc = createMockChildProcess(0);
      sandbox.stub(cp, "spawn").returns(mockProc);

      const task = new Task("test-task", false, "echo", [], { shell: true });
      const startCallback = sandbox.stub().resolves();
      const fxError = createFxError("stop error");
      const stopCallback = sandbox.stub().resolves(fxError);

      const promise = task.wait(startCallback, stopCallback);
      await flush();
      mockProc.emit("exit");

      const result = await promise;
      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.message).to.equal("stop error");
      }
    });

    it("should collect multiple stdout and stderr chunks", async () => {
      const mockProc = createMockChildProcess(0);
      sandbox.stub(cp, "spawn").returns(mockProc);

      const task = new Task("test-task", false, "cmd", undefined, { shell: true });
      const startCallback = sandbox.stub().resolves();
      const stopCallback = sandbox.stub().resolves(null);

      const promise = task.wait(startCallback, stopCallback);
      await flush();

      mockProc.stdout!.emit("data", Buffer.from("line1\n"));
      mockProc.stdout!.emit("data", Buffer.from("line2\n"));
      mockProc.stderr!.emit("data", Buffer.from("warn1\n"));
      mockProc.emit("exit");

      const result = await promise;
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.stdout).to.deep.equal(["line1\n", "line2\n"]);
        expect(result.value.stderr).to.deep.equal(["warn1\n"]);
      }
    });

    it("should set exitCode to null when task exitCode is undefined", async () => {
      const mockProc = createMockChildProcess(undefined as any);
      (mockProc as any).exitCode = undefined;
      sandbox.stub(cp, "spawn").returns(mockProc);

      const task = new Task("test-task", false, "cmd", undefined, { shell: true });
      const startCallback = sandbox.stub().resolves();
      const stopCallback = sandbox.stub().resolves(null);

      const promise = task.wait(startCallback, stopCallback);
      await flush();
      mockProc.emit("exit");

      const result = await promise;
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.exitCode).to.be.null;
      }
    });
  });

  describe("waitFor", () => {
    it("should resolve ok when stdout matches pattern", async () => {
      const mockProc = createMockChildProcess(null);
      sandbox.stub(cp, "spawn").returns(mockProc);

      const task = new Task("test-task", true, "server", ["start"], { shell: true });
      const startCallback = sandbox.stub().resolves();
      const stopCallback = sandbox.stub().resolves(null);

      const promise = task.waitFor(/ready/, startCallback, stopCallback);
      await flush();

      mockProc.stdout!.emit("data", Buffer.from("starting...\n"));
      mockProc.stdout!.emit("data", Buffer.from("server ready on port 3000\n"));

      const result = await promise;
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.success).to.be.true;
        expect(result.value.exitCode).to.be.null;
        expect(result.value.stdout).to.deep.equal([
          "starting...\n",
          "server ready on port 3000\n",
        ]);
      }
      expect(startCallback.calledOnce).to.be.true;
    });

    it("should resolve err when stopCallback returns error on pattern match", async () => {
      const mockProc = createMockChildProcess(null);
      sandbox.stub(cp, "spawn").returns(mockProc);

      const task = new Task("test-task", true, "server", [], { shell: true });
      const startCallback = sandbox.stub().resolves();
      const fxError = createFxError("pattern match error");
      const stopCallback = sandbox.stub().resolves(fxError);

      const promise = task.waitFor(/ready/, startCallback, stopCallback);
      await flush();
      mockProc.stdout!.emit("data", Buffer.from("ready\n"));

      const result = await promise;
      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.message).to.equal("pattern match error");
      }
    });

    it("should resolve with success=false when task exits without matching pattern", async () => {
      const mockProc = createMockChildProcess(1);
      sandbox.stub(cp, "spawn").returns(mockProc);

      const task = new Task("test-task", true, "server", [], { shell: true });
      const startCallback = sandbox.stub().resolves();
      const stopCallback = sandbox.stub().resolves(null);

      const promise = task.waitFor(/ready/, startCallback, stopCallback);
      await flush();

      mockProc.stdout!.emit("data", Buffer.from("not matching\n"));
      mockProc.emit("exit");

      const result = await promise;
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.success).to.be.false;
        expect(result.value.exitCode).to.equal(1);
      }
    });

    it("should resolve via timeout when pattern is not matched", async () => {
      const clock = sandbox.useFakeTimers();
      const mockProc = createMockChildProcess(null);
      sandbox.stub(cp, "spawn").returns(mockProc);

      const task = new Task("test-task", true, "server", [], { shell: true });
      const startCallback = sandbox.stub().resolves();
      const stopCallback = sandbox.stub().resolves(null);

      const promise = task.waitFor(/ready/, startCallback, stopCallback, 1000);
      await flush();

      mockProc.stdout!.emit("data", Buffer.from("not matching\n"));
      clock.tick(1001);

      const result = await promise;
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.success).to.be.true;
        expect(result.value.exitCode).to.be.null;
      }
    });

    it("should not resolve again after pattern already matched", async () => {
      const mockProc = createMockChildProcess(0);
      sandbox.stub(cp, "spawn").returns(mockProc);

      const task = new Task("test-task", true, "server", [], { shell: true });
      const startCallback = sandbox.stub().resolves();
      const stopCallback = sandbox.stub().resolves(null);

      const promise = task.waitFor(/ready/, startCallback, stopCallback);
      await flush();

      // First match resolves
      mockProc.stdout!.emit("data", Buffer.from("ready\n"));
      // Exit should not re-resolve
      mockProc.emit("exit");

      const result = await promise;
      expect(result.isOk()).to.be.true;
      // stopCallback should only be called once (for the pattern match)
      expect(stopCallback.calledOnce).to.be.true;
    });

    it("should write to serviceLogWriter when provided", async () => {
      const mockProc = createMockChildProcess(null);
      sandbox.stub(cp, "spawn").returns(mockProc);

      const serviceLogWriter = {
        write: sandbox.stub().resolves(),
      } as unknown as ServiceLogWriter;

      const task = new Task("test-task", true, "server", ["--port", "3000"], { shell: true });
      const startCallback = sandbox.stub().resolves();
      const stopCallback = sandbox.stub().resolves(null);

      const promise = task.waitFor(
        /ready/,
        startCallback,
        stopCallback,
        undefined,
        serviceLogWriter
      );
      await flush();

      mockProc.stdout!.emit("data", Buffer.from("ready\n"));

      const result = await promise;
      expect(result.isOk()).to.be.true;

      const writeStub = serviceLogWriter.write as sinon.SinonStub;
      // First call writes the command
      expect(writeStub.calledWith("test-task", "server --port 3000\n")).to.be.true;
      // Subsequent call writes stdout data
      expect(writeStub.calledWith("test-task", "ready\n")).to.be.true;
    });

    it("should log to logProvider when provided", async () => {
      const mockProc = createMockChildProcess(null);
      sandbox.stub(cp, "spawn").returns(mockProc);

      const logProvider = {
        necessaryLog: sandbox.stub(),
      } as unknown as CLILogProvider;

      const task = new Task("test-task", true, "server", [], { shell: true });
      const startCallback = sandbox.stub().resolves();
      const stopCallback = sandbox.stub().resolves(null);

      const promise = task.waitFor(
        /ready/,
        startCallback,
        stopCallback,
        undefined,
        undefined,
        logProvider
      );
      await flush();

      mockProc.stdout!.emit("data", Buffer.from("ready\n"));

      await promise;

      const logStub = logProvider.necessaryLog as sinon.SinonStub;
      expect(logStub.calledOnce).to.be.true;
      expect(logStub.firstCall.args[0]).to.equal(LogLevel.Info);
      expect(logStub.firstCall.args[1]).to.equal("ready");
    });

    it("should collect stderr and write to serviceLogWriter", async () => {
      const mockProc = createMockChildProcess(1);
      sandbox.stub(cp, "spawn").returns(mockProc);

      const serviceLogWriter = {
        write: sandbox.stub().resolves(),
      } as unknown as ServiceLogWriter;

      const task = new Task("test-task", true, "server", [], { shell: true });
      const startCallback = sandbox.stub().resolves();
      const stopCallback = sandbox.stub().resolves(null);

      const promise = task.waitFor(
        /ready/,
        startCallback,
        stopCallback,
        undefined,
        serviceLogWriter
      );
      await flush();

      mockProc.stderr!.emit("data", Buffer.from("warning msg\n"));
      mockProc.emit("exit");

      const result = await promise;
      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.stderr).to.deep.equal(["warning msg\n"]);
      }

      const writeStub = serviceLogWriter.write as sinon.SinonStub;
      expect(writeStub.calledWith("test-task", "warning msg\n")).to.be.true;
    });

    it("should resolve err on timeout when stopCallback returns error", async () => {
      const clock = sandbox.useFakeTimers();
      const mockProc = createMockChildProcess(null);
      sandbox.stub(cp, "spawn").returns(mockProc);

      const task = new Task("test-task", true, "server", [], { shell: true });
      const startCallback = sandbox.stub().resolves();
      const fxError = createFxError("timeout error");
      const stopCallback = sandbox.stub().resolves(fxError);

      const promise = task.waitFor(/ready/, startCallback, stopCallback, 500);
      await flush();

      clock.tick(501);

      const result = await promise;
      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.message).to.equal("timeout error");
      }
    });
  });

  describe("terminate", () => {
    it("should resolve immediately when task has exitCode", async () => {
      const mockProc = createMockChildProcess(0);
      sandbox.stub(cp, "spawn").returns(mockProc);

      const task = new Task("test-task", false, "echo", [], { shell: true });
      const startCallback = sandbox.stub().resolves();
      const stopCallback = sandbox.stub().resolves(null);

      const waitPromise = task.wait(startCallback, stopCallback);
      await flush();
      mockProc.emit("exit");
      await waitPromise;

      await task.terminate();
      // Should resolve without error
    });

    it("should resolve immediately when pid is undefined", async () => {
      const task = new Task("test-task", false, "echo", [], { shell: true });
      // task.task is undefined, so pid is undefined
      await task.terminate();
      // Should resolve without error
    });

    it("should call treeKill with pid and resolve on success", async () => {
      const mockProc = createMockChildProcess(null);
      (mockProc as any).exitCode = null;
      (mockProc as any).pid = 99999;
      sandbox.stub(cp, "spawn").returns(mockProc);

      // Dynamically require task module after replacing tree-kill in cache
      const treeKillPath = require.resolve("tree-kill");
      const taskPath = require.resolve("../../../../src/cmds/preview/task");
      const originalTreeKill = require.cache[treeKillPath]!.exports;
      const originalTask = require.cache[taskPath];

      // Remove cached task module so it re-requires tree-kill
      delete require.cache[taskPath];

      const treeKillStub = sinon.stub().callsFake((...args: any[]) => {
        const callback = args.length === 2 ? args[1] : args[2];
        if (typeof callback === "function") callback();
      });
      require.cache[treeKillPath]!.exports = treeKillStub;

      // Re-require task module with stubbed tree-kill
      const { Task: TaskWithStub } = require("../../../../src/cmds/preview/task");

      const task = new TaskWithStub("test-task", false, "long-cmd", [], { shell: true });
      const startCallback = sandbox.stub().resolves();
      const stopCallback = sandbox.stub().resolves(null);

      task.wait(startCallback, stopCallback);
      await flush();

      await task.terminate();
      expect(treeKillStub.calledOnce).to.be.true;
      expect(treeKillStub.firstCall.args[0]).to.equal(99999);

      // Restore caches
      require.cache[treeKillPath]!.exports = originalTreeKill;
      if (originalTask) {
        require.cache[taskPath] = originalTask;
      } else {
        delete require.cache[taskPath];
      }
    });

    it("should resolve even when treeKill returns error", async () => {
      const mockProc = createMockChildProcess(null);
      (mockProc as any).exitCode = null;
      (mockProc as any).pid = 99999;
      sandbox.stub(cp, "spawn").returns(mockProc);

      // Dynamically require task module after replacing tree-kill in cache
      const treeKillPath = require.resolve("tree-kill");
      const taskPath = require.resolve("../../../../src/cmds/preview/task");
      const originalTreeKill = require.cache[treeKillPath]!.exports;
      const originalTask = require.cache[taskPath];

      delete require.cache[taskPath];

      const treeKillStub = sinon.stub().callsFake((...args: any[]) => {
        const callback = args.length === 2 ? args[1] : args[2];
        if (typeof callback === "function") callback(new Error("kill failed"));
      });
      require.cache[treeKillPath]!.exports = treeKillStub;

      const { Task: TaskWithStub } = require("../../../../src/cmds/preview/task");

      const task = new TaskWithStub("test-task", false, "long-cmd", [], { shell: true });
      const startCallback = sandbox.stub().resolves();
      const stopCallback = sandbox.stub().resolves(null);

      task.wait(startCallback, stopCallback);
      await flush();

      // Should still resolve without throwing
      await task.terminate();

      // Restore caches
      require.cache[treeKillPath]!.exports = originalTreeKill;
      if (originalTask) {
        require.cache[taskPath] = originalTask;
      } else {
        delete require.cache[taskPath];
      }
    });
  });
});
