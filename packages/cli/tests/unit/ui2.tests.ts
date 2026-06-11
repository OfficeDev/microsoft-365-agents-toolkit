// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  InputTextConfig,
  MultiSelectConfig,
  SingleSelectConfig,
  UserError,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import { SelectSubscriptionError, UnhandledError } from "@microsoft/teamsfx-core";
import { assert } from "chai";
import child_process from "child_process";
import * as sinon from "sinon";
import { logger } from "../../src/commonlib/logger";
import UI from "../../src/userInteraction";

describe("UserInteraction(CLI) 2", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(UI, "createProgressBar").returns({
      start: async (s) => {},
      next: async (s) => {},
      end: async (s) => {},
    });
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("loadSelectDynamicData", async () => {
    it("happy path", async () => {
      const config: SingleSelectConfig = {
        name: "test",
        title: "test",
        options: async () => ["a", "b", "c"],
        default: async () => "a",
      };
      const result = await UI.loadSelectDynamicData(config);
      assert.isTrue(result.isOk());
      assert.deepEqual(config.options, ["a", "b", "c"]);
      assert.equal(config.default, "a");
    });
    it("throw error", async () => {
      const config: SingleSelectConfig = {
        name: "test",
        title: "test",
        options: async () => {
          throw new Error("test");
        },
      };
      const result = await UI.loadSelectDynamicData(config);
      assert.isTrue(result.isErr());
    });
    it("no need to call function", async () => {
      const config: SingleSelectConfig = {
        name: "test",
        title: "test",
        options: ["a", "b", "c"],
      };
      const result = await UI.loadSelectDynamicData(config);
      assert.isTrue(result.isOk());
      assert.deepEqual(config.options, ["a", "b", "c"]);
    });
  });

  describe("loadDefaultValue", async () => {
    it("happy path", async () => {
      const config: SingleSelectConfig = {
        name: "test",
        title: "test",
        options: ["a", "b", "c"],
        default: async () => "a",
      };
      const result = await UI.loadDefaultValue(config);
      assert.isTrue(result.isOk());
      assert.equal(config.default, "a");
    });
    it("throw error", async () => {
      const config: SingleSelectConfig = {
        name: "test",
        title: "test",
        options: ["a", "b", "c"],
        default: async () => {
          throw new Error("test");
        },
      };
      const result = await UI.loadDefaultValue(config);
      assert.isTrue(result.isErr());
    });
    it("no need to call function", async () => {
      const config: SingleSelectConfig = {
        name: "test",
        title: "test",
        options: ["a", "b", "c"],
        default: "a",
      };
      const result = await UI.loadDefaultValue(config);
      assert.isTrue(result.isOk());
    });
  });

  describe("selectOptions", () => {
    it("loadSelectDynamicData throw error", async () => {
      sandbox.stub(UI, "loadSelectDynamicData").resolves(err(new UserError({})));
      const config: MultiSelectConfig = {
        name: "test",
        title: "test",
        options: async () => {
          throw new Error("test");
        },
      };
      const result = await UI.selectOptions(config);
      assert.isTrue(result.isErr());
    });
    it("success with default=all", async () => {
      sandbox.stub(UI, "multiSelect").resolves(ok(["a", "b", "c"]));
      const config: MultiSelectConfig = {
        name: "test",
        title: "test",
        options: async () => {
          return ["a", "b", "c"];
        },
        default: "all",
      };
      const result = await UI.selectOptions(config);
      assert.isTrue(result.isOk());
    });
    it("success with default=all", async () => {
      sandbox.stub(UI, "multiSelect").resolves(ok(["a", "b"]));
      const config: MultiSelectConfig = {
        name: "test",
        title: "test",
        options: async () => {
          return [
            { id: "a", label: "a" },
            { id: "b", label: "b" },
          ];
        },
        default: "all",
      };
      const result = await UI.selectOptions(config);
      assert.isTrue(result.isOk());
    });
    it("success with default=none", async () => {
      sandbox.stub(UI, "multiSelect").resolves(ok([]));
      const config: MultiSelectConfig = {
        name: "test",
        title: "test",
        options: async () => {
          return ["a", "b", "c"];
        },
        default: "none",
      };
      const result = await UI.selectOptions(config);
      assert.isTrue(result.isOk());
    });
  });

  describe("selectOption", () => {
    it("loadSelectDynamicData throw error", async () => {
      sandbox.stub(UI, "loadSelectDynamicData").resolves(err(new UserError({})));
      const config: SingleSelectConfig = {
        name: "test",
        title: "test",
        options: async () => {
          throw new Error("test");
        },
      };
      const result = await UI.selectOption(config);
      assert.isTrue(result.isErr());
    });
    it("SelectSubscriptionError", async () => {
      const config: SingleSelectConfig = {
        name: "subscription",
        title: "select subscription",
        options: [],
      };
      const result = await UI.selectOption(config);
      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.isTrue(result.error instanceof SelectSubscriptionError);
      }
    });
    it("happy return options", async () => {
      const config: SingleSelectConfig = {
        name: "test",
        title: "test",
        options: ["a", "b"],
      };
      sandbox.stub(UI, "singleSelect").resolves(ok("a"));
      const result = await UI.selectOption(config);
      assert.isTrue(result.isOk());
      if (result.isOk()) {
        assert.deepEqual(result.value.options, ["a", "b"]);
      }
    });
  });

  describe("inputText", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("load default value error", async () => {
      const res = await UI.inputText({
        title: "test",
        name: "test",
        default: async () => {
          throw new Error();
        },
      });
      assert.isTrue(res.isErr());
    });
    it("UnhandledError", async () => {
      sandbox.stub(UI, "input").resolves(err(new UnhandledError(new Error("test"))));
      const config: InputTextConfig = {
        name: "testInput",
        title: "input text",
      };
      const result = await UI.inputText(config);
      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.isTrue(result.error instanceof UnhandledError);
      }
    });
  });

  describe("selectFile", () => {
    it("load default value error", async () => {
      const res = await UI.selectFile({
        title: "test",
        name: "test",
        default: async () => {
          throw new Error();
        },
      });
      assert.isTrue(res.isErr());
    });
  });

  describe("selectFiles", () => {
    it("load default value error", async () => {
      const res = await UI.selectFiles({
        title: "test",
        name: "test",
        default: async () => {
          throw new Error();
        },
      });
      assert.isTrue(res.isErr());
    });
  });

  describe("selectFolder", () => {
    it("load default value error", async () => {
      const res = await UI.selectFolder({
        title: "test",
        name: "test",
        default: async () => {
          throw new Error();
        },
      });
      assert.isTrue(res.isErr());
    });
  });

  describe("selectFileOrInput", () => {
    it("happy path", async () => {
      sandbox.stub(UI, "input").resolves(ok("somevalue"));
      const res = await UI.selectFileOrInput({
        name: "test",
        title: "test",
        inputBoxConfig: {
          title: "test",
          name: "test",
          validation: (input: string) => {
            return undefined;
          },
        },
        inputOptionItem: {
          id: "test",
          label: "test",
        },
      });
      assert.isTrue(res.isOk());
    });

    it("load default value error", async () => {
      const res = await UI.selectFileOrInput({
        name: "test",
        title: "test",
        inputBoxConfig: {
          title: "test",
          name: "test",
          default: async () => {
            throw new Error();
          },
        },
        inputOptionItem: {
          id: "test",
          label: "test",
        },
      });
      assert.isTrue(res.isErr());
    });
  });
});

describe("runCommand", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });
  it("happy path win32", async () => {
    const mockChildProcess = {
      on: sandbox.stub().callsFake((event, callback) => {
        if (event === "close") {
          callback(0); // Simulate successful execution
        }
      }),
    };
    sandbox.stub(process, "platform").value("win32");
    sandbox.stub(logger, "info").returns();
    const spawnStub = sandbox.stub(child_process, "spawn").returns(mockChildProcess as any);
    const res = await UI.runCommand({ cmd: 'echo "Hello"' });
    assert.isTrue(res.isOk());
    assert.isTrue(spawnStub.calledOnce);
    assert.equal(spawnStub.firstCall.args[0], "cmd.exe");
  });
  it("uses custom shell", async () => {
    const mockChildProcess = {
      on: sandbox.stub().callsFake((event, callback) => {
        if (event === "close") {
          callback(0);
        }
      }),
    };
    sandbox.stub(process, "platform").value("win32");
    sandbox.stub(logger, "info").returns();
    const spawnStub = sandbox.stub(child_process, "spawn").returns(mockChildProcess as any);
    const res = await UI.runCommand({ cmd: "echo hello", shell: "powershell.exe" });
    assert.isTrue(res.isOk());
    assert.isTrue(spawnStub.calledOnce);
    assert.equal(spawnStub.firstCall.args[0], "powershell.exe");
  });
  it("error linux", async () => {
    const mockChildProcess = {
      on: sandbox.stub().callsFake((event, callback) => {
        if (event === "close") {
          callback(1); // Simulate successful execution
        }
      }),
    };
    sandbox.stub(process, "platform").value("linux");
    sandbox.stub(logger, "info").returns();
    sandbox.stub(logger, "error").returns();
    const spawnStub = sandbox.stub(child_process, "spawn").returns(mockChildProcess as any);
    const res = await UI.runCommand({ cmd: 'echo "Hello"' });
    assert.isTrue(res.isErr());
    assert.isTrue(spawnStub.calledOnce);
    assert.equal(spawnStub.firstCall.args[0], "/bin/bash");
  });
  it("captures stdout data", async () => {
    const stdoutData = "Output from stdout";
    const mockChildProcess = {
      stdout: {
        on: sandbox.stub().callsFake((event, callback) => {
          if (event === "data") {
            callback(Buffer.from(stdoutData));
          }
        }),
      },
      stderr: {
        on: sandbox.stub(),
      },
      on: sandbox.stub().callsFake((event, callback) => {
        if (event === "close") {
          callback(0);
        }
      }),
    };
    sandbox.stub(process, "platform").value("win32");
    sandbox.stub(logger, "info").returns();
    sandbox.stub(process.stdout, "write").returns(true);
    sandbox.stub(process.stderr, "write").returns(true);
    sandbox.stub(child_process, "spawn").returns(mockChildProcess as any);
    const res = await UI.runCommand({ cmd: 'echo "test"' });
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.equal(res.value, stdoutData);
    }
  });
  it("captures stderr data", async () => {
    const stderrData = "Error from stderr";
    const mockChildProcess = {
      stdout: {
        on: sandbox.stub(),
      },
      stderr: {
        on: sandbox.stub().callsFake((event, callback) => {
          if (event === "data") {
            callback(Buffer.from(stderrData));
          }
        }),
      },
      on: sandbox.stub().callsFake((event, callback) => {
        if (event === "close") {
          callback(0);
        }
      }),
    };
    sandbox.stub(process, "platform").value("linux");
    sandbox.stub(logger, "info").returns();
    sandbox.stub(process.stdout, "write").returns(true);
    sandbox.stub(process.stderr, "write").returns(true);
    sandbox.stub(child_process, "spawn").returns(mockChildProcess as any);
    const res = await UI.runCommand({ cmd: 'echo "test"' });
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.equal(res.value, stderrData);
    }
  });
  it("captures both stdout and stderr data", async () => {
    const stdoutData = "Output from stdout\n";
    const stderrData = "Error from stderr\n";
    const mockChildProcess = {
      stdout: {
        on: sandbox.stub().callsFake((event, callback) => {
          if (event === "data") {
            callback(Buffer.from(stdoutData));
          }
        }),
      },
      stderr: {
        on: sandbox.stub().callsFake((event, callback) => {
          if (event === "data") {
            callback(Buffer.from(stderrData));
          }
        }),
      },
      on: sandbox.stub().callsFake((event, callback) => {
        if (event === "close") {
          callback(0);
        }
      }),
    };
    sandbox.stub(process, "platform").value("linux");
    sandbox.stub(logger, "info").returns();
    sandbox.stub(process.stdout, "write").returns(true);
    sandbox.stub(process.stderr, "write").returns(true);
    sandbox.stub(child_process, "spawn").returns(mockChildProcess as any);
    const res = await UI.runCommand({ cmd: 'echo "test"' });
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.equal(res.value, stdoutData + stderrData);
    }
  });
  it("passes workingDirectory to spawn", async () => {
    const workingDir = "/path/to/working/dir";
    const mockChildProcess = {
      stdout: { on: sandbox.stub() },
      stderr: { on: sandbox.stub() },
      on: sandbox.stub().callsFake((event, callback) => {
        if (event === "close") {
          callback(0);
        }
      }),
    };
    sandbox.stub(process, "platform").value("linux");
    sandbox.stub(logger, "info").returns();
    sandbox.stub(process.stdout, "write").returns(true);
    sandbox.stub(process.stderr, "write").returns(true);
    const spawnStub = sandbox.stub(child_process, "spawn").returns(mockChildProcess as any);
    const res = await UI.runCommand({ cmd: "pwd", workingDirectory: workingDir });
    assert.isTrue(res.isOk());
    assert.isTrue(spawnStub.calledOnce);
    const spawnOptions = spawnStub.firstCall.args[2];
    assert.equal(spawnOptions.cwd, workingDir);
  });
  it("passes timeout to spawn", async () => {
    const timeout = 5000;
    const mockChildProcess = {
      stdout: { on: sandbox.stub() },
      stderr: { on: sandbox.stub() },
      on: sandbox.stub().callsFake((event, callback) => {
        if (event === "close") {
          callback(0);
        }
      }),
    };
    sandbox.stub(process, "platform").value("linux");
    sandbox.stub(logger, "info").returns();
    sandbox.stub(process.stdout, "write").returns(true);
    sandbox.stub(process.stderr, "write").returns(true);
    const spawnStub = sandbox.stub(child_process, "spawn").returns(mockChildProcess as any);
    const res = await UI.runCommand({ cmd: "sleep 1", timeout: timeout });
    assert.isTrue(res.isOk());
    assert.isTrue(spawnStub.calledOnce);
    const spawnOptions = spawnStub.firstCall.args[2];
    assert.equal(spawnOptions.timeout, timeout);
  });
  it("passes env variables to spawn", async () => {
    const envVars = { TEST_VAR: "test_value", ANOTHER_VAR: "another_value" };
    const mockChildProcess = {
      stdout: { on: sandbox.stub() },
      stderr: { on: sandbox.stub() },
      on: sandbox.stub().callsFake((event, callback) => {
        if (event === "close") {
          callback(0);
        }
      }),
    };
    sandbox.stub(process, "platform").value("win32");
    sandbox.stub(logger, "info").returns();
    sandbox.stub(process.stdout, "write").returns(true);
    sandbox.stub(process.stderr, "write").returns(true);
    const spawnStub = sandbox.stub(child_process, "spawn").returns(mockChildProcess as any);
    const res = await UI.runCommand({ cmd: "echo %TEST_VAR%", env: envVars });
    assert.isTrue(res.isOk());
    assert.isTrue(spawnStub.calledOnce);
    const spawnOptions = spawnStub.firstCall.args[2];
    assert.deepEqual(spawnOptions.env, envVars);
  });
  it("handles non-zero exit code with output", async () => {
    const errorOutput = "Command failed with error";
    const mockChildProcess = {
      stdout: {
        on: sandbox.stub().callsFake((event, callback) => {
          if (event === "data") {
            callback(Buffer.from(errorOutput));
          }
        }),
      },
      stderr: { on: sandbox.stub() },
      on: sandbox.stub().callsFake((event, callback) => {
        if (event === "close") {
          callback(2); // Non-zero exit code
        }
      }),
    };
    sandbox.stub(process, "platform").value("linux");
    sandbox.stub(logger, "info").returns();
    sandbox.stub(logger, "error").returns();
    sandbox.stub(process.stdout, "write").returns(true);
    sandbox.stub(process.stderr, "write").returns(true);
    sandbox.stub(child_process, "spawn").returns(mockChildProcess as any);
    const res = await UI.runCommand({ cmd: "invalid-command" });
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.include(res.error.message, "Execute task failed with exit code:2");
    }
  });
});
