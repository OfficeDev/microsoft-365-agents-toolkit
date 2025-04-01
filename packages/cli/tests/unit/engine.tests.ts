import {
  CLICommand,
  CLICommandOption,
  CLIContext,
  CLIFoundCommand,
  LogLevel,
  SystemError,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import {
  featureFlagManager,
  FxCore,
  InputValidationError,
  MissingEnvironmentVariablesError,
  UserCancelError,
  VersionState,
} from "@microsoft/teamsfx-core";
import { assert } from "chai";
import "mocha";
import * as sinon from "sinon";
import * as activate from "../../src/activate";
import { getFxCore, resetFxCore } from "../../src/activate";
import { engine } from "../../src/commands/engine";
import { start } from "../../src/commands/index";
import {
  listSamplesCommand,
  listTemplatesCommand,
  m365SideloadingCommand,
} from "../../src/commands/models";
import { getCreateCommand } from "../../src/commands/models/create";
import { createSampleCommand } from "../../src/commands/models/createSample";
import { rootCommand } from "../../src/commands/models/root";
import { logger } from "../../src/commonlib/logger";
import { CliTelemetryReporter } from "../../src/commonlib/telemetry";
import {
  InvalidChoiceError,
  UnknownArgumentError,
  UnknownCommandError,
  UnknownOptionError,
} from "../../src/error";
import * as main from "../../src/index";
import CliTelemetry from "../../src/telemetry/cliTelemetry";
import { getVersion } from "../../src/utils";
import mockedEnv, { RestoreFn } from "mocked-env";
import { shareCommand } from "../../src/commands/models/share";
import { setSensitivityLabelCommand } from "../../src/commands/models/setSensitivityLabel";

describe("CLI Engine", () => {
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(process, "exit");
    sandbox.stub(CliTelemetry, "flush").resolves();
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("findCommand", async () => {
    it("should find new sample command", async () => {
      const result = engine.findCommand(rootCommand, ["new", "sample"]);
      assert.equal(result.cmd.name, createSampleCommand.name);
      assert.deepEqual(result.remainingArgs, []);
    });
    it("should find sideloading command alias", async () => {
      const result = engine.findCommand(rootCommand, ["sideloading"]);
      assert.equal(result.cmd.name, m365SideloadingCommand.name);
      assert.deepEqual(result.remainingArgs, []);
    });

    it("should not find share command if feature flag is off", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      const result = engine.findCommand(rootCommand, ["share"]);
      assert.equal(result.cmd.name, rootCommand.name);
      assert.deepEqual(result.remainingArgs, ["share"]);
    });
    it("should find share command if feature flag is on", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      const result = engine.findCommand(rootCommand, ["share"]);
      assert.equal(result.cmd.name, shareCommand.name);
      assert.deepEqual(result.remainingArgs, []);
    });

    it("should not find set sensitivity label command if feature flag is off", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      const result = engine.findCommand(rootCommand, ["set", "sensitivityLabel"]);
      assert.equal(result.cmd.name, rootCommand.name);
      assert.deepEqual(result.remainingArgs, ["set", "sensitivityLabel"]);
    });
    it("should find set sensitivity label command if feature flag is on", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      const result = engine.findCommand(rootCommand, ["set", "sensitivityLabel"]);
      assert.equal(result.cmd.name, setSensitivityLabelCommand.name);
      assert.deepEqual(result.remainingArgs, []);
    });
  });
  describe("parseArgs", async () => {
    it("array type options", async () => {
      const mockedEnvRestore = mockedEnv({
        CI_ENABLED: "true",
      });
      const command: CLIFoundCommand = {
        name: "test",
        fullName: "test",
        description: "test command",
        options: [
          {
            type: "array",
            name: "option1",
            description: "test option",
          },
        ],
      };
      const ctx: CLIContext = {
        command: command,
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const result = engine.parseArgs(ctx, rootCommand, ["--option1", "a,b,c"]);
      assert.isTrue(result.isOk());
      assert.deepEqual(ctx.optionValues["option1"], ["a", "b", "c"]);
      assert.isFalse(ctx.globalOptionValues.interactive);
      mockedEnvRestore();
    });
    it("array type options 2", async () => {
      const command: CLIFoundCommand = {
        name: "test",
        fullName: "test",
        description: "test command",
        options: [
          {
            type: "array",
            name: "option1",
            description: "test option",
          },
        ],
      };
      const ctx: CLIContext = {
        command: command,
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const result = engine.parseArgs(ctx, rootCommand, [
        "--option1",
        "a",
        "--option1",
        "b",
        "--option1",
        "c",
      ]);
      assert.isTrue(result.isOk());
      assert.deepEqual(ctx.optionValues["option1"], ["a", "b", "c"]);
    });
    it("array type options 3", async () => {
      const command: CLIFoundCommand = {
        name: "test",
        fullName: "test",
        description: "test command",
        options: [
          {
            type: "array",
            name: "option1",
            description: "test option",
          },
        ],
      };
      const ctx: CLIContext = {
        command: command,
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const result = engine.parseArgs(ctx, rootCommand, ["--option1=a,b,c"]);
      assert.isTrue(result.isOk());
      assert.deepEqual(ctx.optionValues["option1"], ["a", "b", "c"]);
    });
    it("array type argument", async () => {
      const command: CLIFoundCommand = {
        name: "test",
        fullName: "test",
        description: "test command",
        arguments: [
          {
            type: "array",
            name: "arg1",
            description: "test argument",
          },
          {
            type: "string",
            name: "arg2",
            description: "test argument2",
            required: true,
            default: "default",
          },
        ],
      };
      const ctx: CLIContext = {
        command: command,
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const result = engine.parseArgs(ctx, rootCommand, ["a,b,c"]);
      assert.isTrue(result.isOk());
      assert.deepEqual(ctx.argumentValues[0], ["a", "b", "c"]);
      assert.equal(ctx.argumentValues[1], "default");
    });
    it("boolean type option", async () => {
      const command: CLIFoundCommand = {
        name: "test",
        fullName: "test",
        description: "test command",
        options: [
          {
            type: "boolean",
            name: "option1",
            description: "test option",
          },
        ],
      };
      const ctx: CLIContext = {
        command: command,
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const result = engine.parseArgs(ctx, rootCommand, ["--option1", "true"]);
      assert.isTrue(result.isOk());
      assert.equal(ctx.optionValues["option1"], true);
    });
    it("UnknownCommandError", async () => {
      const command: CLIFoundCommand = {
        name: "test",
        fullName: "test",
        description: "test command",
        options: [
          {
            type: "boolean",
            name: "option1",
            description: "test option",
          },
        ],
        commands: [
          {
            name: "subcommand",
            description: "test",
          },
        ],
      };
      const ctx: CLIContext = {
        command: command,
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const result = engine.parseArgs(ctx, rootCommand, ["subcomand"]);
      assert.isTrue(
        result.isErr() &&
          result.error instanceof UnknownCommandError &&
          result.error.message.includes("subcomand")
      );
    });
    it("UnknownArgumentError", async () => {
      const command: CLIFoundCommand = {
        name: "test",
        fullName: "test",
        description: "test command",
        arguments: [
          {
            type: "boolean",
            name: "option1",
            description: "test option",
          },
        ],
      };
      const ctx: CLIContext = {
        command: command,
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const result = engine.parseArgs(ctx, rootCommand, ["abc", "def"]);
      assert.isTrue(result.isErr() && result.error instanceof UnknownArgumentError);
    });
  });
  describe("validateOption", async () => {
    it("InvalidChoiceError", async () => {
      const option: CLICommandOption = {
        type: "array",
        description: "test",
        name: "test",
        choices: ["a", "b", "c"],
        value: ["d"],
      };
      const result = engine.validateOption(
        { name: "test", fullName: "test", description: "" },
        option,
        "option"
      );
      assert.isTrue(result.isErr() && result.error instanceof InvalidChoiceError);
    });
  });
  describe("isTelemetryEnabled", async () => {
    it("true", async () => {
      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "abc" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = engine.isTelemetryEnabled(ctx);
      assert.isTrue(res);
    });
    it("true", async () => {
      const res = engine.isTelemetryEnabled();
      assert.isTrue(res);
    });
    it("false", async () => {
      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "abc" },
        optionValues: {},
        globalOptionValues: { telemetry: false },
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = engine.isTelemetryEnabled(ctx);
      assert.isFalse(res);
    });
  });
  describe("processResult", async () => {
    it("sendTelemetryErrorEvent", async () => {
      const sendTelemetryErrorEventStub = sandbox
        .stub(CliTelemetry, "sendTelemetryErrorEvent")
        .returns();
      sandbox.stub(logger, "outputError").returns();
      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "abc" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      await engine.processResult(ctx, new InputValidationError("test", "no reason"));
      assert.isTrue(sendTelemetryErrorEventStub.calledOnce);
    });
    it("sendTelemetryEvent", async () => {
      const sendTelemetryEventStub = sandbox.stub(CliTelemetry, "sendTelemetryEvent").returns();
      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "abc" },
        optionValues: { env: "dev" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      await engine.processResult(ctx, undefined);
      assert.isTrue(sendTelemetryEventStub.calledOnce);
    });
    it("skip telemetry when reporter is disabled", async () => {
      CliTelemetry.reporter = new CliTelemetryReporter("real", "real", "real", "real");
      CliTelemetry.enable = false;
      const spy = sandbox.spy(CliTelemetry.reporter.reporter, "sendTelemetryEvent");
      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "abc" },
        optionValues: {},
        globalOptionValues: { telemetry: false },
        argumentValues: [],
        telemetryProperties: {},
      };
      await engine.processResult(ctx, undefined);
      assert.isTrue(spy.notCalled);
    });
    it("skip telemetry when context is undefined", async () => {
      CliTelemetry.reporter = new CliTelemetryReporter("real", "real", "real", "real");
      CliTelemetry.enable = false;
      const spy = sandbox.spy(CliTelemetry.reporter.reporter, "sendTelemetryEvent");
      await engine.processResult(undefined, undefined);
      assert.isTrue(spy.notCalled);
    });
    it("skip telemetry when command telemetry is undefined", async () => {
      CliTelemetry.reporter = new CliTelemetryReporter("real", "real", "real", "real");
      CliTelemetry.enable = false;
      const spy = sandbox.spy(CliTelemetry.reporter.reporter, "sendTelemetryEvent");
      const command: CLICommand = {
        name: "test",
        description: "test",
      };
      const ctx: CLIContext = {
        command: { ...command, fullName: "test" },
        optionValues: {},
        globalOptionValues: { telemetry: false },
        argumentValues: [],
        telemetryProperties: {},
      };
      await engine.processResult(ctx, undefined);
      assert.isTrue(spy.notCalled);
    });
  });
  describe("start", async () => {
    it("command not found", async () => {
      sandbox.stub(process, "argv").value(["node", "cli", "abc123"]);
      const stub = sandbox.stub(engine, "printError").returns();
      await engine.start(rootCommand);
      assert.isTrue(stub.called);
    });
    it("command has no handler", async () => {
      sandbox.stub(process, "argv").value(["node", "cli", "list", "templates"]);
      sandbox.stub(listTemplatesCommand, "handler").value(undefined);
      await engine.start(rootCommand);
    });
    it("parseArg return error", async () => {
      sandbox.stub(process, "argv").value(["node", "cli", "new", "--xxx"]);
      let error;
      sandbox.stub(engine, "processResult").callsFake(async (ctx, fxError) => {
        error = fxError;
      });
      await engine.start(rootCommand);
      assert.instanceOf(error, UnknownOptionError);
    });
    it("should display version", async () => {
      sandbox.stub(process, "argv").value(["node", "cli", "--version"]);
      const loggerStub = sandbox.stub(logger, "info");
      await engine.start(rootCommand);
      assert.isTrue(loggerStub.calledWith(getVersion()));
    });
    it("should display help message", async () => {
      sandbox.stub(process, "argv").value(["node", "cli", "-h"]);
      const loggerStub = sandbox.stub(logger, "info");
      await engine.start(rootCommand);
      assert.isTrue(loggerStub.calledOnce);
    });
    it("should validation failed for capability", async () => {
      sandbox
        .stub(process, "argv")
        .value(["node", "cli", "new", "-c", "tab", "-n", "myapp", "-i", "false"]);
      let error: any = {};
      sandbox.stub(engine, "processResult").callsFake(async (context, fxError) => {
        error = fxError;
      });
      await engine.start(rootCommand);
      assert.isTrue(error && error instanceof InvalidChoiceError);
    });
    it("should run command with argument success", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "createSampleProject").resolves(ok({ projectPath: "..." }));
      sandbox
        .stub(process, "argv")
        .value(["node", "cli", "new", "sample", "hello-world-tab-with-backend", "-i", "false"]);
      const loggerStub = sandbox.stub(logger, "info");
      await engine.start(rootCommand);
      assert.isTrue(loggerStub.calledOnce);
    });
    it("should validate argument failed", async () => {
      sandbox.stub(createSampleCommand, "arguments").value([
        {
          type: "string",
          name: "sample",
          description: "Select a sample app to create",
          choices: ["a", "b", "c"],
        },
      ]);
      sandbox.stub(FxCore.prototype, "createSampleProject").resolves(ok({ projectPath: "..." }));
      sandbox.stub(process, "argv").value(["node", "cli", "new", "sample", "d", "-i", "false"]);
      let error: any = {};
      sandbox.stub(engine, "processResult").callsFake(async (context, fxError) => {
        error = fxError;
      });
      sandbox.stub(logger, "info");
      await engine.start(rootCommand);
      assert.isTrue(error instanceof InvalidChoiceError);
    });
    it("should discard useless args and options for interactive mode", async () => {
      sandbox.stub(FxCore.prototype, "createSampleProject").resolves(ok({ projectPath: "..." }));
      sandbox.stub(process, "argv").value(["node", "cli", "new", "sample", "abc"]);
      const stub = sandbox.stub(logger, "info");
      await engine.start(rootCommand);
      assert.isTrue(stub.called);
    });
    it("should run handler return error", async () => {
      sandbox.stub(process, "argv").value(["node", "cli"]);
      const command: CLIFoundCommand = {
        name: "test",
        description: "test",
        fullName: "test",
        handler: async () => err(new UserCancelError()),
      };
      let error: any = {};
      sandbox.stub(engine, "processResult").callsFake(async (context, fxError) => {
        error = fxError;
      });
      await engine.start(command);
      assert.isTrue(error instanceof UserCancelError);
    });
    it("should run handler throw error", async () => {
      sandbox.stub(process, "argv").value(["node", "cli"]);
      const command: CLIFoundCommand = {
        name: "test",
        description: "test",
        fullName: "test",
        handler: async () => {
          throw new UserCancelError();
        },
      };
      let error: any = {};
      sandbox.stub(engine, "processResult").callsFake(async (context, fxError) => {
        error = fxError;
      });
      await engine.start(command);
      assert.isTrue(error instanceof UserCancelError);
    });
    it("run version check and return error", async () => {
      sandbox.stub(FxCore.prototype, "projectVersionCheck").resolves(err(new UserCancelError()));
      sandbox.stub(process, "argv").value(["node", "cli", "provision", "--folder", "abc"]);
      let error: any = {};
      sandbox.stub(engine, "processResult").callsFake(async (context, fxError) => {
        error = fxError;
      });
      await engine.start(rootCommand);
      assert.isTrue(error instanceof UserCancelError);
    });
    it("run version check and return upgradeable and upgrade return error", async () => {
      sandbox.stub(FxCore.prototype, "projectVersionCheck").resolves(
        ok({
          isSupport: VersionState.upgradeable,
          currentVersion: "1",
          trackingId: "1",
          versionSource: "1",
        })
      );
      sandbox.stub(FxCore.prototype, "phantomMigrationV3").resolves(err(new UserCancelError()));
      sandbox.stub(process, "argv").value(["node", "cli", "provision", "--folder", "abc"]);
      let error: any = {};
      sandbox.stub(engine, "processResult").callsFake(async (context, fxError) => {
        error = fxError;
      });
      await engine.start(rootCommand);
      assert.isTrue(error instanceof UserCancelError);
    });
    it("skip options in interactive mode", async () => {
      sandbox.stub(FxCore.prototype, "createProject").resolves(ok({} as any));
      sandbox.stub(process, "argv").value(["node", "cli", "new", "--folder", "abc"]);
      let error: any = undefined;
      sandbox.stub(engine, "processResult").callsFake(async (context, fxError) => {
        error = fxError;
      });
      await engine.start(rootCommand);
      assert.isUndefined(undefined);
    });
    it("skip arguments in interactive mode", async () => {
      sandbox.stub(FxCore.prototype, "createSampleProject").resolves(ok({} as any));
      sandbox.stub(process, "argv").value(["node", "cli", "new", "sample", "abc"]);
      let error: any = undefined;
      sandbox.stub(engine, "processResult").callsFake(async (context, fxError) => {
        error = fxError;
      });
      await engine.start(rootCommand);
      assert.isUndefined(undefined);
    });
    it("no need to skip options or arguments in interactive mode", async () => {
      sandbox.stub(FxCore.prototype, "createProject").resolves(ok({} as any));
      sandbox.stub(process, "argv").value(["node", "cli", "new"]);
      let error: any = undefined;
      sandbox.stub(engine, "processResult").callsFake(async (context, fxError) => {
        error = fxError;
      });
      await engine.start(rootCommand);
      assert.isUndefined(undefined);
    });
    it("use defaultInteractiveOption", async () => {
      const comand = listSamplesCommand;
      sandbox.stub(comand, "handler").resolves(ok(undefined));
      sandbox.stub(process, "argv").value(["node", "cli", "list", "samples"]);
      let error: any = undefined;
      sandbox.stub(engine, "processResult").callsFake(async (context, fxError) => {
        error = fxError;
      });
      await engine.start(rootCommand);
      assert.isUndefined(undefined);
    });
  });
  describe("index.start", async () => {
    it("happy path", async () => {
      sandbox.stub(main, "initTelemetryReporter").returns();
      sandbox.stub(engine, "start").resolves();
      await start();
      assert.isTrue(true);
    });
  });
  describe("getFxCore", async () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("new logger", async () => {
      resetFxCore();
      getFxCore();
    });
  });
  describe("printError", async () => {
    it("happy path user error", async () => {
      sandbox.stub(logger, "info").resolves();
      sandbox.stub(logger, "debug").resolves();
      const stub = sandbox.stub(logger, "outputError").returns();
      engine.printError(new MissingEnvironmentVariablesError("test", "test"));
      assert.isTrue(stub.called);
    });
    it("happy path system error", async () => {
      sandbox.stub(logger, "logLevel").value(LogLevel.Debug);
      const stub = sandbox.stub(logger, "debug").resolves();
      sandbox.stub(logger, "outputError").returns();
      const error = new SystemError({ issueLink: "http://aka.ms/teamsfx-cli-help" });
      engine.printError(error);
      assert.isTrue(stub.called);
    });
    it("happy path inner error", async () => {
      sandbox.stub(logger, "logLevel").value(LogLevel.Debug);
      const stub = sandbox.stub(logger, "debug").resolves();
      sandbox.stub(logger, "outputError").returns();
      const error = new SystemError({ issueLink: "http://aka.ms/teamsfx-cli-help" });
      const innerError = new Error("test");
      error.innerError = innerError;
      error.message = "";
      error.stack = undefined;
      engine.printError(error);
      innerError.stack = undefined;
      engine.printError(error);
      assert.isTrue(stub.called);
    });
    it("canceled", async () => {
      const stub = sandbox.stub(logger, "info").resolves();
      engine.printError(new UserCancelError("test"));
      assert.isTrue(stub.called);
    });
  });
});
