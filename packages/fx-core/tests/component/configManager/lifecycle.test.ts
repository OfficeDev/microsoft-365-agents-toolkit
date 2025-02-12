// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yefuwang@microsoft.com
 */

import { assert } from "chai";
import { describe, it } from "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import sinon from "sinon";
import { Lifecycle } from "../../../src/component/configManager/lifecycle";
import Container from "typedi";
import { DriverDefinition } from "../../../src/component/configManager/interface";
import {
  MockedLogProvider,
  MockedTelemetryReporter,
  MockedUserInteraction,
} from "../../plugins/solution/util";
import { DriverContext } from "../../../src/component/driver/interface/commonArgs";
import {
  Platform,
  Result,
  FxError,
  ok,
  err,
  SystemError,
  IProgressHandler,
} from "@microsoft/teamsfx-api";
import { ExecutionResult, StepDriver } from "../../../src/component/driver/interface/stepDriver";
import { SummaryConstant } from "../../../src/component/configManager/constant";
import { MockedAzureAccountProvider, MockedM365Provider } from "../../core/utils";

const mockedDriverContext: DriverContext = {
  m365TokenProvider: new MockedM365Provider(),
  azureAccountProvider: new MockedAzureAccountProvider(),
  ui: new MockedUserInteraction(),
  progressBar: undefined,
  logProvider: new MockedLogProvider(),
  telemetryReporter: new MockedTelemetryReporter(),
  projectPath: "",
  platform: Platform.VSCode,
};

class DriverA implements StepDriver {
  progressTitle = "mocked progress title";

  async execute(args: unknown, ctx: DriverContext): Promise<ExecutionResult> {
    return {
      result: ok(new Map([["OUTPUT_A", "VALUE_A"]])),
      summaries: [],
    };
  }
}

class DriverAWithSummary extends DriverA {
  async execute(args: unknown, ctx: DriverContext): Promise<ExecutionResult> {
    return {
      result: ok(new Map([["OUTPUT_A", "VALUE_A"]])),
      summaries: ["Environment variable OUTPUT_A set in env/.env file"],
    };
  }
}

class DriverB implements StepDriver {
  progressTitle = "mocked progress title";

  async execute(args: unknown, ctx: DriverContext): Promise<ExecutionResult> {
    return {
      result: ok(new Map([["OUTPUT_B", "VALUE_B"]])),
      summaries: [],
    };
  }
}

class DriverBWithSummary extends DriverB {
  async execute(args: unknown, ctx: DriverContext): Promise<ExecutionResult> {
    return {
      result: ok(new Map([["OUTPUT_B", "VALUE_B"]])),
      summaries: ["Environment variable OUTPUT_B set in env/.env file"],
    };
  }
}

class DriverThatCapitalize implements StepDriver {
  async execute(args: { INPUT_A: string }, ctx: DriverContext): Promise<ExecutionResult> {
    return {
      result: ok(new Map([["OUTPUT", args.INPUT_A.toUpperCase()]])),
      summaries: [],
    };
  }
}

class DriverThatCapitalizeWithSummary extends DriverThatCapitalize {
  async execute(args: { INPUT_A: string }, ctx: DriverContext): Promise<ExecutionResult> {
    return {
      result: ok(new Map([["OUTPUT", args.INPUT_A.toUpperCase()]])),
      summaries: ["Environment variable OUTPUT set in env/.env file"],
    };
  }
}

class DriverThatLowercase implements StepDriver {
  async execute(args: { INPUT_A: string }, ctx: DriverContext): Promise<ExecutionResult> {
    return {
      result: ok(new Map([["OUTPUT_C", args.INPUT_A.toLowerCase()]])),
      summaries: [],
    };
  }
}

class DriverThatHasNestedArgs implements StepDriver {
  async execute(args: { key: [{ key1: string }] }, ctx: DriverContext): Promise<ExecutionResult> {
    return {
      result: ok(new Map([["OUTPUT_D", args.key.map((e) => e.key1).join(",")]])),
      summaries: [],
    };
  }
}

class DriverThatReturnsErrorWithSummary implements StepDriver {
  async execute(args: unknown, context: DriverContext): Promise<ExecutionResult> {
    const fxError: FxError = {
      name: "fakeError",
      message: "fake message",
      source: "xxx",
      timestamp: new Date(),
    };
    return { result: err(fxError), summaries: [] };
  }
}

const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");

class DriverThatUsesEnvField implements StepDriver {
  progressTitle = "mocked progress title";

  async execute(args: unknown, context: DriverContext): Promise<ExecutionResult> {
    if (process.env["ENV_VAR1"]) {
      return {
        result: ok(new Map([["OUTPUT_E", process.env["ENV_VAR1"]]])),
        summaries: [],
      };
    } else {
      return { result: err(mockedError), summaries: [] };
    }
  }
}

class DriverThatUsesWriteToEnvironmentFileField implements StepDriver {
  async run(args: unknown, context: DriverContext): Promise<Result<Map<string, string>, FxError>> {
    throw new Error(`not implemented`);
  }

  async execute(
    args: unknown,
    context: DriverContext,
    outputVarNames: Map<string, string>
  ): Promise<ExecutionResult> {
    const ret = [...outputVarNames.values()].map(
      (value) => [value, value.toLocaleLowerCase()] as const
    );
    return { result: ok(new Map([...ret])), summaries: [] };
  }
}

describe("v3 lifecyle", () => {
  describe("when driver name not found", () => {
    const sandbox = sinon.createSandbox();
    before(() => {
      sandbox.stub(Container, "has").returns(false);
    });

    afterEach(() => {
      sandbox.restore();
    });
    it("should return error", async () => {
      const driverDefs: DriverDefinition[] = [];
      driverDefs.push({
        name: "xxx",
        uses: "xxx",
        with: {},
      });

      const lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");

      const { result: execResult, summaries } = await lifecycle.execute(mockedDriverContext);
      assert(
        execResult.isErr() &&
          execResult.error.kind === "Failure" &&
          execResult.error.error.name === "InvalidYmlActionNameError"
      );

      assert(summaries.length === 0, "summary list should be empty");
    });
  });

  describe("when run/execute with multiple drivers", () => {
    const sandbox = sinon.createSandbox();
    before(() => {
      sandbox
        .stub(Container, "has")
        .withArgs(sandbox.match("DriverA"))
        .returns(true)
        .withArgs(sandbox.match("DriverB"))
        .returns(true);
      sandbox
        .stub(Container, "get")
        .withArgs(sandbox.match("DriverA"))
        .returns(new DriverA())
        .withArgs(sandbox.match("DriverB"))
        .returns(new DriverB());
    });

    after(() => {
      mockedDriverContext.progressBar = undefined;
      sandbox.restore();
    });

    it("should return combined output", async () => {
      const driverDefs: DriverDefinition[] = [];
      driverDefs.push({
        name: "xxx",
        uses: "DriverA",
        with: {},
      });
      driverDefs.push({
        name: "xxx",
        uses: "DriverB",
        with: {},
      });

      const lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");
      const nextStub = sandbox.stub();
      mockedDriverContext.progressBar = {
        next: nextStub,
      } as unknown as IProgressHandler;

      const { result: execResult, summaries } = await lifecycle.execute(mockedDriverContext);
      assert(
        execResult.isOk() &&
          execResult.value.size === 2 &&
          execResult.value.get("OUTPUT_A") === "VALUE_A" &&
          execResult.value.get("OUTPUT_B") === "VALUE_B"
      );

      assert(
        summaries.length === 2 && summaries[0].length === 0 && summaries[1].length === 0,
        "summary list should have 2 empty items, since DriverA and DriverB doesn't implement execute()"
      );

      assert.isTrue(nextStub.calledTwice);
    });
  });

  describe("when run/execute with valid placeholders", async () => {
    const sandbox = sinon.createSandbox();
    let restoreFn: RestoreFn | undefined = undefined;

    before(() => {
      restoreFn = mockedEnv({
        SOME_ENV_VAR: "xxx",
      });
      sandbox.stub(Container, "has").withArgs(sandbox.match("DriverThatCapitalize")).returns(true);
      sandbox
        .stub(Container, "get")
        .withArgs(sandbox.match("DriverThatCapitalize"))
        .returns(new DriverThatCapitalize());
    });

    after(() => {
      if (restoreFn) {
        restoreFn();
      }
      sandbox.restore();
    });

    it("should replace all placeholders", async () => {
      const driverDefs: DriverDefinition[] = [];
      driverDefs.push({
        uses: "DriverThatCapitalize",
        with: { INPUT_A: "hello ${{ SOME_ENV_VAR }}" },
      });

      const lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");

      const { result: execResult, summaries } = await lifecycle.execute(mockedDriverContext);
      assert(execResult.isOk() && execResult.value.get("OUTPUT") === "HELLO XXX");
      assert(summaries.length === 1 && summaries[0].length === 0);

      assert((driverDefs[0].with as any).INPUT_A === "hello xxx");
    });
  });

  describe("when dealing with multiple valid placeholders", async () => {
    const sandbox = sinon.createSandbox();
    let restoreFn: RestoreFn | undefined = undefined;

    before(() => {
      restoreFn = mockedEnv({
        SOME_ENV_VAR: "xxx",
        OTHER_ENV_VAR: "yyy",
      });
      sandbox
        .stub(Container, "has")
        .withArgs(sandbox.match("DriverThatCapitalize"))
        .returns(true)
        .withArgs(sandbox.match("DriverThatLowercase"))
        .returns(true)
        .withArgs(sandbox.match("DriverThatHasNestedArgs"))
        .returns(true)
        .withArgs(sandbox.match("DriverThatUsesEnvField"))
        .returns(true);
      sandbox
        .stub(Container, "get")
        .withArgs(sandbox.match("DriverThatCapitalize"))
        .returns(new DriverThatCapitalize())
        .withArgs(sandbox.match("DriverThatLowercase"))
        .returns(new DriverThatLowercase())
        .withArgs(sandbox.match("DriverThatHasNestedArgs"))
        .returns(new DriverThatHasNestedArgs())
        .withArgs(sandbox.match("DriverThatUsesEnvField"))
        .returns(new DriverThatUsesEnvField());
    });

    after(() => {
      if (restoreFn) {
        restoreFn();
      }
      sandbox.restore();
    });

    it("should replace all placeholders for a single driver", async () => {
      let driverDefs: DriverDefinition[] = [];
      driverDefs.push({
        uses: "DriverThatCapitalize",
        with: { INPUT_A: "hello ${{ SOME_ENV_VAR }} and ${{OTHER_ENV_VAR}}" },
      });

      let lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");

      driverDefs = [];
      driverDefs.push({
        uses: "DriverThatCapitalize",
        with: { INPUT_A: "hello ${{ SOME_ENV_VAR }} and ${{OTHER_ENV_VAR}}" },
      });

      lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");
      const { result: execResult, summaries } = await lifecycle.execute(mockedDriverContext);
      assert(execResult.isOk() && execResult.value.get("OUTPUT") === "HELLO XXX AND YYY");
      assert(summaries.length === 1 && summaries[0].length === 0);
    });

    it("should replace all placeholders for every driver", async () => {
      let driverDefs: DriverDefinition[] = [];
      driverDefs.push({
        uses: "DriverThatCapitalize",
        with: { INPUT_A: "hello ${{ SOME_ENV_VAR }}" },
      });
      driverDefs.push({
        uses: "DriverThatLowercase",
        with: { INPUT_A: "Hello ${{OTHER_ENV_VAR}}" },
      });

      let lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");

      driverDefs = [];
      driverDefs.push({
        uses: "DriverThatCapitalize",
        with: { INPUT_A: "hello ${{ SOME_ENV_VAR }}" },
      });
      driverDefs.push({
        uses: "DriverThatLowercase",
        with: { INPUT_A: "Hello ${{OTHER_ENV_VAR}}" },
      });

      lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");
      const { result: execResult, summaries } = await lifecycle.execute(mockedDriverContext);
      assert(
        execResult.isOk() &&
          execResult.value.get("OUTPUT") === "HELLO XXX" &&
          execResult.value.get("OUTPUT_C") === "hello yyy"
      );

      assert(summaries.length === 2 && summaries[0].length === 0 && summaries[1].length === 0);
    });

    it("should replace all placeholders for every driver with nested args", async () => {
      let driverDefs: DriverDefinition[] = [];
      driverDefs.push({
        uses: "DriverThatHasNestedArgs",
        with: {
          key: [{ key1: "hello ${{ SOME_ENV_VAR }}" }, { key1: "hello ${{ OTHER_ENV_VAR }}" }],
        },
      });

      let lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");

      driverDefs = [];
      driverDefs.push({
        uses: "DriverThatHasNestedArgs",
        with: {
          key: [{ key1: "hello ${{ SOME_ENV_VAR }}" }, { key1: "hello ${{ OTHER_ENV_VAR }}" }],
        },
      });

      lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");
      const { result: execResult, summaries } = await lifecycle.execute(mockedDriverContext);
      assert(execResult.isOk() && execResult.value.get("OUTPUT_D") === "hello xxx,hello yyy");
      assert(summaries.length === 1 && summaries[0].length === 0);
    });

    describe("execute()", async () => {
      it("should resolve inter-driver dependency", async () => {
        const driverDefs: DriverDefinition[] = [];
        driverDefs.push({
          uses: "DriverThatCapitalize",
          with: { INPUT_A: "hello ${{ SOME_ENV_VAR }}" },
        });
        // OUTPUT is a placeholder for the output of the previous driver
        driverDefs.push({
          uses: "DriverThatLowercase",
          with: { INPUT_A: "Hello ${{OUTPUT}}" },
        });

        const lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");
        const { result, summaries } = await lifecycle.execute(mockedDriverContext);
        assert(
          result.isOk() &&
            result.value.get("OUTPUT") === "HELLO XXX" &&
            result.value.get("OUTPUT_C") === "hello hello xxx"
        );

        assert(summaries.length === 2 && summaries[0].length === 0 && summaries[1].length === 0);
      });

      it("should resolve placeholders in env field", async () => {
        const driverDefs: DriverDefinition[] = [];
        driverDefs.push({
          uses: "DriverThatUsesEnvField",
          with: {},
          env: {
            ENV_VAR1: "hello ${{ SOME_ENV_VAR }}",
          },
        });

        const lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");
        const { result, summaries } = await lifecycle.execute(mockedDriverContext);
        assert(result.isOk() && result.value.get("OUTPUT_E") === "hello xxx");
        assert(summaries.length === 1 && summaries[0].length === 0);
      });
    });
  });

  describe("when dealing with unresolved placeholders", async () => {
    const sandbox = sinon.createSandbox();
    let restoreFn: RestoreFn | undefined = undefined;

    before(() => {
      sandbox
        .stub(Container, "has")
        .withArgs(sandbox.match("DriverThatCapitalize"))
        .returns(true)
        .withArgs(sandbox.match("DriverThatLowercase"))
        .returns(true)
        .withArgs(sandbox.match("DriverThatUsesEnvField"))
        .returns(true);
      sandbox
        .stub(Container, "get")
        .withArgs(sandbox.match("DriverThatCapitalize"))
        .returns(new DriverThatCapitalize())
        .withArgs(sandbox.match("DriverThatLowercase"))
        .returns(new DriverThatLowercase())
        .withArgs(sandbox.match("DriverThatUsesEnvField"))
        .returns(new DriverThatUsesEnvField());
    });

    after(() => {
      mockedDriverContext.progressBar = undefined;
      sandbox.restore();
      if (restoreFn) {
        restoreFn();
      }
    });

    restoreFn = mockedEnv({
      AZURE_OPENAI_DEPLOYMENT_NAME: "gpt-4o",
    });

    it("should return unresolved placeholders", async () => {
      const driverDefs: DriverDefinition[] = [];
      driverDefs.push({
        uses: "DriverThatCapitalize",
        with: { INPUT_A: "hello ${{ SOME_ENV_VAR }} ${{AAA}} ${{BBB}}" },
      });
      driverDefs.push({
        uses: "DriverThatLowercase",
        with: { INPUT_A: "${{CCC}} Hello ${{OTHER_ENV_VAR}}" },
      });

      const lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");

      const unresolved = lifecycle.resolvePlaceholders();
      assert(
        unresolved.length === 5 &&
          unresolved.some((x) => x === "SOME_ENV_VAR") &&
          unresolved.some((x) => x === "AAA") &&
          unresolved.some((x) => x === "BBB") &&
          unresolved.some((x) => x === "CCC") &&
          unresolved.some((x) => x === "OTHER_ENV_VAR")
      );

      const { result: execResult, summaries } = await lifecycle.execute(mockedDriverContext);
      assert(
        execResult.isErr() &&
          execResult.error.kind === "PartialSuccess" &&
          execResult.error.reason.kind === "UnresolvedPlaceholders" &&
          execResult.error.reason.unresolvedPlaceHolders.length === 3 &&
          execResult.error.reason.unresolvedPlaceHolders.some((x) => x === "SOME_ENV_VAR") &&
          execResult.error.reason.unresolvedPlaceHolders.some((x) => x === "AAA") &&
          execResult.error.reason.unresolvedPlaceHolders.some((x) => x === "BBB") &&
          execResult.error.reason.failedDriver.uses === "DriverThatCapitalize",
        "execute() should fail at first driver because of unresolved placeholders and stop execution"
      );

      assert(
        summaries.length === 1 &&
          summaries[0].length === 1 &&
          summaries[0][0].includes("Unresolved placeholders")
      );
    });

    it("should return unresolved placeholders with nested argument", async () => {
      const driverDefs: DriverDefinition[] = [];
      driverDefs.push({
        uses: "DriverThatCapitalize",
        with: { INPUT_A: ["hello ${{ SOME_ENV_VAR }} ${{AAA}} ${{BBB}}"] },
      });
      driverDefs.push({
        uses: "DriverThatLowercase",
        with: { INPUT_A: { a: "${{CCC}} Hello ${{OTHER_ENV_VAR}}" } },
      });

      const lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");

      const unresolved = lifecycle.resolvePlaceholders();
      assert(
        unresolved.length === 5 &&
          unresolved.some((x) => x === "SOME_ENV_VAR") &&
          unresolved.some((x) => x === "AAA") &&
          unresolved.some((x) => x === "BBB") &&
          unresolved.some((x) => x === "CCC") &&
          unresolved.some((x) => x === "OTHER_ENV_VAR")
      );

      const { result: execResult, summaries } = await lifecycle.execute(mockedDriverContext);
      // execute() will fail at first driver because of unresolved placeholders and stops
      assert(
        execResult.isErr() &&
          execResult.error.kind === "PartialSuccess" &&
          execResult.error.reason.kind === "UnresolvedPlaceholders" &&
          execResult.error.reason.unresolvedPlaceHolders.length === 3 &&
          execResult.error.reason.unresolvedPlaceHolders.some((x) => x === "SOME_ENV_VAR") &&
          execResult.error.reason.unresolvedPlaceHolders.some((x) => x === "AAA") &&
          execResult.error.reason.unresolvedPlaceHolders.some((x) => x === "BBB") &&
          execResult.error.reason.failedDriver.uses === "DriverThatCapitalize"
      );

      assert(
        summaries.length === 1 &&
          summaries[0].length === 1 &&
          summaries[0][0].includes("Unresolved placeholders")
      );
    });

    it("should escape OpenAI related env keys", async () => {
      const driverDefs: DriverDefinition[] = [];
      driverDefs.push({
        uses: "file/createOrUpdateEnvironmentFile",
        with: {
          target: ".env",
          envs: {
            OPENAI_API_KEY: "${{SECRET_OPENAI_API_KEY}}",
            AZURE_OPENAI_DEPLOYMENT_NAME: "${{AZURE_OPENAI_DEPLOYMENT_NAME}}",
          },
        },
      });

      const lifecycle = new Lifecycle("deploy", driverDefs, "1.0.0");

      const unresolved = lifecycle.resolvePlaceholders();
      assert.equal(unresolved.length, 0);
    });

    describe("execute()", async () => {
      it("should return unresolved placeholders in env field", async () => {
        const driverDefs: DriverDefinition[] = [];
        driverDefs.push({
          uses: "DriverThatUsesEnvField",
          with: {},
          env: {
            ENV_VAR1: "hello ${{ SOME_ENV_VAR }}",
          },
        });

        const lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");
        const nextStub = sandbox.stub();
        mockedDriverContext.progressBar = {
          next: nextStub,
        } as unknown as IProgressHandler;
        const { result, summaries } = await lifecycle.execute(mockedDriverContext);
        assert(
          result.isErr() &&
            result.error.kind === "PartialSuccess" &&
            result.error.reason.kind === "UnresolvedPlaceholders" &&
            result.error.reason.unresolvedPlaceHolders.some((x) => x === "SOME_ENV_VAR") &&
            result.error.reason.failedDriver.uses === "DriverThatUsesEnvField"
        );

        assert(
          summaries.length === 1 &&
            summaries[0].length === 1 &&
            summaries[0][0].includes("Unresolved placeholders")
        );

        assert.isTrue(nextStub.calledOnce);
      });
    });
  });
});

describe("Summary", () => {
  const sandbox = sinon.createSandbox();
  const restoreFn = mockedEnv({});

  before(() => {
    sandbox
      .stub(Container, "has")
      .withArgs(sandbox.match("DriverAWithSummary"))
      .returns(true)
      .withArgs(sandbox.match("DriverBWithSummary"))
      .returns(true)
      .withArgs(sandbox.match("DriverThatCapitalizeWithSummary"))
      .returns(true)
      .withArgs(sandbox.match("DriverThatReturnsErrorWithSummary"))
      .returns(true);

    sandbox
      .stub(Container, "get")
      .withArgs(sandbox.match("DriverAWithSummary"))
      .returns(new DriverAWithSummary())
      .withArgs(sandbox.match("DriverBWithSummary"))
      .returns(new DriverBWithSummary())
      .withArgs(sandbox.match("DriverThatCapitalizeWithSummary"))
      .returns(new DriverThatCapitalizeWithSummary())
      .withArgs(sandbox.match("DriverThatReturnsErrorWithSummary"))
      .returns(new DriverThatReturnsErrorWithSummary());
  });

  after(() => {
    sandbox.restore();
    if (restoreFn) {
      restoreFn();
    }
  });

  it("should be returned if all drivers' execute() return ok", async () => {
    const driverDefs: DriverDefinition[] = [];
    driverDefs.push({
      uses: "DriverAWithSummary",
      with: {},
    });
    driverDefs.push({
      uses: "DriverBWithSummary",
      with: {},
    });

    const lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");
    const { result, summaries } = await lifecycle.execute(mockedDriverContext);

    assert(
      result.isOk() &&
        result.value.get("OUTPUT_A") === "VALUE_A" &&
        result.value.get("OUTPUT_B") === "VALUE_B"
    );

    assert(
      summaries.length === 2 &&
        summaries[0].length === 1 &&
        summaries[0][0] ===
          `${SummaryConstant.Succeeded} Environment variable OUTPUT_A set in env/.env file` &&
        summaries[1].length === 1 &&
        summaries[1][0] ===
          `${SummaryConstant.Succeeded} Environment variable OUTPUT_B set in env/.env file`
    );
  });

  it("should contain error summary if any driver's execute() returns error", async () => {
    const driverDefs: DriverDefinition[] = [];
    driverDefs.push({
      uses: "DriverAWithSummary",
      with: {},
    });
    driverDefs.push({
      uses: "DriverBWithSummary",
      with: {},
    });
    driverDefs.push({
      uses: "DriverThatReturnsErrorWithSummary",
      with: {},
    });

    const lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");
    const { result, summaries } = await lifecycle.execute(mockedDriverContext);

    assert(
      result.isErr() &&
        result.error.kind === "PartialSuccess" &&
        result.error.reason.kind === "DriverError" &&
        result.error.reason.failedDriver.uses === "DriverThatReturnsErrorWithSummary" &&
        result.error.reason.error.name === "fakeError" &&
        result.error.env.size === 2 &&
        result.error.env.get("OUTPUT_A") === "VALUE_A" &&
        result.error.env.get("OUTPUT_B") === "VALUE_B"
    );

    assert(
      summaries.length === 3 &&
        summaries[0].length === 1 &&
        summaries[0][0] ===
          `${SummaryConstant.Succeeded} Environment variable OUTPUT_A set in env/.env file` &&
        summaries[1].length === 1 &&
        summaries[1][0] ===
          `${SummaryConstant.Succeeded} Environment variable OUTPUT_B set in env/.env file` &&
        summaries[2].length === 1 &&
        summaries[2][0].includes(`${SummaryConstant.Failed} fake message`)
    );
  });

  it("should contain error summary if there are unresolved placeholders", async () => {
    const driverDefs: DriverDefinition[] = [];
    driverDefs.push({
      uses: "DriverAWithSummary",
      with: {},
    });
    driverDefs.push({
      uses: "DriverBWithSummary",
      with: {
        BBB: "${{ AAA }} ${{ CCC }}",
      },
    });
    driverDefs.push({
      uses: "DriverThatReturnsErrorWithSummary",
      with: {},
    });

    const lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");
    const { result, summaries } = await lifecycle.execute(mockedDriverContext);

    assert(
      result.isErr() &&
        result.error.kind === "PartialSuccess" &&
        result.error.reason.kind === "UnresolvedPlaceholders" &&
        result.error.reason.failedDriver.uses === "DriverBWithSummary" &&
        result.error.env.size === 1 &&
        result.error.env.get("OUTPUT_A") === "VALUE_A" &&
        result.error.reason.unresolvedPlaceHolders.some((x) => x === "AAA")
    );

    assert(
      summaries.length === 2 &&
        summaries[0].length === 1 &&
        summaries[0][0] ===
          `${SummaryConstant.Succeeded} Environment variable OUTPUT_A set in env/.env file` &&
        summaries[1].length === 1 &&
        summaries[1][0] === `${SummaryConstant.Failed} Unresolved placeholders: AAA,CCC`,
      `Summary should only contain 2 items, because of execution stops at DriverBWithSummary`
    );
  });
});

describe("writeToEnvironmentFile", () => {
  const sandbox = sinon.createSandbox();
  const restoreFn = mockedEnv({});

  before(() => {
    sandbox
      .stub(Container, "has")
      .withArgs(sandbox.match("DriverThatUsesWriteToEnvironmentFileField"))
      .returns(true);

    sandbox
      .stub(Container, "get")
      .withArgs(sandbox.match("DriverThatUsesWriteToEnvironmentFileField"))
      .returns(new DriverThatUsesWriteToEnvironmentFileField());
  });

  after(() => {
    sandbox.restore();
    if (restoreFn) {
      restoreFn();
    }
  });

  it("should work", async () => {
    const driverDefs: DriverDefinition[] = [];
    driverDefs.push({
      uses: "DriverThatUsesWriteToEnvironmentFileField",
      with: {},
      writeToEnvironmentFile: {
        key1: "AAA",
        key2: "BBB",
      },
    });

    const lifecycle = new Lifecycle("configureApp", driverDefs, "1.0.0");
    const { result } = await lifecycle.execute(mockedDriverContext);

    assert(
      result.isOk() &&
        result.value.size === 2 &&
        result.value.has("AAA") &&
        result.value.get("AAA") === "aaa" &&
        result.value.has("BBB") &&
        result.value.get("BBB") === "bbb"
    );
  });
});
