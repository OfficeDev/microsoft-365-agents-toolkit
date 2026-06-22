// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, Inputs, Platform, SystemError, UserInteraction } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import { Result, err, ok } from "neverthrow";
import { Answers, BuildTarget, DeclarativeLocator } from "../../src/v4";
import { ModifyFrontDoorDeps, modifyProjectFrontDoor } from "../../src/core/modifyProjectFrontDoor";

const EMPTY_FLOOR = Buffer.alloc(0);

const V4_TARGET: BuildTarget = {
  templateId: "add-mcp-server",
  engine: "v4",
  answers: { addCapability: "add-action", actionSource: "mcp" },
};

const V3_CORE_TARGET: BuildTarget = {
  templateId: "addPlugin",
  engine: "v3-core-method",
  answers: { addCapability: "add-action", actionSource: "mcp" },
};

function recorder<A extends unknown[], R>(
  impl: (...args: A) => R
): { fn: (...args: A) => R; calls: A[] } {
  const calls: A[] = [];
  return {
    calls,
    fn: (...args: A): R => {
      calls.push(args);
      return impl(...args);
    },
  };
}

const stubUI = {} as unknown as UserInteraction;
const okTarget = (target: BuildTarget): Promise<Result<BuildTarget, FxError>> =>
  Promise.resolve(ok(target));
const okAnswers = (answers: Answers): Promise<Result<Answers, FxError>> =>
  Promise.resolve(ok(answers));
const okUndefined = (): Promise<Result<undefined, FxError>> => Promise.resolve(ok(undefined));

function selectorRecorder(target: BuildTarget) {
  return recorder(
    (
      _floor: Buffer,
      _ui: UserInteraction,
      _surface: string,
      _deps?: {
        flagReader?: (name: string) => boolean;
        interactive?: boolean;
        prefilled?: Record<string, string>;
      }
    ): Promise<Result<BuildTarget, FxError>> => {
      void _floor;
      void _ui;
      void _surface;
      void _deps;
      return okTarget(target);
    }
  );
}

function inputsRecorder(answers: Answers) {
  return recorder(
    (
      _floor: Buffer,
      _locator: DeclarativeLocator,
      _entry: Answers,
      _ui: UserInteraction,
      _deps?: { flagReader?: (name: string) => boolean; surface?: string }
    ): Promise<Result<Answers, FxError>> => {
      void _floor;
      void _locator;
      void _entry;
      void _ui;
      void _deps;
      return okAnswers(answers);
    }
  );
}

const failRunSelector = (): Promise<Result<BuildTarget, FxError>> => {
  throw new Error("runSelector must not run on this path");
};
const failRunInputs = (): Promise<Result<Answers, FxError>> => {
  throw new Error("runInputs must not run on this path");
};
const failScaffoldV4 = (): Promise<Result<undefined, FxError>> => {
  throw new Error("scaffoldV4 must not run on this path");
};
const failCoreMethod = (): Promise<Result<undefined, FxError>> => {
  throw new Error("callCoreMethod must not run on this path");
};

function deps(overrides: Partial<ModifyFrontDoorDeps>): ModifyFrontDoorDeps {
  return {
    readFloorBytes: () => EMPTY_FLOOR,
    flagReader: () => true,
    ui: stubUI,
    runSelector: failRunSelector,
    runInputs: failRunInputs,
    scaffoldV4: failScaffoldV4,
    callCoreMethod: failCoreMethod,
    ...overrides,
  };
}

describe("modifyProjectFrontDoor", () => {
  it("MDE-01: engine v4 runs modify Q2 via runInputs then scaffoldV4", async () => {
    const runSelector = selectorRecorder(V4_TARGET);
    const q2: Answers = {
      mcpServerUrl: "https://example.com/mcp",
      teamsManifestPath: "manifest.json",
      authType: "none",
    };
    const runInputs = inputsRecorder(q2);
    const scaffoldV4 = recorder((_inputs: Inputs, _target: BuildTarget, _answers: Answers) => {
      void _inputs;
      void _target;
      void _answers;
      return okUndefined();
    });
    const coreMethod = recorder((_inputs: Inputs, _target: BuildTarget) => {
      void _inputs;
      void _target;
      return okUndefined();
    });

    const res = await modifyProjectFrontDoor(
      { platform: Platform.VSCode },
      { addCapability: "add-action", actionSource: "mcp" },
      { mcpServerUrl: "https://example.com/mcp" },
      deps({
        runSelector: runSelector.fn,
        runInputs: runInputs.fn,
        scaffoldV4: scaffoldV4.fn,
        callCoreMethod: coreMethod.fn,
      })
    );

    assert.isTrue(res.isOk());
    assert.equal(runSelector.calls.length, 1);
    assert.deepEqual(runSelector.calls[0][3]?.prefilled, {
      addCapability: "add-action",
      actionSource: "mcp",
    });
    assert.equal(runSelector.calls[0][2], "vscode");
    assert.equal(runInputs.calls.length, 1);
    assert.deepEqual(runInputs.calls[0][1], { kind: "modify", templateId: "add-mcp-server" });
    assert.deepEqual(runInputs.calls[0][2], {
      addCapability: "add-action",
      actionSource: "mcp",
      mcpServerUrl: "https://example.com/mcp",
    });
    assert.equal(runInputs.calls[0][4]?.surface, "vscode");
    assert.equal(scaffoldV4.calls.length, 1);
    assert.deepEqual(scaffoldV4.calls[0][1], V4_TARGET);
    assert.deepEqual(scaffoldV4.calls[0][2], q2);
    assert.equal(coreMethod.calls.length, 0);
  });

  it("MDE-02: engine v3-core-method dispatches to the core-method handler without v4 Q2", async () => {
    const runSelector = selectorRecorder(V3_CORE_TARGET);
    const coreMethod = recorder((_inputs: Inputs, _target: BuildTarget) => {
      void _inputs;
      void _target;
      return okUndefined();
    });

    const res = await modifyProjectFrontDoor(
      { platform: Platform.CLI },
      { addCapability: "add-action", actionSource: "mcp" },
      {},
      deps({
        runSelector: runSelector.fn,
        callCoreMethod: coreMethod.fn,
      })
    );

    assert.isTrue(res.isOk());
    assert.equal(runSelector.calls[0][2], "cli");
    assert.equal(coreMethod.calls.length, 1);
    assert.deepEqual(coreMethod.calls[0][1], V3_CORE_TARGET);
  });

  it("MDE-03: unsupported selector engines fail loudly", async () => {
    const target: BuildTarget = { templateId: "x", engine: "surface-action" };

    const res = await modifyProjectFrontDoor(
      { platform: Platform.VSCode },
      {},
      {},
      deps({ runSelector: () => okTarget(target) })
    );

    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.instanceOf(res.error, SystemError);
      assert.equal(res.error.name, "UnsupportedModifyEngine");
    }
  });

  it("MDE-04: selector errors are returned without dispatching", async () => {
    const selectorError = new SystemError({
      source: "Test",
      name: "SelectorFailed",
      message: "selector failed",
    });

    const res = await modifyProjectFrontDoor(
      { platform: Platform.VSCode },
      {},
      {},
      deps({ runSelector: () => Promise.resolve(err(selectorError)) })
    );

    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.name, "SelectorFailed");
    }
  });
});
