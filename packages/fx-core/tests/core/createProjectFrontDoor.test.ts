// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CreateProjectResult, FxError, Inputs, Platform, UserError } from "@microsoft/teamsfx-api";
import { UserInteraction } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import { Result, err, ok } from "neverthrow";
import { Answers, BuildTarget, DeclarativeLocator } from "../../src/v4";
import { CreateFrontDoorDeps, createProjectFrontDoor } from "../../src/core/createProjectFrontDoor";
import { FeatureFlags } from "../../src/common/featureFlags";

/**
 * Tests for docs/03-specs/operations/scaffolding/dispatch-create-by-engine.md.
 * One `it` per DCE-* acceptance-criteria row (DCE-09 is an L3 VS Code E2E,
 * documented in the spec, not exercised here). Every effectful seam is an
 * injected stub, so this is an L1 engine-tier test with no I/O (INV-4 / INV-6).
 */

const EMPTY_FLOOR = Buffer.alloc(0);

const V4_TARGET: BuildTarget = {
  templateId: "da/mcp-server",
  engine: "v4",
  answers: { projectType: "copilot-agent-type", daTemplate: "add-action", actionSource: "mcp" },
};
const V3_TARGET: BuildTarget = {
  templateId: "default-bot",
  engine: "v3",
  answers: {
    projectType: "teams-agent-and-app-type",
    teamsApp: "other",
    teamsOtherAppType: "default-bot",
  },
};
const V3_MCP_TWIN_TARGET: BuildTarget = {
  templateId: "declarative-agent-with-action-from-mcp",
  engine: "v3",
  answers: { projectType: "copilot-agent-type", daTemplate: "add-action", actionSource: "mcp" },
};
const SURFACE_ACTION_TARGET: BuildTarget = {
  templateId: "open-github-copilot-chat",
  engine: "surface-action",
  answers: { projectType: "start-with-github-copilot" },
};

function baseInputs(platform: Platform = Platform.VSCode): Inputs {
  return { platform };
}

/** Inputs with `template-name` preset (the CLI non-interactive contract); defaults to the CLI surface. */
function presetInputs(templateId: string, platform: Platform = Platform.CLI): Inputs {
  return { platform, "template-name": templateId };
}

/** A call-recording stub: `fn` runs `impl` and records each call's arguments. */
function recorder<A extends unknown[], R>(
  impl: (...args: A) => R
): {
  fn: (...args: A) => R;
  calls: A[];
} {
  const calls: A[] = [];
  return {
    calls,
    fn: (...args: A): R => {
      calls.push(args);
      return impl(...args);
    },
  };
}

const okResult = (projectPath: string): Promise<Result<CreateProjectResult, FxError>> =>
  Promise.resolve(ok({ projectPath }));
const okAnswers = (answers: Answers): Promise<Result<Answers, FxError>> =>
  Promise.resolve(ok(answers));
const okTarget = (target: BuildTarget): Promise<Result<BuildTarget, FxError>> =>
  Promise.resolve(ok(target));
const okFloor = (): Promise<Result<undefined, FxError>> => Promise.resolve(ok(undefined));

/** A typed `runCreateSelector` stub that records its `(floor, ui, surface, deps)` args. */
function selectorRecorder(target: BuildTarget) {
  return recorder(
    (
      _floor: Buffer,
      _ui: UserInteraction,
      _surface: string,
      _deps?: { flagReader?: (name: string) => boolean; interactive?: boolean }
    ): Promise<Result<BuildTarget, FxError>> => okTarget(target)
  );
}

/** A typed `runCreateInputs` stub that records its `(floor, locator, entry, ui, deps)` args. */
function inputsRecorder(answers: Answers) {
  return recorder(
    (
      _floor: Buffer,
      _locator: DeclarativeLocator,
      _entry: Answers,
      _ui: UserInteraction,
      _deps?: { flagReader?: (name: string) => boolean; surface?: string }
    ): Promise<Result<Answers, FxError>> => okAnswers(answers)
  );
}

/** A typed `resolveCreateTargetByTemplateId` stub that records its `(floor, templateId)` args. */
function resolveByTemplateIdRecorder(target: BuildTarget) {
  return recorder(
    (_floor: Buffer, _templateId: string): Result<BuildTarget, FxError> => ok(target)
  );
}

// Handlers that must never run on the path under test — invoking one throws and fails the test.
const failCreateV3 = (_inputs: Inputs): Promise<Result<CreateProjectResult, FxError>> => {
  throw new Error("createV3 must not run on this path");
};
const failScaffoldV4 = (
  _inputs: Inputs,
  _target: BuildTarget,
  _answers: Answers
): Promise<Result<CreateProjectResult, FxError>> => {
  throw new Error("scaffoldV4 must not run on this path");
};
const failCollectFloor = (
  _inputs: Inputs,
  _ui: UserInteraction
): Promise<Result<undefined, FxError>> => {
  throw new Error("collectCreateFloor must not run on this path");
};
const failPreFill = (_inputs: Inputs, _target: BuildTarget): void => {
  throw new Error("applyV3PreFill must not run on this path");
};
const failRunInputs = (): Promise<Result<Answers, FxError>> => {
  throw new Error("runInputs must not run on this path");
};
const failRunSelector = (): Promise<Result<BuildTarget, FxError>> => {
  throw new Error("runSelector must not run on this path");
};
const failResolveByTemplateId = (): Result<BuildTarget, FxError> => {
  throw new Error("resolveByTemplateId must not run on this path");
};

// A do-nothing host UI; the flag-on path only hands it to the (stubbed) walks
// (the test-only cast keeps the src no-`as` rule out of scope here).
const stubUI = {} as unknown as UserInteraction;

/** Front-door deps with every seam defaulting to fail-if-called; override per test. */
function deps(overrides: Partial<CreateFrontDoorDeps>): CreateFrontDoorDeps {
  return {
    createV3: failCreateV3,
    scaffoldV4: failScaffoldV4,
    collectCreateFloor: failCollectFloor,
    applyV3PreFill: failPreFill,
    flagReader: () => true,
    readFloorBytes: () => EMPTY_FLOOR,
    ui: stubUI,
    runSelector: failRunSelector,
    resolveByTemplateId: failResolveByTemplateId,
    runInputs: failRunInputs,
    ...overrides,
  };
}

describe("createProjectFrontDoor (dispatch-create-by-engine)", () => {
  it("DCE-01: flag off delegates to createV3 and never walks the selector", async () => {
    const createV3 = recorder((_inputs: Inputs) => okResult("/v3"));
    let selectorWalked = false;

    const res = await createProjectFrontDoor(
      baseInputs(),
      deps({
        createV3: createV3.fn,
        flagReader: () => false,
        runSelector: () => {
          selectorWalked = true;
          return okTarget(V4_TARGET);
        },
      })
    );

    assert.isTrue(res.isOk());
    assert.equal(createV3.calls.length, 1);
    assert.isFalse(selectorWalked, "the selector is never walked on the flag-off path");
  });

  it("DCE-02: engine v4 runs Q2 via runInputs then scaffoldV4, never createV3", async () => {
    const q2: Answers = {
      mcpServerType: "remote",
      mcpServerUrl: "https://api/mcp",
      authType: "none",
    };
    const createV3 = recorder((_inputs: Inputs) => okResult("/v3"));
    const runInputs = inputsRecorder(q2);
    const scaffoldV4 = recorder((_i: Inputs, _t: BuildTarget, _a: Answers) => okResult("/v4"));
    const runSelector = selectorRecorder(V4_TARGET);

    const res = await createProjectFrontDoor(
      baseInputs(),
      deps({
        createV3: createV3.fn,
        scaffoldV4: scaffoldV4.fn,
        runSelector: runSelector.fn,
        runInputs: runInputs.fn,
        collectCreateFloor: okFloor,
      })
    );

    assert.isTrue(res.isOk());
    assert.equal(runInputs.calls.length, 1);
    assert.equal(scaffoldV4.calls.length, 1);
    assert.equal(createV3.calls.length, 0);
    assert.equal(runSelector.calls[0][2], "vscode"); // host platform → selector surface
    assert.equal(runInputs.calls[0][4]?.surface, "vscode"); // host platform → inputs surface (gates csharp)
    assert.deepEqual(scaffoldV4.calls[0][1], V4_TARGET);
    assert.deepEqual(scaffoldV4.calls[0][2], q2);
  });

  it("DCE-03: the Q2 answers reach scaffoldV4 under the create locator", async () => {
    const q2: Answers = {
      mcpServerType: "remote",
      mcpServerUrl: "https://api/mcp",
      authType: "none",
    };
    const runInputs = inputsRecorder(q2);
    const scaffoldV4 = recorder((_i: Inputs, _t: BuildTarget, _a: Answers) => okResult("/v4"));

    const res = await createProjectFrontDoor(
      baseInputs(),
      deps({
        scaffoldV4: scaffoldV4.fn,
        runSelector: () => okTarget(V4_TARGET),
        runInputs: runInputs.fn,
        collectCreateFloor: okFloor,
      })
    );

    assert.isTrue(res.isOk());
    assert.deepEqual(runInputs.calls[0][1], { kind: "create", templateId: "da/mcp-server" });
    assert.deepEqual(runInputs.calls[0][2], V4_TARGET.answers);
    assert.deepEqual(scaffoldV4.calls[0][2], q2);
  });

  it("DCE-04: engine v3 pre-fills from the Q1 picks then delegates to createV3", async () => {
    const prefill = recorder((_i: Inputs, _t: BuildTarget) => undefined);
    const createV3 = recorder((_inputs: Inputs) => okResult("/v3"));
    const inputs = baseInputs();

    const res = await createProjectFrontDoor(
      inputs,
      deps({
        createV3: createV3.fn,
        applyV3PreFill: prefill.fn,
        runSelector: () => okTarget(V3_TARGET),
      })
    );

    assert.isTrue(res.isOk());
    assert.equal(prefill.calls.length, 1);
    assert.deepEqual(prefill.calls[0][1], V3_TARGET);
    assert.equal(createV3.calls.length, 1);
    // pre-fill mutates the same inputs object that is then handed to createV3.
    assert.strictEqual(prefill.calls[0][0], inputs);
    assert.strictEqual(createV3.calls[0][0], inputs);
  });

  it("DCE-05: DT-off DA+MCP resolves the v3 twin, then pre-fills and delegates to createV3", async () => {
    const prefill = recorder((_i: Inputs, _t: BuildTarget) => undefined);
    const createV3 = recorder((_inputs: Inputs) => okResult("/v3"));

    const res = await createProjectFrontDoor(
      baseInputs(),
      deps({
        createV3: createV3.fn,
        applyV3PreFill: prefill.fn,
        runSelector: () => okTarget(V3_MCP_TWIN_TARGET),
      })
    );

    assert.isTrue(res.isOk());
    assert.deepEqual(prefill.calls[0][1], V3_MCP_TWIN_TARGET);
    assert.equal(createV3.calls.length, 1);
  });

  it("DCE-06: a surface-action returns shouldInvokeTeamsAgent and scaffolds nothing", async () => {
    // createV3 / scaffoldV4 / applyV3PreFill / runInputs all default to fail-if-called.
    const res = await createProjectFrontDoor(
      baseInputs(),
      deps({ runSelector: () => okTarget(SURFACE_ACTION_TARGET) })
    );

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.equal(res.value.projectPath, "");
      assert.isTrue(res.value.shouldInvokeTeamsAgent);
    }
  });

  it("DCE-07: the DA+MCP v4 route bypasses createV3 (the v3 generator path)", async () => {
    // Under V4 the front door scaffolds DA+MCP directly through the v4 engine, so
    // createV3 (the v3 generator carrier) is never reached.
    const createV3 = recorder((_inputs: Inputs) => okResult("/v3"));

    const res = await createProjectFrontDoor(
      baseInputs(),
      deps({
        createV3: createV3.fn,
        scaffoldV4: (_i, _t, _a) => okResult("/v4"),
        runSelector: () => okTarget(V4_TARGET),
        runInputs: () => okAnswers({}),
        collectCreateFloor: okFloor,
      })
    );

    assert.isTrue(res.isOk());
    assert.equal(createV3.calls.length, 0);
  });

  it("DCE-08: a Q1 cancellation surfaces as the Result error and runs no hand-off", async () => {
    // every hand-off defaults to fail-if-called.
    const cancel = new UserError({ source: "Test", name: "UserCancelError", message: "cancel" });

    const res = await createProjectFrontDoor(
      baseInputs(),
      deps({ runSelector: () => Promise.resolve(err(cancel)) })
    );

    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.name, "UserCancelError");
    }
  });

  it("DCE-10: a preset template-name resolving to v3 skips Q1 + pre-fill, then delegates to createV3", async () => {
    // runSelector + applyV3PreFill default to fail-if-called: the preset path walks
    // neither (the v3 traverse short-circuits on template-name downstream).
    const resolveByTemplateId = resolveByTemplateIdRecorder({
      templateId: "default-bot",
      engine: "v3",
      answers: {},
    });
    const createV3 = recorder((_inputs: Inputs) => okResult("/v3"));
    const inputs = presetInputs("default-bot");

    const res = await createProjectFrontDoor(
      inputs,
      deps({ createV3: createV3.fn, resolveByTemplateId: resolveByTemplateId.fn })
    );

    assert.isTrue(res.isOk());
    assert.equal(resolveByTemplateId.calls.length, 1);
    assert.equal(resolveByTemplateId.calls[0][1], "default-bot"); // the preset id is forwarded
    assert.equal(createV3.calls.length, 1);
    assert.strictEqual(createV3.calls[0][0], inputs);
  });

  it("DCE-11: a preset template-name resolving to v4 runs Q2 then scaffoldV4, never createV3", async () => {
    const q2: Answers = {
      mcpServerType: "remote",
      mcpServerUrl: "https://api/mcp",
      authType: "none",
    };
    const resolveByTemplateId = resolveByTemplateIdRecorder({
      templateId: "da/mcp-server",
      engine: "v4",
      answers: {},
    });
    const runInputs = inputsRecorder(q2);
    const scaffoldV4 = recorder((_i: Inputs, _t: BuildTarget, _a: Answers) => okResult("/v4"));

    const res = await createProjectFrontDoor(
      presetInputs("da/mcp-server"),
      deps({
        scaffoldV4: scaffoldV4.fn,
        resolveByTemplateId: resolveByTemplateId.fn,
        runInputs: runInputs.fn,
        collectCreateFloor: okFloor,
      })
    );

    assert.isTrue(res.isOk());
    assert.equal(resolveByTemplateId.calls.length, 1);
    assert.equal(runInputs.calls.length, 1);
    assert.equal(scaffoldV4.calls.length, 1);
    assert.deepEqual(runInputs.calls[0][1], { kind: "create", templateId: "da/mcp-server" });
  });

  it("DCE-13: a non-interactive walk (no preset template-name) threads interactive:false into runSelector", async () => {
    const runSelector = selectorRecorder(SURFACE_ACTION_TARGET);
    const inputs: Inputs = { platform: Platform.CLI, nonInteractive: true };

    const res = await createProjectFrontDoor(inputs, deps({ runSelector: runSelector.fn }));

    assert.isTrue(res.isOk());
    assert.equal(runSelector.calls.length, 1);
    assert.equal(runSelector.calls[0][3]?.interactive, false);
  });

  it("DCE-14: engine v4 collects the create floor after Q2 and before scaffoldV4", async () => {
    const order: string[] = [];
    const runInputs = recorder(
      (
        _floor: Buffer,
        _locator: DeclarativeLocator,
        _entry: Answers,
        _ui: UserInteraction,
        _deps?: { flagReader?: (name: string) => boolean }
      ): Promise<Result<Answers, FxError>> => {
        order.push("q2");
        return okAnswers({});
      }
    );
    const collectFloor = recorder((_i: Inputs, _ui: UserInteraction) => {
      order.push("floor");
      return okFloor();
    });
    const scaffoldV4 = recorder((_i: Inputs, _t: BuildTarget, _a: Answers) => {
      order.push("scaffold");
      return okResult("/v4");
    });

    const res = await createProjectFrontDoor(
      baseInputs(),
      deps({
        scaffoldV4: scaffoldV4.fn,
        runSelector: () => okTarget(V4_TARGET),
        runInputs: runInputs.fn,
        collectCreateFloor: collectFloor.fn,
      })
    );

    assert.isTrue(res.isOk());
    assert.equal(collectFloor.calls.length, 1);
    assert.deepEqual(order, ["q2", "floor", "scaffold"]); // floor sits between Q2 and the scaffold
    // the floor mutates the same inputs bag scaffoldV4 then scaffolds from.
    assert.strictEqual(collectFloor.calls[0][0], scaffoldV4.calls[0][0]);
  });

  it("DCE-15: a create-floor cancellation propagates and does not scaffold", async () => {
    const cancel = new UserError({ source: "Test", name: "UserCancelError", message: "cancel" });
    const scaffoldV4 = recorder((_i: Inputs, _t: BuildTarget, _a: Answers) => okResult("/v4"));

    const res = await createProjectFrontDoor(
      baseInputs(),
      deps({
        scaffoldV4: scaffoldV4.fn,
        runSelector: () => okTarget(V4_TARGET),
        runInputs: () => okAnswers({}),
        collectCreateFloor: () => Promise.resolve(err(cancel)),
      })
    );

    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.name, "UserCancelError");
    }
    assert.equal(scaffoldV4.calls.length, 0);
  });

  it("maps the host platform onto the selector surface (cli, vs)", async () => {
    const cli = selectorRecorder(SURFACE_ACTION_TARGET);
    await createProjectFrontDoor(baseInputs(Platform.CLI), deps({ runSelector: cli.fn }));
    assert.equal(cli.calls[0][2], "cli");

    const vs = selectorRecorder(SURFACE_ACTION_TARGET);
    await createProjectFrontDoor(baseInputs(Platform.VS), deps({ runSelector: vs.fn }));
    assert.equal(vs.calls[0][2], "vs");
  });

  it("defaults the flag reader to featureFlagManager (V4 off ⇒ pass-through)", async () => {
    const saved = process.env[FeatureFlags.V4Enabled.name];
    delete process.env[FeatureFlags.V4Enabled.name];
    try {
      const createV3 = recorder((_inputs: Inputs) => okResult("/v3"));

      const res = await createProjectFrontDoor(
        baseInputs(),
        // no flagReader override → the real featureFlagManager default reads V4 off.
        deps({ createV3: createV3.fn, flagReader: undefined, runSelector: failRunSelector })
      );

      assert.isTrue(res.isOk());
      assert.equal(createV3.calls.length, 1);
    } finally {
      if (saved === undefined) {
        delete process.env[FeatureFlags.V4Enabled.name];
      } else {
        process.env[FeatureFlags.V4Enabled.name] = saved;
      }
    }
  });

  it("propagates a Q2 error and does not scaffold", async () => {
    const q2Failed = new UserError({ source: "Test", name: "Q2Failed", message: "bad inputs" });
    const scaffoldV4 = recorder((_i: Inputs, _t: BuildTarget, _a: Answers) => okResult("/v4"));

    const res = await createProjectFrontDoor(
      baseInputs(),
      deps({
        scaffoldV4: scaffoldV4.fn,
        runSelector: () => okTarget(V4_TARGET),
        runInputs: () => Promise.resolve(err(q2Failed)),
      })
    );

    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.name, "Q2Failed");
    }
    assert.equal(scaffoldV4.calls.length, 0);
  });

  it("fails loudly on an unsupported create engine (v3-core-method)", async () => {
    const target: BuildTarget = { templateId: "some-core-method", engine: "v3-core-method" };

    const res = await createProjectFrontDoor(
      baseInputs(),
      deps({ runSelector: () => okTarget(target) })
    );

    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.name, "UnsupportedCreateEngine");
    }
  });

  it("fails loudly on an unhandled surface action", async () => {
    const target: BuildTarget = { templateId: "some-other-action", engine: "surface-action" };

    const res = await createProjectFrontDoor(
      baseInputs(),
      deps({ runSelector: () => okTarget(target) })
    );

    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.name, "UnsupportedCreateAction");
    }
  });
});
