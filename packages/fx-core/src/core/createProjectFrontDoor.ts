// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  CreateProjectResult,
  FxError,
  Inputs,
  Platform,
  SystemError,
  UserInteraction,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import path from "path";
import { Result, err, ok } from "neverthrow";
import {
  Answers,
  BuildTarget,
  DeclarativeLocator,
  bundledFloorDir,
  resolveCreateTargetByTemplateId,
  runCreateInputs,
  runCreateSelector,
} from "../v4";
import { FeatureFlags, featureFlagManager } from "../common/featureFlags";
import { TOOLS } from "../common/globalVars";
import { QuestionNames } from "../question/questionNames";

/**
 * Operation `dispatch-create-by-engine` — the create front door.
 *
 * Spec: docs/03-specs/operations/scaffolding/dispatch-create-by-engine.md
 * Decision: docs/02-architecture/adr/ADR-0014-dispatcher-buildtarget-resolution.md
 *
 * The single entry the create surfaces call in place of `FxCore.createProject`.
 * Behind `TEAMSFX_V4_ENABLED`, the v4 create selector is the live Q1 front door
 * and the resolved `BuildTarget` is dispatched by its `engine` (INV-3):
 *   - `v4`             → run the template's own Q2 (`runCreateInputs`) over the
 *                        same floor, collect the create floor (`folder`/`app-name`,
 *                        the step the v3 `QuestionMW` owns), then `scaffoldV4` the
 *                        authored package;
 *   - `v3`             → translate the Q1 picks onto the v3 `QuestionNames.*`
 *                        (`applyV3PreFill`, INV-5) and hand off to `createV3`,
 *                        whose `QuestionMW` then skips Q1 and asks only Q2;
 *   - `surface-action` → return the action's surface signal (no scaffold).
 * Flag off is a pure pass-through to the unmodified `createV3` (INV-1): the
 * selector is never walked, so v3 behavior is byte-identical.
 *
 * This orchestrator is a seam outside the v4 world (INV-4): it touches the v3
 * `Inputs` and calls `createV3`, which v4 may not. Every effectful step is an
 * injected dependency, so the dispatch is verifiable without I/O; the floor read
 * is injectable too (INV-6).
 */

const SOURCE = "Scaffold";

/** The only shipped create `surface-action`: open GitHub Copilot Chat (the v3 `startWithGithubCopilot` shape). */
const OPEN_GITHUB_COPILOT_CHAT = "open-github-copilot-chat";

/**
 * The create front door's injected seams. `createV3` is required — it is both the
 * flag-off pass-through and the engine=v3 hand-off, and injecting it (rather than
 * importing `FxCore`) keeps this seam free of an import cycle. `scaffoldV4`,
 * `collectCreateFloor`, and `applyV3PreFill` are the flag-on hand-offs the
 * composition root (`FxCore`) supplies. The remaining members default to the real
 * wiring, so a production caller passes only the four handlers.
 */
export interface CreateFrontDoorDeps {
  /** The flag-off pass-through + the engine=v3 hand-off: the unmodified `FxCore.createProject`. */
  createV3: (inputs: Inputs) => Promise<Result<CreateProjectResult, FxError>>;
  /** The engine=v4 hand-off: build the scaffold context + run the authored declarative package. */
  scaffoldV4: (
    inputs: Inputs,
    target: BuildTarget,
    answers: Answers
  ) => Promise<Result<CreateProjectResult, FxError>>;
  /**
   * The engine=v4 create-floor collection: ask `folder` + `app-name`. The v4 path
   * carries no `QuestionMW`, so the front door collects the floor itself — the same
   * step `createProject`'s `QuestionMW` runs (last) for v3. Interactive surfaces are
   * prompted; a non-interactive surface (a CLI preset `template-name` / `-f` / `-n`)
   * is skipped exactly as v3 is.
   */
  collectCreateFloor: (
    inputs: Inputs,
    ui: UserInteraction
  ) => Promise<Result<undefined, FxError>>;
  /** The engine=v3 adapter: translate the Q1 dimension picks onto the v3 `QuestionNames.*` (INV-5). */
  applyV3PreFill: (inputs: Inputs, target: BuildTarget) => void;
  /** The feature-flag reader (default: `featureFlagManager`-backed, so VS Code-settings flags apply). */
  flagReader?: (name: string) => boolean;
  /** The bundled-floor channel-zip reader (default: the shipped `templates.zip`; injectable for tests, INV-6). */
  readFloorBytes?: () => Buffer;
  /** The host surface (default: `TOOLS.ui`). */
  ui?: UserInteraction;
  /** The Q1 selector walk (default: the real `runCreateSelector`). */
  runSelector?: typeof runCreateSelector;
  /** Resolve a target directly from a preset `template-name`, bypassing Q1 (default: the real `resolveCreateTargetByTemplateId`). */
  resolveByTemplateId?: typeof resolveCreateTargetByTemplateId;
  /** The Q2 inputs walk (default: the real `runCreateInputs`). */
  runInputs?: typeof runCreateInputs;
}

/** The default `featureFlagManager`-backed reader (a flag is on per its env var / VS Code setting). */
function defaultFlagReader(name: string): boolean {
  return featureFlagManager.getBooleanValue({ name, defaultValue: "false" });
}

/** Read the shipped bundled-floor channel zip (the default `readFloorBytes`). */
function readBundledFloorBytes(): Buffer {
  return fs.readFileSync(path.join(bundledFloorDir(), "templates.zip"));
}

/** Map the host `Platform` onto the selector's `surface` axis (drives option `condition`s). */
function surfaceOf(platform: Platform | undefined): string {
  switch (platform) {
    case Platform.CLI:
    case Platform.CLI_HELP:
      return "cli";
    case Platform.VS:
      return "vs";
    default:
      return "vscode";
  }
}

/** Dispatch a resolved `surface-action` target onto its surface signal (no scaffold). */
function dispatchSurfaceAction(target: BuildTarget): Result<CreateProjectResult, FxError> {
  if (target.templateId === OPEN_GITHUB_COPILOT_CHAT) {
    return ok({ projectPath: "", shouldInvokeTeamsAgent: true });
  }
  return err(
    new SystemError({
      source: SOURCE,
      name: "UnsupportedCreateAction",
      message: `The create front door does not handle the '${target.templateId}' surface action.`,
    })
  );
}

/**
 * Run the create front door for `inputs`, dispatching the resolved engine.
 *
 * @param inputs the v3 create inputs (carries `platform`; mutated in place by the
 *               engine=v3 pre-fill before the hand-off to `createV3`)
 * @param deps   the injected seams (see `CreateFrontDoorDeps`)
 * @returns the created project (drop-in for `FxCore.createProject`), or a
 *          `UserError` / `SystemError` (a surface cancellation or a route break)
 */
export async function createProjectFrontDoor(
  inputs: Inputs,
  deps: CreateFrontDoorDeps
): Promise<Result<CreateProjectResult, FxError>> {
  const flagReader = deps.flagReader ?? defaultFlagReader;

  // INV-1: flag off ⇒ a pure pass-through to the unmodified v3 createProject.
  if (!flagReader(FeatureFlags.V4Enabled.name)) {
    return deps.createV3(inputs);
  }

  const ui = deps.ui ?? TOOLS.ui;
  const surface = surfaceOf(inputs.platform);
  const floorBytes = (deps.readFloorBytes ?? readBundledFloorBytes)();
  const interactive = !inputs.nonInteractive;

  // A surface that already resolved the leaf template — the CLI in non-interactive
  // mode presets `template-name` from its `-c` capability — pins the BuildTarget by
  // id: the Q1 selector is a *router*, so re-walking it would re-prompt, or (non-
  // interactive) fail on a missing dimension, for a target already chosen. Resolve
  // the engine from the template's route and reuse the v3 `traverse` short-circuit
  // on `template-name` for the v3 hand-off. Otherwise walk Q1 (INV-2), threading
  // `interactive` so a non-interactive surface never silently prompts.
  const presetTemplateId = inputs[QuestionNames.TemplateName];
  let target: Result<BuildTarget, FxError>;
  if (presetTemplateId) {
    const resolveByTemplateId = deps.resolveByTemplateId ?? resolveCreateTargetByTemplateId;
    target = resolveByTemplateId(floorBytes, presetTemplateId);
  } else {
    const runSelector = deps.runSelector ?? runCreateSelector;
    target = await runSelector(floorBytes, ui, surface, { flagReader, interactive });
  }
  if (target.isErr()) {
    return err(target.error);
  }

  // INV-3: exactly one resolved BuildTarget, dispatched by its engine.
  switch (target.value.engine) {
    case "surface-action":
      return dispatchSurfaceAction(target.value);
    case "v3":
      // INV-5: translate the Q1 picks onto the v3 inputs (a translation, not a
      // re-ask), then hand off to createProject — its QuestionMW skips Q1. A preset
      // `template-name` carries no Q1 picks and already short-circuits the v3
      // traverse, so the pre-fill is skipped.
      if (!presetTemplateId) {
        deps.applyV3PreFill(inputs, target.value);
      }
      return deps.createV3(inputs);
    case "v4": {
      const runInputs = deps.runInputs ?? runCreateInputs;
      const locator: DeclarativeLocator = { kind: "create", templateId: target.value.templateId };
      // Q2: the template's own inputs, over the same floor.
      const answers = await runInputs(floorBytes, locator, target.value.answers ?? {}, ui, {
        flagReader,
        surface,
      });
      if (answers.isErr()) {
        return err(answers.error);
      }
      // The v4 path carries no QuestionMW (createProject's, which asks the v3 tree's
      // folder/app-name last), so collect the create floor here — interactive
      // surfaces are prompted; a non-interactive surface (CLI preset / -f / -n) is
      // skipped (the traverse short-circuit + questionVisitor preset-skip mirror v3).
      const floorRes = await deps.collectCreateFloor(inputs, ui);
      if (floorRes.isErr()) {
        return err(floorRes.error);
      }
      return deps.scaffoldV4(inputs, target.value, answers.value);
    }
    case "v3-core-method":
      // The shipped create selector carries no v3-core-method route; fail loudly
      // (no silent fallback) rather than mis-route through createProject.
      return err(
        new SystemError({
          source: SOURCE,
          name: "UnsupportedCreateEngine",
          message: `The create front door does not dispatch the '${target.value.engine}' engine.`,
        })
      );
  }
}
