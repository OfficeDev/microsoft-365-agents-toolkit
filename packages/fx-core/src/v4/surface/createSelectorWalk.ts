// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError, UserError, UserInteraction } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import {
  PresentationOption,
  PresentationQuestion,
  SelectorPresentation,
} from "../buildTarget/parseSelector";
import {
  BuildTarget,
  PromptResult,
  RouteQuestion,
  RouteResolverPort,
  resolveBuildTarget,
} from "../buildTarget/resolveBuildTarget";
import { openCreateSelector, openCreateSelectorPresentation } from "../distribution/createSelector";
import { openDeclarativePackage } from "../distribution/declarativePackage";
import { ExpressionRuntimePort, Scope, evaluateExpression } from "../expression/evaluateExpression";

/**
 * Operation `walk-create-selector` — the live Q1 prompt face.
 *
 * Spec: docs/03-specs/operations/scaffolding/walk-create-selector.md
 *
 * Run a create kind's Q1 (`selector.json`) over the host `UserInteraction`,
 * resolving the engine decision — a `BuildTarget { templateId, engine,
 * answers }`. This is the surface wiring of `resolveBuildTarget`'s
 * prefill-aware `walk` source: it loads the shipped `selector.json` from the
 * floor, builds a `RouteResolverPort` whose `prompt` renders each un-pre-filled
 * routing question over `UserInteraction`, and walks Q1 into the dispatched
 * target — the front door (principle 1: the `templateId` comes from the
 * selector, never the v3 tree). A pre-filled dimension is used as-is and a
 * non-interactive run never prompts (resolveBuildTarget AC-03a/03b).
 *
 * v4-owned (INV-1): `UserInteraction` is `@microsoft/teamsfx-api` (upstream of
 * both worlds), and the feature-flag reader is injected — v4 imports no
 * `featureFlagManager`. The floor read is injectable (INV-5), so the operation
 * is CI-testable from an in-memory floor with no built artifact.
 */

const SOURCE = "Scaffold";

/**
 * The create-selector options (all defaulted): the feature-flag reader behind
 * `featureFlag('…')` + route `when`, plus the prefill-aware walk inputs
 * (walk-create-selector §"Inputs").
 */
export interface CreateSelectorDeps {
  /** The feature-flag reader (default: env-backed); v4 imports no `featureFlagManager`. */
  flagReader?: (name: string) => boolean;
  /** Q1-dimension answers known up front; a pre-filled dimension is used as-is, never prompted (default `{}`). */
  prefilled?: Record<string, string>;
  /** Whether an un-pre-filled gated dimension may be prompted; `false` ⇒ a missing required dimension is a `UserError` (default `true`). */
  interactive?: boolean;
}

/** The default env-backed feature-flag reader (a flag is on iff its env var is exactly `"true"`). */
function envFlagReader(name: string): boolean {
  return process.env[name] === "true";
}

/** Convert a thrown prompt failure back to an `FxError` for the `Result` boundary. */
function toFxError(e: unknown): FxError {
  if (e instanceof UserError || e instanceof SystemError) {
    return e;
  }
  const message = e instanceof Error ? e.message : String(e);
  return new SystemError({ source: SOURCE, name: "CreateSelectorWalkFailed", message });
}

/**
 * Build the live interactive `RouteResolverPort` over the floor + the host
 * `UserInteraction`. The `prompt` face renders each routing question's
 * presentation, filtering its `staticOptions` by each option's environment
 * `condition` (the shared evaluator over a `{ surface }` scope + the injected
 * `flagReader`) and returning the chosen `id`; a surface cancellation throws its
 * `FxError`, which `runCreateSelector` converts back to a `Result` error.
 * v4 membership is descriptor-derived from the floor (INV-5); the `v3Registry` /
 * `v3CoreMethodRegistry` are unused by the create selector (the table carries no
 * v3 / core-method route today). Language is not resolved here — it is the
 * downstream `collect-inputs` Q0 question.
 */
function buildPort(
  floorBytes: Buffer,
  presentation: SelectorPresentation,
  ui: UserInteraction,
  surface: string,
  flagReader: (name: string) => boolean
): RouteResolverPort {
  const exprPort: ExpressionRuntimePort = { functions: () => undefined, flags: flagReader };
  const byName = new Map<string, PresentationQuestion>(
    presentation.questions.map((q) => [q.name, q])
  );

  async function prompt(question: RouteQuestion, step: number): Promise<PromptResult> {
    const pq = byName.get(question.name);
    if (pq === undefined) {
      throw new SystemError({
        source: SOURCE,
        name: "MissingSelectorPresentation",
        message: `The selector has no presentation for question '${question.name}'.`,
      });
    }
    const scope: Scope = { surface };
    const visible: PresentationOption[] = [];
    for (const option of pq.staticOptions) {
      if (option.condition !== undefined) {
        const gate = evaluateExpression(option.condition, scope, exprPort);
        if (gate.isErr()) {
          throw gate.error;
        }
        if (gate.value !== true) {
          continue;
        }
      }
      visible.push(option);
    }
    const selected = await ui.selectOption({
      name: pq.name,
      title: pq.title ?? pq.name,
      placeholder: pq.placeholder,
      step,
      options: visible.map((option) => ({
        id: option.id,
        label: option.label,
        detail: option.detail,
        groupName: option.groupName,
      })),
      returnObject: false,
    });
    if (selected.isErr()) {
      throw selected.error;
    }
    if (selected.value.type === "back") {
      return { kind: "back" };
    }
    const result = selected.value.result;
    if (typeof result === "string") {
      return { kind: "value", value: result };
    }
    return { kind: "value", value: result === undefined ? "" : result.id };
  }

  return {
    prompt,
    featureFlag: flagReader,
    v4Registry(templateId: string): boolean {
      return openDeclarativePackage(floorBytes, { kind: "create", templateId }).isOk();
    },
    v3Registry(_templateId: string): boolean {
      return false;
    },
    v3CoreMethodRegistry(_coreMethod: string): boolean {
      return false;
    },
  };
}

/**
 * Run the create Q1 over `ui`, resolving a dispatched `BuildTarget`. The walk is
 * prefill-aware: a dimension in `deps.prefilled` is used as-is (no prompt), an
 * un-pre-filled one is prompted when `deps.interactive` (default), else a
 * missing-dimension `UserError` (resolveBuildTarget AC-03a/03b).
 *
 * @param floorBytes the bundled-floor channel zip (injectable for CI tests)
 * @param ui         the host surface (`@microsoft/teamsfx-api`)
 * @param surface    the surface id (`"vscode"` / `"cli"` / `"vs"`) for option `condition`s
 * @param deps       injected `flagReader` (default env-backed) + `prefilled` (default `{}`) + `interactive` (default `true`)
 * @returns `ok(BuildTarget { …, answers })`, or a `UserError` / `SystemError`
 *          (a selector / route break, a missing dimension, or a surface cancellation).
 */
export async function runCreateSelector(
  floorBytes: Buffer,
  ui: UserInteraction,
  surface: string,
  deps: CreateSelectorDeps = {}
): Promise<Result<BuildTarget, FxError>> {
  const flagReader = deps.flagReader ?? envFlagReader;
  const prefilled = deps.prefilled ?? {};
  const interactive = deps.interactive ?? true;
  const spec = openCreateSelector(floorBytes);
  if (spec.isErr()) {
    return err(spec.error);
  }
  const presentation = openCreateSelectorPresentation(floorBytes);
  if (presentation.isErr()) {
    return err(presentation.error);
  }
  const port = buildPort(floorBytes, presentation.value, ui, surface, flagReader);
  try {
    return await resolveBuildTarget(spec.value, prefilled, interactive, port);
  } catch (e) {
    return err(toFxError(e));
  }
}

/**
 * Resolve a `BuildTarget` directly from an already-chosen `templateId`, without
 * walking Q1.
 *
 * A surface that already resolved the leaf template — the CLI in non-interactive
 * mode presets `template-name` from its `-c` capability — has no Q1 dimensions
 * left to ask: the selector walk is a *router*, so re-walking it would re-prompt
 * (interactive) or fail on a missing dimension (non-interactive) for a target the
 * caller already pinned. This finds the route whose `templateId` matches and
 * returns its `engine`, so the front door can hand off (a v3 target then reuses
 * the v3 `traverse` short-circuit on `template-name`).
 *
 * A `templateId` with no selector route is **not** an error: it resolves to
 * `engine:"v3"`, the coexistence default — any v3 `TemplateNames` value the CLI
 * presets scaffolds through `createV3` exactly as under the flag-off path,
 * whether or not the selector happens to expose it. `answers` is empty (no Q1
 * picks were made); the v3 hand-off is driven by `template-name`, not by picks.
 *
 * @param floorBytes the bundled-floor channel zip (injectable for CI tests)
 * @param templateId the already-chosen leaf template id (a v3 `TemplateNames`
 *                   value or a v4 `<kind>/<id>`)
 * @returns `ok(BuildTarget { templateId, engine, answers:{} })`, or a
 *          `SystemError` (a malformed / unreadable selector in the floor).
 */
export function resolveCreateTargetByTemplateId(
  floorBytes: Buffer,
  templateId: string
): Result<BuildTarget, FxError> {
  const spec = openCreateSelector(floorBytes);
  if (spec.isErr()) {
    return err(spec.error);
  }
  const route = spec.value.routes.find((r) => r.templateId === templateId);
  return ok({ templateId, engine: route?.engine ?? "v3", answers: {} });
}
