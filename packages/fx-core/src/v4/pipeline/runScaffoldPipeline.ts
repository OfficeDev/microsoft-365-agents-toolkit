// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError, UserError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { RenderVars, TemplateFileEntry } from "../model/dataModel";

/**
 * The v4 scaffold-pipeline executor — one template package's `pipeline.json`
 * materialized against a resolved render context.
 *
 * Spec: docs/03-specs/operations/scaffolding/run-scaffold-pipeline.md
 * Decision: docs/02-architecture/adr/ADR-0017-named-pipeline-step-whitelist.md
 *           (the step `when` grammar is shared with ADR-0016 §4.3; the reverse
 *           `minEngineVersion` gate that makes every referenced step present is
 *           ADR-0015 AC-16..AC-18, upstream of this operation)
 *
 * Two phases, in order (proposal §13.1):
 *   1. a fixed, engine-owned **render phase** — each `.tpl` entry has its path
 *      and body rendered through the one Mustache surface (the `.tpl` suffix
 *      stripped on write), every other file is copied **verbatim**, and only
 *      files **absent** from the target are written (exists → skip + warning);
 *      then
 *   2. the ordered **post-render steps** from `pipeline.steps`, each a
 *      domain-typed whitelist entry that mutates *existing* files, routing any
 *      manifest mutation through a `packages/manifest` wrapper (INV-3).
 *
 * v4-owned (INV-7): no v3 generator / `onDidSelection` reuse. It consumes the
 * one shared evaluator for each step's `when` and adds no grammar (INV-4). It
 * owns the render phase but neither decides *which* package to run nor builds
 * the render-var map (those are `resolve-template-source` /
 * `build-render-context`, upstream).
 */

const SOURCE = "Scaffold";

/**
 * The one engine-builtin guard the executor evaluates **directly**, before the
 * render phase — a create-contract precondition with no params and no manifest
 * mutation, so a violation writes nothing (AC-14). Every other step is
 * dispatched through `stepRegistry`.
 */
const STEP_REQUIRE_EMPTY_TARGET = "require-empty-target";

const TPL_SUFFIX = ".tpl";

const SKIP_WARNING = "exists, not overwritten; delete or rename it to rebuild";

/**
 * A `with` value — a Mustache string, a JSON literal (both appear in the
 * fixtures), or a `string[]` when the value is exactly one `{{Token}}` whose
 * render-var is a `multiSelect` selection (resolved structurally, AC-22).
 */
export type ParamValue = string | boolean | string[];

/** A step's author-supplied `with` block, or its resolved form. */
export type StepParams = Record<string, ParamValue>;

/**
 * One `pipeline.steps` entry (ADR-0017 decisions 1/5). `with` values are
 * Mustache over the render-var space; `when` is the shared closed grammar.
 * `produces` is the forward-looking cross-step form, loader-rejected here
 * (INV-8 / AC-16).
 */
export interface PipelineStep {
  step: string;
  comment?: string;
  when?: string;
  with?: StepParams;
  produces?: string[];
}

/** A parsed, schema-valid `pipeline.json` (ADR-0017). */
export interface Pipeline {
  pipeline: string;
  comment?: string;
  steps: PipelineStep[];
}

/** The output directory + the file set it already contains (the create entry, ADR-0014). */
export interface TargetDir {
  path: string;
  existing: string[];
}

/** A render-phase path skipped because it already existed (carries its warning, AC-05). */
export interface SkippedFile {
  path: string;
  warning: string;
}

/** The result of a scaffold run (spec Outputs). */
export interface ScaffoldOutcome {
  written: string[];
  skipped: SkippedFile[];
  stepsRun: string[];
  stepsSkipped: string[];
}

/** A resolved named orchestration — only its presence (known vs unknown) is read here. */
export interface Orchestration {
  name: string;
}

/**
 * A `packages/manifest` wrapper, threaded to a manifest-mutating step (INV-3).
 * The operation forwards it verbatim; only the minimal slice the conformance
 * step (`da-action/register-plugin-manifest`) calls is modeled — the full
 * wrapper API lives in `packages/manifest`, not at this boundary.
 */
export interface ManifestWrapper {
  addAction(action: Record<string, string>): void;
}

/** The capabilities the executor hands each registered step's `apply`. */
export interface StepContext {
  write(path: string, data: Buffer): void;
  manifestWrapper(kind: string): ManifestWrapper;
  /**
   * Read the current bytes at a target path — `undefined` if absent. The
   * read-modify-write face a non-manifest step (e.g. `mcp-auth/inject-yml-action`
   * rewriting the render-phase `m365agents.yml`) needs; manifest mutation still
   * routes through `manifestWrapper` (INV-3).
   */
  read(path: string): Buffer | undefined;
}

/**
 * An engine-registered step (ADR-0017 decision 2): it validates its **resolved**
 * `with` against its own `paramsSchema` and applies its mutation. The whitelist
 * of these is engine-owned; this operation only **dispatches** within it.
 */
export interface RegisteredStep {
  validateParams(resolved: StepParams): string | undefined;
  apply(
    resolved: StepParams,
    ctx: StepContext
  ): Result<void, FxError> | Promise<Result<void, FxError>>;
}

/**
 * The narrow port this operation declares (interface-segregation over the full
 * `ScaffoldRuntime`). `render` is the **single** Mustache surface used for both
 * `.tpl` paths/bodies and `with` values — the spec's `fs.render` and
 * `renderValue` faces are deliberately one function so INV-4 ("two surfaces, no
 * third") holds by construction. Existence is read from `TargetDir.existing` (a
 * pure snapshot), so there is no separate stateful `fs.exists` oracle (AC-17).
 */
export interface PipelineRuntimePort {
  pipelineRegistry(pipelineName: string): Orchestration | undefined;
  stepRegistry(stepName: string): RegisteredStep | undefined;
  evalWhen(expr: string, renderVars: RenderVars): Result<boolean, FxError>;
  render(mustache: string, renderVars: RenderVars): Result<string, FxError>;
  manifestWrapper(kind: string): ManifestWrapper;
  write(path: string, data: Buffer): void;
  /** The read-modify-write face: current bytes at a path, `undefined` if absent (AC-21). */
  read(path: string): Buffer | undefined;
}

/** `SystemError` names — an engine-side break a build gate should have caught (spec Outputs). */
export const PIPELINE_UNKNOWN_PIPELINE = "PipelineUnknownPipeline";
export const PIPELINE_UNKNOWN_STEP = "PipelineUnknownStep";
export const PIPELINE_PARAMS_VIOLATION = "PipelineParamsViolation";
export const PIPELINE_CROSS_STEP_REFERENCE = "PipelineCrossStepReference";

/** `UserError` name — the create-contract violation (a non-empty target, AC-14). */
export const REQUIRE_EMPTY_TARGET = "RequireEmptyTarget";

function systemError(name: string, message: string): SystemError {
  return new SystemError({ source: SOURCE, name, message });
}

/** Evaluate a step's `when` (absent guard ⇒ active). */
function whenActive(
  step: PipelineStep,
  renderVars: RenderVars,
  port: PipelineRuntimePort
): Result<boolean, FxError> {
  return step.when ? port.evalWhen(step.when, renderVars) : ok(true);
}

/** Resolve a step's `with`: Mustache over render vars for strings, literals verbatim (AC-09). */
function resolveParams(
  params: StepParams,
  renderVars: RenderVars,
  port: PipelineRuntimePort
): Result<StepParams, FxError> {
  const resolved: StepParams = {};
  for (const key of Object.keys(params)) {
    const value = params[key];
    if (typeof value === "string") {
      // a `with` value that is exactly one `{{Token}}` whose render-var is a
      // string[] (a multiSelect selection carried via {from}, build-render-context
      // RCTX-11) resolves structurally — the typed list is handed to the step
      // verbatim rather than flattened by the scalar Mustache surface (AC-22).
      const soleToken = value.match(/^\{\{(\w+)\}\}$/);
      if (soleToken) {
        const listValue = renderVars[soleToken[1]];
        if (Array.isArray(listValue)) {
          resolved[key] = listValue;
          continue;
        }
      }
      const rendered = port.render(value, renderVars);
      if (rendered.isErr()) {
        return err(rendered.error);
      }
      resolved[key] = rendered.value;
    } else {
      // a JSON literal (e.g. `includeCredentialRefs: true`) passes through unrendered
      resolved[key] = value;
    }
  }
  return ok(resolved);
}

/**
 * Execute one template package's pipeline against a resolved render context.
 *
 * @param pipeline   the parsed, schema-valid `pipeline.json`
 * @param content    the template's renderable file entries (unrendered, `.tpl` intact)
 * @param renderVars the resolved render-var map (raw answers ∪ `replaceMap` ∪ `derived.*`)
 * @param targetDir  the output path + its current file set (the create entry, ADR-0014)
 * @param port       the narrow `PipelineRuntimePort` (an in-memory fake in tests)
 * @returns `ok(ScaffoldOutcome)`, or a `UserError` (create-contract violation) /
 *          `SystemError` (an engine-side invariant break that reached runtime).
 */
export async function runScaffoldPipeline(
  pipeline: Pipeline,
  content: TemplateFileEntry[],
  renderVars: RenderVars,
  targetDir: TargetDir,
  port: PipelineRuntimePort
): Promise<Result<ScaffoldOutcome, FxError>> {
  // INV-2 — the pipeline name resolves within the engine whitelist; an unknown
  // one reaching execution is an engine break, never a silent no-op (AC-02).
  if (!port.pipelineRegistry(pipeline.pipeline)) {
    return err(
      systemError(
        PIPELINE_UNKNOWN_PIPELINE,
        `Unknown pipeline '${pipeline.pipeline}'. The name is a closed engine whitelist; reaching execution with an unknown one is an engine inconsistency.`
      )
    );
  }

  // INV-8 — cross-step data flow (`produces` / `<stepId>.<field>`) is forward-
  // looking; render vars are frozen before any step runs, so it is loader-
  // rejected until the §13.1 modify-flow step library lands (AC-16). The
  // `<stepId>.<field>` form degenerates to a missing-token error at render time
  // (render vars are flat), so the explicit `produces` declaration is the signal.
  for (const step of pipeline.steps) {
    if (step.produces !== undefined) {
      return err(
        systemError(
          PIPELINE_CROSS_STEP_REFERENCE,
          `Step '${step.step}' declares a forward-looking cross-step reference ('produces'); render vars are frozen before steps run, so this form is not yet supported.`
        )
      );
    }
  }

  // The create-contract guard runs BEFORE the render phase so a violation writes
  // nothing (AC-14). It is recorded into `stepsRun` / `stepsSkipped` later, in
  // its declared position, by the main step loop.
  for (const step of pipeline.steps) {
    if (step.step !== STEP_REQUIRE_EMPTY_TARGET) {
      continue;
    }
    const active = whenActive(step, renderVars, port);
    if (active.isErr()) {
      return err(active.error);
    }
    if (active.value && targetDir.existing.length > 0) {
      return err(
        new UserError({
          source: SOURCE,
          name: REQUIRE_EMPTY_TARGET,
          message: `The target folder '${targetDir.path}' is not empty. Scaffolding a new project requires an empty folder; remove the existing files or choose another folder, then try again.`,
        })
      );
    }
  }

  // Phase 1 — render. New files only; existence keys off the frozen snapshot
  // (INV-5). A `.tpl` entry has BOTH its path and body rendered (the dynamic
  // `ai-plugin-{{MCPNamespace}}.json.tpl` filename, §13.1); every other file is
  // copied verbatim (AC-19).
  const written: string[] = [];
  const skipped: SkippedFile[] = [];
  for (const entry of content) {
    if (entry.path.endsWith(TPL_SUFFIX)) {
      const renderedPath = port.render(entry.path.slice(0, -TPL_SUFFIX.length), renderVars);
      if (renderedPath.isErr()) {
        return err(renderedPath.error); // an unproducible path token is an engine break
      }
      const writePath = renderedPath.value;
      if (targetDir.existing.includes(writePath)) {
        skipped.push({ path: writePath, warning: SKIP_WARNING }); // AC-05
        continue;
      }
      const renderedBody = port.render(entry.data.toString("utf8"), renderVars); // AC-18
      if (renderedBody.isErr()) {
        return err(renderedBody.error); // AC-20 — an unproducible body token is an engine break
      }
      port.write(writePath, Buffer.from(renderedBody.value, "utf8"));
      written.push(writePath);
    } else {
      if (targetDir.existing.includes(entry.path)) {
        skipped.push({ path: entry.path, warning: SKIP_WARNING });
        continue;
      }
      port.write(entry.path, entry.data); // AC-19 — verbatim, no substitution, no suffix change
      written.push(entry.path);
    }
  }

  // Phase 2 — post-render steps, in declared order.
  const ctx: StepContext = {
    write: (path, data) => port.write(path, data),
    manifestWrapper: (kind) => port.manifestWrapper(kind),
    read: (path) => port.read(path),
  };
  const stepsRun: string[] = [];
  const stepsSkipped: string[] = [];
  for (const step of pipeline.steps) {
    const active = whenActive(step, renderVars, port);
    if (active.isErr()) {
      return err(active.error); // AC-11 — `when` references an absent identifier
    }

    if (step.step === STEP_REQUIRE_EMPTY_TARGET) {
      // already enforced before render; only record its declared-order status
      (active.value ? stepsRun : stepsSkipped).push(step.step);
      continue;
    }

    const registered = port.stepRegistry(step.step);
    if (!registered) {
      return err(
        systemError(
          PIPELINE_UNKNOWN_STEP,
          `Unknown step '${step.step}'. The reverse minEngineVersion gate (ADR-0015) guarantees every referenced step is present; an absent one is an engine inconsistency.`
        )
      ); // AC-03
    }

    if (!active.value) {
      stepsSkipped.push(step.step); // AC-07
      continue;
    }

    const resolved = resolveParams(step.with ?? {}, renderVars, port);
    if (resolved.isErr()) {
      return err(resolved.error); // AC-11 — `with` references an absent identifier
    }

    const violation = registered.validateParams(resolved.value);
    if (violation !== undefined) {
      return err(
        systemError(
          PIPELINE_PARAMS_VIOLATION,
          `Step '${step.step}' resolved parameters violate its schema: ${violation}. The build-time typed-context check (ADR-0016) should have caught this.`
        )
      ); // AC-11
    }

    const applied = await registered.apply(resolved.value, ctx);
    if (applied.isErr()) {
      return err(applied.error);
    }
    stepsRun.push(step.step); // AC-06 / AC-08
  }

  return ok({ written, skipped, stepsRun, stepsSkipped });
}
