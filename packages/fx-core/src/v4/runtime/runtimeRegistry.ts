// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError } from "@microsoft/teamsfx-api";
import { Result } from "neverthrow";
import { RenderVars } from "../model/dataModel";
import { ExpressionRuntimePort, Scope, evaluateExpression } from "../expression/evaluateExpression";
import {
  ManifestWrapper,
  Orchestration,
  PipelineRuntimePort,
  RegisteredStep,
} from "../pipeline/runScaffoldPipeline";
import { renderMustache } from "./renderMustache";
import {
  STEP_INJECT_YML_ACTION,
  STEP_PERSIST_CREDENTIAL_ENV,
  mcpAuthInjectYmlAction,
  mcpAuthPersistCredentialEnv,
} from "./steps/mcpAuth";
import { STEP_MATERIALIZE_LOCAL_SERVERS, mcpLocalMaterializeServers } from "./steps/mcpLocal";
import { STEP_GENERATE_OPENAPI_PLUGIN_FILES, openApiGeneratePluginFiles } from "./steps/openApi";

/**
 * The engine's wiring registry — the closed orchestration / step whitelist and
 * the `PipelineRuntimePort` factory shared by every `ScaffoldRuntime` (the
 * in-memory test seam and the on-disk production seam). Hoisted here so the
 * registry is one source of truth: adding a pipeline or a post-render step is a
 * single edit both runtimes pick up (ADR-0017 closed whitelist).
 *
 * Only the file sink (`write` / `read`) differs between runtimes; everything
 * else — the registries, the real Mustache renderer, the real `when` evaluator —
 * is identical, so it lives here and the sink is injected.
 *
 * v4-owned (INV-7): imports no v3 symbol.
 */

/** The orchestration names the engine knows (ADR-0017 closed whitelist). */
export const KNOWN_PIPELINES = new Set(["default", "openapi", "typespec", "officeAddin", "spfx"]);

/** The post-render step whitelist (ADR-0017 decision 2). */
export const STEP_REGISTRY = new Map<string, RegisteredStep>([
  [STEP_INJECT_YML_ACTION, mcpAuthInjectYmlAction],
  [STEP_PERSIST_CREDENTIAL_ENV, mcpAuthPersistCredentialEnv],
  [STEP_MATERIALIZE_LOCAL_SERVERS, mcpLocalMaterializeServers],
  [STEP_GENERATE_OPENAPI_PLUGIN_FILES, openApiGeneratePluginFiles],
]);

/**
 * A no-op manifest wrapper: the `da/mcp-server` create pipeline has no
 * manifest-mutating step (that is the `add`/modify flow), so this face is never
 * called here; it exists only to satisfy the port (INV-3).
 */
export const NOOP_MANIFEST_WRAPPER: ManifestWrapper = {
  addAction: () => undefined,
};

/**
 * The one outside-world face that differs between runtimes: a synchronous file
 * sink the render phase writes through and a read-modify-write step reads back
 * (run-scaffold-pipeline AC-21). The in-memory seam backs it with a `Map`; the
 * on-disk seam backs it with `fs`.
 */
export interface FileSink {
  /** Persist `data` at `path` (a target-relative, forward-slash path). */
  write(path: string, data: Buffer): void;
  /** Read back a previously written file, or `undefined` when absent (EAFP). */
  read(path: string): Buffer | undefined;
}

/**
 * Project the render context down to its scalar subset for the `when` evaluator.
 * A `multiSelect` answer carried through `{from}` is a `string[]` (collect-inputs
 * INV-7); arrays are off the scalar expression grammar, so they are dropped here
 * rather than handed to `evaluateExpression`, which stays scalar-valued.
 */
function scalarScope(renderVars: RenderVars): Scope {
  const scope: Scope = {};
  for (const [key, value] of Object.entries(renderVars)) {
    if (Array.isArray(value)) {
      continue;
    }
    scope[key] = value;
  }
  return scope;
}

/**
 * Build the two-phase pipeline executor's port over an injected file sink. The
 * registries, the real renderer, and the real `when` evaluator are shared; only
 * the sink varies, so a whole-template scaffold is byte-identical across the
 * in-memory and on-disk runtimes.
 */
export function buildPipelinePort(
  exprPort: ExpressionRuntimePort,
  sink: FileSink
): PipelineRuntimePort {
  return {
    pipelineRegistry: (name: string): Orchestration | undefined =>
      KNOWN_PIPELINES.has(name) ? { name } : undefined,
    stepRegistry: (name: string): RegisteredStep | undefined => STEP_REGISTRY.get(name),
    evalWhen: (expr: string, renderVars: RenderVars): Result<boolean, FxError> =>
      evaluateExpression({ expr }, scalarScope(renderVars), exprPort).map(
        (value) => value === true
      ),
    render: (mustache: string, renderVars: RenderVars): Result<string, FxError> =>
      renderMustache(mustache, renderVars),
    manifestWrapper: (): ManifestWrapper => NOOP_MANIFEST_WRAPPER,
    write: (path: string, data: Buffer): void => sink.write(path, data),
    read: (path: string): Buffer | undefined => sink.read(path),
  };
}
