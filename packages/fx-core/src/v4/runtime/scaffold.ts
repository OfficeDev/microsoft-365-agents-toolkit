// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError } from "@microsoft/teamsfx-api";
import { Result, err } from "neverthrow";
import { Answers, CallerFloor, TemplateFileEntry } from "../model/dataModel";
import { ExpressionRuntimePort } from "../expression/evaluateExpression";
import { buildRenderContext } from "../renderContext/buildRenderContext";
import {
  PipelineRuntimePort,
  ScaffoldOutcome,
  TargetDir,
  runScaffoldPipeline,
} from "../pipeline/runScaffoldPipeline";
import { parseDeclaredKeys, parsePipeline, parseReplaceMap } from "./packageParse";
import { COMMON_LANGUAGE, selectLanguageContent } from "./selectLanguageContent";

/**
 * The v4 scaffold composition — one opened template package + its answers →
 * a materialized file set, run against an injected runtime (ADR-0018).
 *
 * It is the thin weld over the already-specced operations: parse the package's
 * `descriptor` / `pipeline` shapes, derive the render-var map
 * (`build-render-context`), then execute the two-phase pipeline
 * (`run-scaffold-pipeline`). It adds no behavior of its own — every fact a
 * caller asserts is owned by a composed operation.
 *
 * The sink is the injected runtime's port (an in-memory `Map` for the T3
 * scenario tier; an on-disk sink for the real surface). The caller inspects the
 * resulting file set through whatever the runtime exposes.
 *
 * Spec: docs/03-specs/scenarios/da/create-mcp-server.md (SCN-CREATE-MCP-01..10)
 *
 * v4-owned (INV-7): imports no v3 symbol. Welding this onto the real
 * distribution chain + an on-disk file sink (so v3 calls it) is the follow-up.
 */

/** The two ports `scaffold` composes — satisfied structurally by `InMemoryRuntime`. */
export interface ScaffoldRuntime {
  /** The pure expression port (whitelist + flags) for `{expr}` render-var derivation. */
  exprPort: ExpressionRuntimePort;
  /** The two-phase pipeline executor's port (render surface + step registry + file sink). */
  port: PipelineRuntimePort;
}

/** One scaffold invocation's inputs (an opened package + its resolved answers + target). */
export interface ScaffoldRequest {
  /** The package's parsed `descriptor.json` (its `replaceMap` is the render-var source). */
  descriptor: unknown;
  /** The package's parsed `pipeline.json`. */
  pipeline: unknown;
  /** The opened `content/**` entries (raw bytes, `.tpl` suffix intact). */
  content: TemplateFileEntry[];
  /** The resolved answer object (`collect-inputs` output). */
  answers: Answers;
  /** The caller-injected identifier floor (`appName`, the `language` axis, …). */
  callerFloor: CallerFloor;
  /** The output directory + the files it already contains (the create-empty contract). */
  targetDir: TargetDir;
}

/**
 * Scaffold one template package against an injected runtime.
 *
 * @returns `ok(ScaffoldOutcome)` (the `written` / `skipped` / `stepsRun` /
 *          `stepsSkipped` record), or the first operation's `FxError` — a
 *          `UserError` for a non-empty target (the create contract), or a
 *          `SystemError` for an engine-shape break.
 */
export async function scaffold(
  request: ScaffoldRequest,
  runtime: ScaffoldRuntime
): Promise<Result<ScaffoldOutcome, FxError>> {
  const replaceMap = parseReplaceMap(request.descriptor);
  if (replaceMap.isErr()) {
    return err(replaceMap.error);
  }
  const pipeline = parsePipeline(request.pipeline);
  if (pipeline.isErr()) {
    return err(pipeline.error);
  }

  // The declared optionsSchema.properties ids seed the render-context identifier
  // domain so a conditionally-skipped option (the local branch leaves
  // `mcpServerUrl` unanswered) renders "" instead of an undeclared SystemError
  // (build-render-context INV-3 / RCTX-12).
  const declaredKeys = parseDeclaredKeys(request.descriptor);
  const baseVars = buildRenderContext(
    replaceMap.value,
    request.answers,
    request.callerFloor,
    runtime.exprPort,
    declaredKeys
  );
  if (baseVars.isErr()) {
    return err(baseVars.error);
  }
  // Overlay the caller floor last (build-render-context deliberately omits it,
  // its no-shadow guard already ran) so the render phase resolves `{{appName}}`.
  const renderVars = { ...baseVars.value, ...request.callerFloor };

  // Narrow a language-partitioned package's `content/{language}/` to the active
  // language before render (ADR-0016 §5); a `["common"]` package is returned
  // flat, unchanged. The language is the caller floor's Q0 answer.
  const language = request.callerFloor.language || COMMON_LANGUAGE;
  const content = selectLanguageContent(request.descriptor, request.content, language);

  return runScaffoldPipeline(
    pipeline.value,
    content,
    renderVars,
    request.targetDir,
    runtime.port
  );
}
