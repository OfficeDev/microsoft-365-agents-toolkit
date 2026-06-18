// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ExpressionRuntimePort } from "../expression/evaluateExpression";
import { PipelineRuntimePort } from "../pipeline/runScaffoldPipeline";
import { createExpressionPort } from "./whitelist";
import { FileSink, buildPipelinePort } from "./runtimeRegistry";

/**
 * The pure `InMemoryRuntime` (ADR-0018): a `PipelineRuntimePort` whose render
 * surface is the real Mustache renderer, whose `when` is the real evaluator,
 * and whose read/write share one `Map` so a step's read-modify-write observes
 * the render phase's bytes (run-scaffold-pipeline AC-21). No fs / http / clock,
 * which is exactly what makes a whole-template scaffold fuzzable at the T3 tier.
 *
 * The only face that differs from the on-disk production runtime is the file
 * sink (a `Map` here, `fs` there); everything else is the shared
 * `buildPipelinePort` wiring.
 *
 * Spec: docs/03-specs/scenarios/da/create-mcp-server.md (the SCN-CREATE-MCP-* run host)
 * Decision: docs/02-architecture/adr/ADR-0018-scaffold-runtime-test-pyramid.md
 *
 * v4-owned (INV-7): imports no v3 symbol.
 */

/** A composed in-memory runtime: the pipeline port, its backing file map, and the shared expr port. */
export interface InMemoryRuntime {
  /** The pipeline executor's port, backed by the shared file map. */
  port: PipelineRuntimePort;
  /** The shared render + write file map (`path → bytes`); the read face reads back from it. */
  files: Map<string, Buffer>;
  /** The pure expression port (whitelist + flags), reused for `{expr}` render-var derivation. */
  exprPort: ExpressionRuntimePort;
}

/**
 * Build an in-memory runtime. The optional `flagReader` backs `featureFlag('…')`
 * (env-backed by default; injected in tests).
 */
export function createInMemoryRuntime(flagReader?: (name: string) => boolean): InMemoryRuntime {
  const files = new Map<string, Buffer>();
  const exprPort = createExpressionPort(flagReader);
  const sink: FileSink = {
    write: (path: string, data: Buffer): void => {
      files.set(path, data);
    },
    read: (path: string): Buffer | undefined => files.get(path),
  };
  const port = buildPipelinePort(exprPort, sink);
  return { port, files, exprPort };
}
