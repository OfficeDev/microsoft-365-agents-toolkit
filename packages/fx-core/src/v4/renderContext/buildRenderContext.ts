// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import {
  EvalValue,
  ExpressionRuntimePort,
  NULL_VALUE,
  Scope,
  evaluateExpression,
} from "../expression/evaluateExpression";
import { Answers, CallerFloor, RenderVars } from "../model/dataModel";

/**
 * The v4 render-context build — answers → render variables.
 *
 * Spec: docs/03-specs/operations/scaffolding/build-render-context.md
 * Decision: docs/02-architecture/adr/ADR-0016-declarative-template-format.md
 *           (decision 3 closed-DSL replaceMap, 4 caller-floor no-shadow, 9 typed context)
 *
 * This is part of the v4 world; it imports no v3 symbol. It realizes the closed
 * `replaceMap` DSL: turning one template's authored entries — together with the
 * resolved answers and the caller-injected identifier floor — into the render
 * variable map that `content/**` and step `with` interpolate against.
 *
 * It is **one** behavior (answers → render vars), distinct from `collect-inputs`
 * (questions → answers) even though both call the shared evaluator. It is pure
 * and deterministic (spec INV-5): no fs / http / clock.
 */

const SOURCE = "Scaffold";

/**
 * One closed-DSL `replaceMap` entry (ADR-0016 decision 3): each is exactly one
 * of `{const}` (literal), `{from}` (verbatim copy of an answer), `{when, value}`
 * (emit `value` iff the `when` guard evaluates true), or `{expr}` (a value-context
 * evaluator call). There is no free-form string-assembly path (spec INV-1).
 */
export type ReplaceMapEntry =
  | { var: string; const: string }
  | { var: string; from: string }
  | { var: string; when: string; value: string }
  | { var: string; expr: string };

/**
 * `SystemError` name: a `replaceMap[].var` names a caller-injected identifier.
 * Reaching this at runtime is an engine bug the build gate should have caught
 * (spec INV-2 / decision 4); a template derives a **new** `PascalCase` var
 * instead of transforming a floor id in place.
 */
export const RCTX_SHADOWS_CALLER_FLOOR = "RenderVarShadowsCallerFloor";

/**
 * Build one template's render variable map from its `replaceMap`, resolved
 * `answers`, and the caller-injected floor.
 *
 * @param replaceMap   the closed-DSL entry list (`descriptor.replaceMap`)
 * @param answers      the resolved answer object (raw answers ∪ `derived.*`)
 * @param callerFloor  the closed `camelCase` caller-injected identifier set
 * @param port         the pure function whitelist + feature-flag reader (spec INV-4)
 * @param declaredKeys the `optionsSchema.properties` id set; each is seeded as
 *                     the `null` marker so a `{from}` / `{expr}` over a declared
 *                     but unanswered option renders `""` instead of raising the
 *                     undeclared-identifier `SystemError` (spec INV-3). Defaults
 *                     to `[]` (no extra seeding) for callers without a schema.
 * @returns `ok(RenderVars)` = raw answers ∪ `replaceMap`-derived ∪ `derived.*`,
 *          or a `SystemError` for a caller-floor shadow, an undeclared
 *          identifier, a non-whitelisted call, or a rejected expression.
 */
export function buildRenderContext(
  replaceMap: ReplaceMapEntry[],
  answers: Answers,
  callerFloor: CallerFloor,
  port: ExpressionRuntimePort,
  declaredKeys: string[] = []
): Result<RenderVars, FxError> {
  // Identifier domain for {from} / {expr}: optionsSchema.properties ∪ derived.* ∪
  // callerFloor (spec INV-3). Computed once, so a later entry cannot reference an
  // earlier entry's derived var — the result stays order-independent (INV-5).
  // Every declared property is seeded as NULL_VALUE first, so a {from} / {expr}
  // over a conditionally-skipped option (e.g. `mcpServerUrl` on the local branch)
  // resolves to "" rather than the undeclared SystemError — the same presence
  // model collect-inputs uses. The caller floor and scalar answers then overlay,
  // so an answered id wins over its NULL_VALUE seed. A multiSelect answer
  // (string[]) is off the scalar grammar, so it is not placed in the evaluator
  // scope (collect-inputs INV-7); it is still copied via {from}. An id outside
  // declaredKeys ∪ callerFloor ∪ answers stays absent, so a typo is still a
  // SystemError.
  const scope: Scope = {};
  for (const key of declaredKeys) {
    scope[key] = NULL_VALUE;
  }
  Object.assign(scope, callerFloor);
  for (const [key, value] of Object.entries(answers)) {
    if (!Array.isArray(value)) {
      scope[key] = value;
    }
  }
  // Seed = raw answers ∪ derived.* (the floor is overlaid downstream, not here).
  const renderVars: RenderVars = { ...answers };

  for (const entry of replaceMap) {
    // INV-2: a replaceMap var never shadows a caller-injected identifier.
    if (entry.var in callerFloor) {
      return err(
        new SystemError({
          source: SOURCE,
          name: RCTX_SHADOWS_CALLER_FLOOR,
          message: `replaceMap var '${entry.var}' shadows the caller-injected identifier '${entry.var}'; derive a new PascalCase var instead.`,
        })
      );
    }

    if ("const" in entry) {
      renderVars[entry.var] = entry.const;
    } else if ("from" in entry) {
      // RCTX-11 / INV-7: a {from} of a multiSelect answer copies the string[]
      // verbatim; a scalar answer goes through the evaluator, which also checks
      // the identifier is in the closed domain (INV-3).
      const direct = answers[entry.from];
      if (Array.isArray(direct)) {
        renderVars[entry.var] = direct;
      } else {
        const r = evaluateExpression({ from: entry.from }, scope, port);
        if (r.isErr()) {
          return err(r.error);
        }
        renderVars[entry.var] = asString(r.value);
      }
    } else if ("when" in entry) {
      const r = evaluateExpression({ expr: entry.when }, scope, port);
      if (r.isErr()) {
        return err(r.error);
      }
      // A false guard emits nothing — an `optional` var (spec §3.4 / RCTX-04).
      if (r.value === true) {
        renderVars[entry.var] = entry.value;
      }
    } else {
      const r = evaluateExpression({ expr: entry.expr }, scope, port);
      if (r.isErr()) {
        return err(r.error);
      }
      renderVars[entry.var] = asString(r.value);
    }
  }

  return ok(renderVars);
}

/** A value-context result is a string; a stray boolean renders as its literal. */
function asString(v: EvalValue): string {
  return typeof v === "string" ? v : String(v);
}
