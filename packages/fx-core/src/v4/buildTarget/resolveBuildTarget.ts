// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, UserError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import {
  ExpressionNode,
  ExpressionRuntimePort,
  NULL_VALUE,
  Scope,
  evaluateExpression,
} from "../expression/evaluateExpression";

/**
 * The v4 front stage — resolve a create/modify `selector.json` walk into a
 * `BuildTarget = { templateId, engine, answers }` and dispatch that `templateId`
 * to the world (v4 / v3 / v3-core-method / surface-action) that handles it.
 *
 * Spec: docs/03-specs/operations/scaffolding/resolve-build-target.md
 * Decision: docs/02-architecture/adr/ADR-0014-dispatcher-buildtarget-resolution.md
 *           (Amendment 2 — one `walk` source; the `language` axis moves to
 *           `collect-inputs` Q0, ADR-0016 decision 5)
 *
 * Part of the v4 world; it imports no v3 symbol (spec INV-8) and reaches v3 only
 * through the engine that dispatch hands off to. Route resolution yields a
 * `templateId` only — it never reads `descriptor.languages` (spec INV-1):
 * language is the downstream `collect-inputs` Q0 question, not a routing input.
 * Route `when` and question `condition` are the one shared `evaluate-expression`
 * grammar (spec INV-3 / one-grammar), so they are evaluated through that
 * evaluator, never a second routing table.
 */

const SOURCE = "Scaffold";

/** The four worlds dispatch can hand off to (spec §"Outputs", INV-4). */
export type DispatchEngine = "v4" | "v3" | "v3-core-method" | "surface-action";

/**
 * One parsed `selector.json` route. `when` is the route predicate; `engine`
 * selects the world and carries exactly its own id key (spec INV-4 / invariant
 * 12): `v4` → `templateId`, `v3` → `templateId` (+ `v3Adapter`), `v3-core-method`
 * → `coreMethod`, `surface-action` → `action`.
 */
export interface SelectorRoute {
  when: string;
  engine: DispatchEngine;
  templateId?: string;
  v3Adapter?: string;
  coreMethod?: string;
  action?: string;
  surfaces?: string[];
}

/**
 * One Q1 routing question. Only `name` (the answer key) and the optional
 * `condition` (a shared-grammar gate) drive resolution; presentation is the
 * `prompt` face's concern. Kept structurally minimal so this operation does not
 * couple to the full `collect-inputs` `QuestionSpec`.
 */
export interface RouteQuestion {
  name: string;
  condition?: ExpressionNode;
}

/** The parsed per-kind `selector.json`: the Q1 tree + the routing table. */
export interface SelectorSpec {
  questions: RouteQuestion[];
  routes: SelectorRoute[];
}

/**
 * One interactive Q1 prompt's outcome: the chosen dimension `value`, or the
 * host's `back` request. On `back` the walk re-asks the previous interactive
 * dimension (spec WCS-14); a `back` at the first prompt cancels the walk.
 */
export type PromptResult = { kind: "value"; value: string } | { kind: "back" };

/**
 * The narrow port this operation actually uses (interface-segregation): it does
 * not depend on the full `ScaffoldRuntime` and never opens the package. The
 * `prompt` face serves the interactive Q1 walk; route resolution is otherwise
 * pure (spec §"Inputs" port table). Language is not read here — it is the
 * downstream `collect-inputs` Q0 question (ADR-0014 Amendment 2).
 */
export interface RouteResolverPort {
  /**
   * The walk source's interactive (un-pre-filled) Q1 prompt. `step` is the
   * 1-based position among prompts shown so far (the host shows a Back button
   * when `step > 1`); the result is the chosen value or a host `back` request.
   */
  prompt(question: RouteQuestion, step: number): Promise<PromptResult>;
  /** Evaluate `featureFlag('…')` inside a route predicate / question condition. */
  featureFlag(name: string): boolean;
  /** Membership test for the v4 world (descriptor-derived, spec INV-5). */
  v4Registry(templateId: string): boolean;
  /** Membership test for the frozen v3 generator allow-list. */
  v3Registry(templateId: string): boolean;
  /** Membership test for the frozen v3 core-method allow-list. */
  v3CoreMethodRegistry(coreMethod: string): boolean;
}

/** The resolved dispatch outcome (spec §"Outputs"). */
export interface BuildTarget {
  /** The `<kind>/<id>` (v4) / v3 `TemplateNames` value / `coreMethod` / `action`. */
  templateId: string;
  /** Which world dispatch hands off to. */
  engine: DispatchEngine;
  /**
   * The Q1 dimension picks that produced this target (pre-filled or prompted).
   * The downstream `engine=v3` adapter pre-fills the v3 inputs from this so the
   * v3 question tree skips Q1 and asks only Q2 (walk-create-selector Notes); a
   * `surface-action` carries the picks but scaffolds nothing.
   */
  answers?: Record<string, string>;
}

/** `UserError` name: a resolved/supplied `templateId` belongs to no world (spec AC-10, INV-6). */
export const BUILD_TARGET_UNKNOWN_TEMPLATE = "BuildTargetUnknownTemplate";
/** `UserError` name: a route carries the wrong engine-specific key set (spec AC-11, invariant 12). */
export const BUILD_TARGET_MALFORMED_ROUTE = "BuildTargetMalformedRoute";
/** `UserError` name: a `v4` route's `templateId` has no descriptor (spec AC-12, invariant 17). */
export const BUILD_TARGET_DANGLING_V4_ROUTE = "BuildTargetDanglingV4Route";
/** `UserError` name: no route's predicate matched the resolved answers (spec INV-6, no silent fallback). */
export const BUILD_TARGET_NO_MATCHING_ROUTE = "BuildTargetNoMatchingRoute";
/** `UserError` name: a non-interactive walk lacks a required gated dimension (spec AC-03b). */
export const BUILD_TARGET_MISSING_DIMENSION = "BuildTargetMissingDimension";
/** `UserError` name: the user backed out of the first Q1 prompt, cancelling the walk (spec WCS-14). */
export const BUILD_TARGET_WALK_CANCELLED = "BuildTargetWalkCancelled";

function userError(name: string, message: string): UserError {
  return new UserError({ source: SOURCE, name, message });
}

/** Compose the evaluator's port from this operation's narrow port: no whitelist functions, flags via `featureFlag`. */
function exprPort(port: RouteResolverPort): ExpressionRuntimePort {
  return {
    functions: () => undefined,
    flags: (name) => port.featureFlag(name),
  };
}

/**
 * Seed every declared Q1 question name with `NULL_VALUE`, then overlay the known
 * answers (spec INV-3 scope). So a route/condition referencing an unasked
 * dimension reads as `null` (a false `==` test), while a truly undeclared
 * identifier is still an evaluator error.
 */
function buildScope(declared: string[], answers: Record<string, string>): Scope {
  const scope: Scope = {};
  for (const name of declared) {
    scope[name] = NULL_VALUE;
  }
  for (const [k, v] of Object.entries(answers)) {
    scope[k] = v;
  }
  return scope;
}

/**
 * Reject a route whose engine does not carry exactly its own id key (spec AC-11
 * / invariant 12): the required key must be present and no other engine's key
 * may appear. Returns a human-readable reason, or `undefined` when well-formed.
 */
function malformedRouteReason(r: SelectorRoute): string | undefined {
  const has = {
    templateId: r.templateId !== undefined,
    v3Adapter: r.v3Adapter !== undefined,
    coreMethod: r.coreMethod !== undefined,
    action: r.action !== undefined,
  };
  switch (r.engine) {
    case "v4":
      if (!has.templateId) {
        return "engine 'v4' requires 'templateId'";
      }
      if (has.v3Adapter || has.coreMethod || has.action) {
        return "engine 'v4' must not carry 'v3Adapter' / 'coreMethod' / 'action'";
      }
      return undefined;
    case "v3":
      if (!has.v3Adapter) {
        return "engine 'v3' requires 'v3Adapter'";
      }
      if (has.coreMethod || has.action) {
        return "engine 'v3' must not carry 'coreMethod' / 'action'";
      }
      return undefined;
    case "v3-core-method":
      if (!has.coreMethod) {
        return "engine 'v3-core-method' requires 'coreMethod'";
      }
      if (has.templateId || has.v3Adapter || has.action) {
        return "engine 'v3-core-method' must not carry 'templateId' / 'v3Adapter' / 'action'";
      }
      return undefined;
    case "surface-action":
      if (!has.action) {
        return "engine 'surface-action' requires 'action'";
      }
      if (has.templateId || has.v3Adapter || has.coreMethod) {
        return "engine 'surface-action' must not carry 'templateId' / 'v3Adapter' / 'coreMethod'";
      }
      return undefined;
  }
}

/**
 * Build/load gate over the whole routing table (spec AC-11/AC-12, run before any
 * route matches so a malformed selector is rejected regardless of the entry):
 * every route is well-shaped, and every `v4` route resolves to a present
 * descriptor (`v4Registry`).
 */
function validateRoutes(routes: SelectorRoute[], port: RouteResolverPort): Result<void, FxError> {
  for (const r of routes) {
    const reason = malformedRouteReason(r);
    if (reason !== undefined) {
      return err(userError(BUILD_TARGET_MALFORMED_ROUTE, `route '${r.when}': ${reason}`));
    }
    if (r.engine === "v4" && r.templateId !== undefined && !port.v4Registry(r.templateId)) {
      return err(
        userError(
          BUILD_TARGET_DANGLING_V4_ROUTE,
          `route '${r.when}' targets v4 templateId '${r.templateId}', but no descriptor for it is present`
        )
      );
    }
  }
  return ok(undefined);
}

/**
 * Walk Q1 prefill-aware (the `walk` source, spec AC-03/03a/03b): for each
 * question whose condition holds, take a pre-filled answer as-is (no prompt),
 * else prompt it when `interactive`, else fail with a missing-dimension
 * `UserError`. Interactive, partial pre-fill, and the non-interactive batch are
 * one code path (spec INV-3).
 *
 * The walk is back-capable (spec WCS-14): each interactive prompt pushes a
 * restore point (its index + the answers snapshot taken *before* its pick), and
 * a host `back` pops to the previous interactive dimension and re-asks it. The
 * `step` passed to `prompt` is `history.length + 1`, so the host shows a Back
 * button from the second prompt on, never on the first. Pre-filled and skipped
 * dimensions push no restore point, so `back` steps over them.
 */
async function walkWithPrefill(
  selector: SelectorSpec,
  prefilled: Record<string, string>,
  interactive: boolean,
  port: RouteResolverPort
): Promise<Result<Record<string, string>, FxError>> {
  const declared = selector.questions.map((q) => q.name);
  let answers: Record<string, string> = {};
  const history: { index: number; snapshot: Record<string, string> }[] = [];
  let index = 0;
  while (index < selector.questions.length) {
    const q = selector.questions[index];
    if (q.condition !== undefined) {
      const gate = evaluateExpression(q.condition, buildScope(declared, answers), exprPort(port));
      if (gate.isErr()) {
        return err(gate.error);
      }
      if (gate.value !== true) {
        index++;
        continue;
      }
    }
    if (prefilled[q.name] !== undefined) {
      answers[q.name] = prefilled[q.name];
      index++;
      continue;
    }
    if (!interactive) {
      return err(
        userError(
          BUILD_TARGET_MISSING_DIMENSION,
          `required dimension '${q.name}' was not provided (non-interactive)`
        )
      );
    }
    const outcome = await port.prompt(q, history.length + 1);
    if (outcome.kind === "back") {
      const restore = history.pop();
      if (restore === undefined) {
        return err(userError(BUILD_TARGET_WALK_CANCELLED, "the create selector was cancelled"));
      }
      answers = restore.snapshot;
      index = restore.index;
      continue;
    }
    history.push({ index, snapshot: { ...answers } });
    answers[q.name] = outcome.value;
    index++;
  }
  return ok(answers);
}

/** Match the first route whose `when` predicate is true against the collected answers. */
function matchRoute(
  selector: SelectorSpec,
  answers: Record<string, string>,
  port: RouteResolverPort
): Result<SelectorRoute, FxError> {
  const scope = buildScope(
    selector.questions.map((q) => q.name),
    answers
  );
  for (const r of selector.routes) {
    const hit = evaluateExpression({ expr: r.when }, scope, exprPort(port));
    if (hit.isErr()) {
      return err(hit.error);
    }
    if (hit.value === true) {
      return ok(r);
    }
  }
  return err(
    userError(
      BUILD_TARGET_NO_MATCHING_ROUTE,
      "no selector route matched the resolved answers (no silent fallback)"
    )
  );
}

/** Dispatch a matched route to its engine + id key (no `descriptorLanguages` read here, spec AC-06). */
function dispatchRoute(
  r: SelectorRoute,
  port: RouteResolverPort
): Result<{ templateId: string; engine: DispatchEngine }, FxError> {
  switch (r.engine) {
    case "v4":
    case "v3":
      if (r.templateId === undefined) {
        return err(
          userError(BUILD_TARGET_MALFORMED_ROUTE, `route '${r.when}': missing 'templateId'`)
        );
      }
      return ok({ templateId: r.templateId, engine: r.engine });
    case "v3-core-method":
      if (r.coreMethod === undefined) {
        return err(
          userError(BUILD_TARGET_MALFORMED_ROUTE, `route '${r.when}': missing 'coreMethod'`)
        );
      }
      if (!port.v3CoreMethodRegistry(r.coreMethod)) {
        return err(
          userError(
            BUILD_TARGET_UNKNOWN_TEMPLATE,
            `core method '${r.coreMethod}' is not in the v3 core-method allow-list`
          )
        );
      }
      return ok({ templateId: r.coreMethod, engine: r.engine });
    case "surface-action":
      if (r.action === undefined) {
        return err(userError(BUILD_TARGET_MALFORMED_ROUTE, `route '${r.when}': missing 'action'`));
      }
      return ok({ templateId: r.action, engine: r.engine });
  }
}

/**
 * Resolve a create/modify `selector.json` walk into a dispatched `BuildTarget`.
 *
 * The walk is the only source (ADR-0014 Amendment 2): each gated Q1 dimension is
 * taken from `prefilled` when present (used as-is, no prompt), else prompted when
 * `interactive`, else a missing-dimension `UserError`. `atk new` is a bare
 * interactive walk; `atk add` / CodeLens / modify are the same walk pre-filled
 * over their kind's `selector.json`. Language is not resolved here — it is the
 * downstream `collect-inputs` Q0 question (ADR-0016 decision 5).
 *
 * @param selector    the parsed per-kind `selector.json` (the Q1 tree + routes)
 * @param prefilled   Q1 dimensions known up front; each is used as-is, never prompted
 * @param interactive whether an un-pre-filled gated dimension may be prompted
 * @param port        the narrow injected `RouteResolverPort`
 * @returns `ok(BuildTarget)`, or a `UserError` (unknown templateId, missing
 *          dimension, malformed / dangling / no-matching route).
 */
export async function resolveBuildTarget(
  selector: SelectorSpec,
  prefilled: Record<string, string>,
  interactive: boolean,
  port: RouteResolverPort
): Promise<Result<BuildTarget, FxError>> {
  // The whole routing table is gated first (AC-11/AC-12), before any match.
  const routesOk = validateRoutes(selector.routes, port);
  if (routesOk.isErr()) {
    return err(routesOk.error);
  }

  // One prefill-aware walk covers interactive, partial pre-fill, and the
  // non-interactive batch (spec INV-3): pre-filled dimensions are used as-is,
  // un-pre-filled ones prompted when interactive, else a missing-dimension error.
  const walked = await walkWithPrefill(selector, prefilled, interactive, port);
  if (walked.isErr()) {
    return err(walked.error);
  }
  const answers = walked.value;

  const matched = matchRoute(selector, answers, port);
  if (matched.isErr()) {
    return err(matched.error);
  }
  const dispatched = dispatchRoute(matched.value, port);
  if (dispatched.isErr()) {
    return err(dispatched.error);
  }

  return ok({ ...dispatched.value, answers });
}
