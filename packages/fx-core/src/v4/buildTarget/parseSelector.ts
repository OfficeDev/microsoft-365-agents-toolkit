// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, UserError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { ExpressionNode } from "../expression/evaluateExpression";
import { DispatchEngine, RouteQuestion, SelectorRoute, SelectorSpec } from "./resolveBuildTarget";

/**
 * The load face of `resolve-build-target`: project a raw, parsed `selector.json`
 * (`unknown`) onto the `SelectorSpec` that `resolveBuildTarget` consumes.
 *
 * It keeps only what routing needs — each question's `{ name, condition? }`
 * (the presentation fields `type` / `title` / `staticOptions` / `keyPrefix` are
 * the surface's concern, spec INV-3) and each route's `when` + closed-set
 * `engine` + that engine's own key. Engine-key *completeness* (invariant 12) is
 * deliberately left to the load gate AC-11 inside `resolveBuildTarget`, so the
 * parser and the resolver compose without duplicating that check.
 *
 * Spec: docs/03-specs/operations/scaffolding/resolve-build-target.md
 *       (AC-19 presentation-stripped, AC-20 malformed)
 *
 * v4-owned (INV-8): imports no v3 symbol. Realized with predicates only
 * (`isRecord` / `isExpressionNode` / `isDispatchEngine`), never an unchecked
 * `as` cast — the same untrusted-JSON discipline `validate-template-package`
 * uses.
 */

const SOURCE = "Scaffold";

/** `UserError` name: a `selector.json` is not the parseable shape (an authoring / load break). */
export const BUILD_TARGET_MALFORMED_SELECTOR = "BuildTargetMalformedSelector";

function userError(message: string): UserError {
  return new UserError({ source: SOURCE, name: BUILD_TARGET_MALFORMED_SELECTOR, message });
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Read a record field as a string, or `undefined` when absent / non-string. */
function stringField(record: Record<string, unknown>, key: string): string | undefined {
  const value = record[key];
  return typeof value === "string" ? value : undefined;
}

/** A `string[]` view of an `unknown`, or `undefined` when it is not an all-string array. */
function stringArray(value: unknown): string[] | undefined {
  if (Array.isArray(value) && value.every((item): item is string => typeof item === "string")) {
    return value;
  }
  return undefined;
}

/** Membership test for the closed dispatch-engine set (spec INV-4). */
function isDispatchEngine(value: unknown): value is DispatchEngine {
  return (
    value === "v4" || value === "v3" || value === "v3-core-method" || value === "surface-action"
  );
}

/**
 * A structural check for the authored `ExpressionNode` union (raw `expr` or one
 * of its sugar forms). The grammar itself is validated later by
 * `evaluate-expression`; here we only confirm the `condition` is one of the
 * closed forms so it can be kept verbatim without an `as` cast.
 */
function isExpressionNode(value: unknown): value is ExpressionNode {
  if (!isRecord(value)) {
    return false;
  }
  return (
    typeof value.expr === "string" ||
    typeof value.from === "string" ||
    typeof value.featureFlag === "string" ||
    typeof value.capability === "string" ||
    isRecord(value.equals) ||
    isRecord(value.enum) ||
    Array.isArray(value.anyOf)
  );
}

/** Project one raw question onto `{ name, condition? }` (presentation dropped, AC-19). */
function parseQuestion(raw: unknown): Result<RouteQuestion, FxError> {
  if (!isRecord(raw)) {
    return err(userError("a selector question must be an object"));
  }
  const name = stringField(raw, "name");
  if (name === undefined) {
    return err(userError("a selector question must have a string 'name'"));
  }
  const condition = raw.condition;
  if (condition === undefined) {
    return ok({ name });
  }
  if (!isExpressionNode(condition)) {
    return err(userError(`selector question '${name}' has a malformed 'condition'`));
  }
  return ok({ name, condition });
}

/**
 * Project one raw route onto `{ when, engine, +engine-key? }` (AC-19). The
 * `when` and the closed-set `engine` are required; the engine-specific keys are
 * carried verbatim when present, and their *completeness* is the resolver's
 * gate (AC-11), not this parser's.
 */
function parseRoute(raw: unknown): Result<SelectorRoute, FxError> {
  if (!isRecord(raw)) {
    return err(userError("a selector route must be an object"));
  }
  const when = stringField(raw, "when");
  if (when === undefined) {
    return err(userError("a selector route must have a string 'when'"));
  }
  const engine = raw.engine;
  if (!isDispatchEngine(engine)) {
    return err(userError(`selector route '${when}' has an invalid 'engine'`));
  }
  const route: SelectorRoute = { when, engine };
  const templateId = stringField(raw, "templateId");
  if (templateId !== undefined) {
    route.templateId = templateId;
  }
  const v3Adapter = stringField(raw, "v3Adapter");
  if (v3Adapter !== undefined) {
    route.v3Adapter = v3Adapter;
  }
  const coreMethod = stringField(raw, "coreMethod");
  if (coreMethod !== undefined) {
    route.coreMethod = coreMethod;
  }
  const action = stringField(raw, "action");
  if (action !== undefined) {
    route.action = action;
  }
  const surfaces = stringArray(raw.surfaces);
  if (surfaces !== undefined) {
    route.surfaces = surfaces;
  }
  return ok(route);
}

/**
 * Parse a raw `selector.json` into a `SelectorSpec` (the Q1 questions + the
 * routing table), dropping presentation fields and keeping each route's
 * routing keys (AC-19); a non-object root, a non-array `questions` / `routes`,
 * or a malformed question / route is an explicit `UserError` (AC-20).
 */
export function parseSelectorSpec(raw: unknown): Result<SelectorSpec, FxError> {
  if (!isRecord(raw)) {
    return err(userError("selector.json must be a JSON object"));
  }
  if (!Array.isArray(raw.questions)) {
    return err(userError("selector.json 'questions' must be an array"));
  }
  if (!Array.isArray(raw.routes)) {
    return err(userError("selector.json 'routes' must be an array"));
  }
  const questions: RouteQuestion[] = [];
  for (const rawQuestion of raw.questions) {
    const parsed = parseQuestion(rawQuestion);
    if (parsed.isErr()) {
      return err(parsed.error);
    }
    questions.push(parsed.value);
  }
  const routes: SelectorRoute[] = [];
  for (const rawRoute of raw.routes) {
    const parsed = parseRoute(rawRoute);
    if (parsed.isErr()) {
      return err(parsed.error);
    }
    routes.push(parsed.value);
  }
  return ok({ questions, routes });
}

/**
 * The presentation projection of `selector.json` — the sibling of
 * `parseSelectorSpec` that keeps what the surface *renders* (`title` /
 * `placeholder` / `staticOptions`) rather than what routing *matches*. Both read
 * the same authored file: `walk-create-selector` drives the live Q1 prompt face
 * from this projection, while `resolveBuildTarget` routes off `parseSelectorSpec`.
 *
 * Spec: docs/03-specs/operations/scaffolding/walk-create-selector.md (WCS-07).
 */

/** One selectable option's presentation: id + label, optional detail / group, and an environment `condition`. */
export interface PresentationOption {
  id: string;
  label: string;
  detail?: string;
  groupName?: string;
  condition?: ExpressionNode;
}

/** One Q1 question's presentation: its name, optional title / placeholder, and its static options. */
export interface PresentationQuestion {
  name: string;
  title?: string;
  placeholder?: string;
  staticOptions: PresentationOption[];
}

/** The presentation projection of the whole create selector. */
export interface SelectorPresentation {
  questions: PresentationQuestion[];
}

/** Project one raw option onto `{ id, label, detail?, groupName?, condition? }`; id + label are required. */
function parsePresentationOption(raw: unknown): Result<PresentationOption, FxError> {
  if (!isRecord(raw)) {
    return err(userError("a selector option must be an object"));
  }
  const id = stringField(raw, "id");
  if (id === undefined) {
    return err(userError("a selector option must have a string 'id'"));
  }
  const label = stringField(raw, "label");
  if (label === undefined) {
    return err(userError(`selector option '${id}' must have a string 'label'`));
  }
  const option: PresentationOption = { id, label };
  const detail = stringField(raw, "detail");
  if (detail !== undefined) {
    option.detail = detail;
  }
  const groupName = stringField(raw, "groupName");
  if (groupName !== undefined) {
    option.groupName = groupName;
  }
  const condition = raw.condition;
  if (condition !== undefined) {
    if (!isExpressionNode(condition)) {
      return err(userError(`selector option '${id}' has a malformed 'condition'`));
    }
    option.condition = condition;
  }
  return ok(option);
}

/** Project one raw question onto `{ name, title?, placeholder?, staticOptions }`; name is required. */
function parsePresentationQuestion(raw: unknown): Result<PresentationQuestion, FxError> {
  if (!isRecord(raw)) {
    return err(userError("a selector question must be an object"));
  }
  const name = stringField(raw, "name");
  if (name === undefined) {
    return err(userError("a selector question must have a string 'name'"));
  }
  const question: PresentationQuestion = { name, staticOptions: [] };
  const title = stringField(raw, "title");
  if (title !== undefined) {
    question.title = title;
  }
  const placeholder = stringField(raw, "placeholder");
  if (placeholder !== undefined) {
    question.placeholder = placeholder;
  }
  if (raw.staticOptions !== undefined) {
    if (!Array.isArray(raw.staticOptions)) {
      return err(userError(`selector question '${name}' has a non-array 'staticOptions'`));
    }
    for (const rawOption of raw.staticOptions) {
      const parsed = parsePresentationOption(rawOption);
      if (parsed.isErr()) {
        return err(parsed.error);
      }
      question.staticOptions.push(parsed.value);
    }
  }
  return ok(question);
}

/**
 * Project a raw `selector.json` onto its `SelectorPresentation` (each question's
 * `title` / `placeholder` / `staticOptions`); a non-object root, a non-array
 * `questions`, or a malformed question / option is an explicit `UserError`
 * (mirrors `parseSelectorSpec`'s AC-20 discipline).
 */
export function parseSelectorPresentation(raw: unknown): Result<SelectorPresentation, FxError> {
  if (!isRecord(raw)) {
    return err(userError("selector.json must be a JSON object"));
  }
  if (!Array.isArray(raw.questions)) {
    return err(userError("selector.json 'questions' must be an array"));
  }
  const questions: PresentationQuestion[] = [];
  for (const rawQuestion of raw.questions) {
    const parsed = parsePresentationQuestion(rawQuestion);
    if (parsed.isErr()) {
      return err(parsed.error);
    }
    questions.push(parsed.value);
  }
  return ok({ questions });
}
