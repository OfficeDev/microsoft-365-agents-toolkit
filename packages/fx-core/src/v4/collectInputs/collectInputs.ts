// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError, UserError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { EvalValue, ExpressionNode, NULL_VALUE, Scope } from "../expression/evaluateExpression";
import { Answers } from "../model/dataModel";

/**
 * The v4 input collection — one template's `questions.json` → the resolved answers.
 *
 * Spec: docs/03-specs/operations/scaffolding/collect-inputs.md
 * Decision: docs/02-architecture/adr/ADR-0016-declarative-template-format.md
 *           (decision 2 optionsSchema, 5 language axis, 6 native QuestionSpec)
 *
 * This module is part of the v4 world; it imports no v3 symbol. It realizes
 * ADR-0016 decision 6: the authored fields **are** the runtime model — a native
 * `QuestionSpec[]` walked by a surface-neutral driver, with no rehydration into
 * v3's `IQTreeNode` and no `func` / `onDidSelection` callback that could change a
 * question's shape at runtime (spec INV-1).
 *
 * It is **one** behavior (questions → answers), distinct from
 * `build-render-context` (answers → render vars); it computes no render variable
 * and writes no file. Every `condition` / `optionsFromParams` is routed through
 * the shared evaluator (`evaluate-expression`); this operation adds no operator
 * (spec INV-6).
 */

const SOURCE = "Scaffold";

/**
 * Display labels for the Q0 language axis (ADR-0016 decision 5). `descriptor.languages`
 * carries lowercase ids (`"typescript"`), but the v3 UI shows proper-cased labels
 * (`"TypeScript"`), so the v4 axis mirrors them — a v4-local copy of v3's
 * `LanguageOptionMap` labels (INV-7 forbids importing the v3 symbol). An id with no
 * entry falls back to itself.
 */
const LANGUAGE_LABELS: Record<string, string> = {
  javascript: "JavaScript",
  typescript: "TypeScript",
  csharp: "C#",
  python: "Python",
};

/** An authored visibility / value guard — the same closed form the evaluator parses. */
export type ConditionNode = ExpressionNode;

/**
 * An identity-only option (spec INV-2): an `id`, presentational fields, and an
 * optional visibility `condition`. **No** configuration payload hangs off it
 * (no v3 `option.data`); computed values flow through provider `derived.*`.
 */
export interface OptionItem {
  id: string;
  label?: string;
  description?: string;
  detail?: string;
  groupName?: string;
  condition?: ConditionNode;
  keyPrefix?: string;
}

/** The native question kinds the surface-neutral driver renders (ADR-0016 decision 6). */
export type QuestionType =
  | "singleSelect"
  | "multiSelect"
  | "text"
  | "confirm"
  | "singleFile"
  | "folder"
  | "singleFileOrText";

/** A validator reference: the registry name, or `{ use, params }` (the `"uri"` shorthand). */
export interface ValidationSpec {
  use: string;
  params?: Record<string, string>;
}

/**
 * One authored question (spec INV-1: `questions.json` parses straight into this).
 * Exactly one of `staticOptions` / `optionsFrom` may be present (spec INV-3).
 */
export interface QuestionSpec {
  name: string;
  type: QuestionType;
  title?: string;
  placeholder?: string;
  prompt?: string;
  default?: string;
  validation?: string | ValidationSpec;
  staticOptions?: OptionItem[];
  optionsFrom?: string;
  optionsFromParams?: Record<string, ConditionNode>;
  skipSingleOption?: boolean;
  optional?: boolean;
  condition?: ConditionNode;
  keyPrefix?: string;
}

/** The Q2 options JSON Schema — only its identifier domain matters to this operation. */
export interface OptionsSchema {
  properties?: Record<string, unknown>;
}

/** What an `optionsFrom` provider yields: identity-only options + optional `derived.*`. */
export interface ResolvedOptions {
  options: OptionItem[];
  derived?: Record<string, string>;
}

/**
 * An engine-registered `optionsFrom` provider (spec §3.3.2). It owns its own I/O
 * (the read-only `http` / `InMemoryRuntime` face lives inside the provider, not on
 * the operation), declares its produced keys in `derivedSchema`, and is total.
 */
export interface OptionsProvider {
  derivedSchema?: string[];
  fetch(params: Record<string, string>): Promise<ResolvedOptions> | ResolvedOptions;
}

/** An engine-registered validator (spec §6.4): an error message, or `undefined` when valid. */
export type Validator = (value: string) => string | undefined;

/**
 * One prompt's outcome: the chosen `value`, or the host's `back` request. A
 * `back` re-asks the previous prompted question (spec INPUT-16); a `back` at the
 * first prompt cancels the walk.
 */
export type Asked<T> = { kind: "value"; value: T } | { kind: "back" };

/** The surface-neutral prompt driver (`ScriptedUI` in tests) — proposal §4.2 / §8. */
export interface PromptUI {
  /**
   * Render one question. `step` is the 1-based position among prompts shown so
   * far (the host shows a Back button when `step > 1`); the result is the chosen
   * value or a host `back` request (spec INPUT-16).
   */
  ask(
    question: QuestionSpec,
    options: OptionItem[] | undefined,
    step?: number
  ): Promise<Result<Asked<string>, FxError>>;
  /**
   * The multi-pick face: a `multiSelect` question resolves to the `string[]` of
   * selected ids (collect-inputs INV-7 / INPUT-15). A distinct face keeps the
   * typed list from ever collapsing into a delimiter-joined scalar.
   */
  askMulti(
    question: QuestionSpec,
    options: OptionItem[] | undefined,
    step?: number
  ): Promise<Result<Asked<string[]>, FxError>>;
}

/**
 * The narrow port (interface-segregation over the full `ScaffoldRuntime`): the
 * prompt driver, the provider / validator registries, and the shared evaluator.
 */
export interface CollectInputsPort {
  ui: PromptUI;
  optionsProvider(providerId: string): OptionsProvider | undefined;
  validator(name: string): Validator | undefined;
  evaluate(node: ConditionNode, scope: Scope): Result<EvalValue, FxError>;
}

/**
 * `SystemError` names — an engine-side break `validate-template-package` should
 * have caught at build time (spec Outputs).
 */
export const INPUT_BOTH_OPTION_SOURCES = "InputBothOptionSources";
export const INPUT_UNKNOWN_PROVIDER = "InputUnknownProvider";
export const INPUT_UNKNOWN_VALIDATOR = "InputUnknownValidator";
export const INPUT_FORWARD_DERIVED_REFERENCE = "InputForwardDerivedReference";
export const INPUT_PROVIDER_FAILED = "InputProviderFailed";

/** `UserError` name — an input-side, user-fixable validation failure (spec Outputs). */
export const INPUT_VALIDATION_FAILED = "InputValidationFailed";

/** `UserError` name — a `back` at the very first prompt cancelled the walk (spec INPUT-16). */
export const INPUT_WALK_CANCELLED = "InputWalkCancelled";

/** The walk's cancel signal: a `back` at the very first prompt (spec INPUT-16). */
function walkCancelled(): UserError {
  return new UserError({
    source: SOURCE,
    name: INPUT_WALK_CANCELLED,
    message: "the input walk was cancelled by going back from the first question",
  });
}

/**
 * Walk one template's questions into the resolved answer object.
 *
 * @param questions     the native `QuestionSpec[]` (`questions.json`, schema-valid)
 * @param optionsSchema the Q2 options schema (its `properties` = the identifier domain)
 * @param entryParams   the option ids a pre-filled CLI arg / URL supplies, keyed by id
 * @param languages     `descriptor.languages` (the engine-owned language axis, decision 5)
 * @param port          the prompt driver + provider / validator registries + evaluator
 * @returns `ok(Answers)` = each asked question's value ∪ `derived.<id>.<key>`, or a
 *          `UserError` (validation failure) / `SystemError` (engine-side break).
 */
export async function collectInputs(
  questions: QuestionSpec[],
  optionsSchema: OptionsSchema,
  entryParams: Answers,
  languages: string[],
  port: CollectInputsPort
): Promise<Result<Answers, FxError>> {
  // entry.params pre-fills seed the answers up front, so a question's condition
  // (e.g. modify's `mcpServerUrl == null`) sees the supplied value (spec INPUT-12).
  let answers: Answers = { ...entryParams };
  const declared = Object.keys(optionsSchema.properties ?? {});
  // INV-5: a (providerId, normalize(params)) key resolves once per run.
  const providerCache = new Map<string, ResolvedOptions>();
  // INV-4: providers resolve in declaration order; a forward `derived.*` ref is rejected.
  const resolvedProviders = new Set<string>();

  // The back history: one restore point per *prompted* step, holding the answers
  // as they stood *before* that prompt recorded its value. A host `back` pops the
  // top and re-asks that step (spec INPUT-16); skipped / pre-filled / auto-selected
  // steps push nothing, so `back` steps straight over them. The provider caches stay
  // monotonic — re-walking a provider step re-resolves it idempotently (cache hit),
  // and a forward `derived.*` reference is a build-time error, so no snapshot is kept.
  const history: { pos: number; answers: Answers }[] = [];

  // The walk positions: pos 0 = the language axis Q0 (decision 5), pos i ≥ 1 =
  // questions[i - 1]. A single cursor over both lets a `back` from the first
  // question cross into Q0. `step` (the host's Back-button gate) = prompts shown
  // so far + 1 = history.length + 1, so the first prompt is step 1 (no Back button).
  let pos = 0;
  while (pos <= questions.length) {
    if (pos === 0) {
      // Q0 — the language axis (decision 5). A non-singleton list asks Q0; a single
      // concrete language auto-selects; the language-agnostic `["common"]` has no axis.
      if (languages.length > 1) {
        const langQuestion: QuestionSpec = {
          name: "language",
          type: "singleSelect",
          title: "Programming Language",
        };
        const asked = await port.ui.ask(
          langQuestion,
          languages.map((l) => ({ id: l, label: LANGUAGE_LABELS[l] ?? l })),
          history.length + 1
        );
        if (asked.isErr()) {
          return err(asked.error);
        }
        if (asked.value.kind === "back") {
          const restore = history.pop();
          if (restore === undefined) {
            return err(walkCancelled());
          }
          answers = restore.answers;
          pos = restore.pos;
          continue;
        }
        history.push({ pos: 0, answers: { ...answers } });
        answers.language = asked.value.value;
      } else if (languages.length === 1 && languages[0] !== "common") {
        answers.language = languages[0];
      }
      pos++;
      continue;
    }

    const q = questions[pos - 1];

    // INV-3 / INPUT-03: exactly one option source (schema-enforced; guarded here too).
    if (q.staticOptions !== undefined && q.optionsFrom !== undefined) {
      return err(
        systemError(
          INPUT_BOTH_OPTION_SOURCES,
          `question '${q.name}' declares both staticOptions and optionsFrom; exactly one option source is allowed`
        )
      );
    }

    // Identifier domain for this question's guards: the declared option ids
    // (unanswered → NULL_VALUE so `x == null` is a presence test) overlaid with
    // the answers resolved so far (incl. provider `derived.*`).
    const scope = buildScope(declared, answers);

    // INPUT-01: a question whose condition is false is skipped whole.
    if (q.condition !== undefined) {
      const r = port.evaluate(q.condition, scope);
      if (r.isErr()) {
        return err(r.error);
      }
      if (r.value !== true) {
        pos++;
        continue;
      }
    }

    // INPUT-12: an entry.params pre-fill is used as-is, never prompted.
    if (q.name in answers) {
      pos++;
      continue;
    }

    // Resolve the option set (if any) — staticOptions filtered by their own
    // condition, or the named provider's `{ options, derived }`.
    let options: OptionItem[] | undefined;
    if (q.staticOptions !== undefined) {
      const filtered: OptionItem[] = [];
      for (const opt of q.staticOptions) {
        // INPUT-02: an option-level condition hides only that option.
        if (opt.condition !== undefined) {
          const r = port.evaluate(opt.condition, scope);
          if (r.isErr()) {
            return err(r.error);
          }
          if (r.value !== true) {
            continue;
          }
        }
        filtered.push(opt);
      }
      options = filtered;
    } else if (q.optionsFrom !== undefined) {
      // INPUT-05 / INPUT-11: a dynamic / machine-state-dependent list is an
      // engine-registered provider referenced by name, never a condition predicate.
      const provider = port.optionsProvider(q.optionsFrom);
      if (provider === undefined) {
        return err(
          systemError(
            INPUT_UNKNOWN_PROVIDER,
            `optionsFrom '${q.optionsFrom}' on question '${q.name}' is not a registered provider`
          )
        );
      }
      // INPUT-06: optionsFromParams close over prior answers via the shared evaluator.
      const paramsResult = resolveParams(q.optionsFromParams, scope, resolvedProviders, port);
      if (paramsResult.isErr()) {
        return err(paramsResult.error);
      }
      const params = paramsResult.value;
      // INPUT-09 / INV-5: resolve once per (providerId, normalize(params)) per run.
      const cacheKey = `${q.optionsFrom}|${stableStringify(params)}`;
      let resolved = providerCache.get(cacheKey);
      if (resolved === undefined) {
        try {
          resolved = await provider.fetch(params);
        } catch (error) {
          if (error instanceof UserError || error instanceof SystemError) {
            return err(error);
          }
          return err(
            systemError(
              INPUT_PROVIDER_FAILED,
              `optionsFrom '${q.optionsFrom}' on question '${q.name}' failed: ${errorMessage(error)}`
            )
          );
        }
        providerCache.set(cacheKey, resolved);
      }
      options = resolved.options;
      // INPUT-07 / INV-4: merge derived under the reserved derived.<provider-id>.<key>
      // namespace — two providers cannot collide by construction.
      if (resolved.derived !== undefined) {
        for (const [key, value] of Object.entries(resolved.derived)) {
          answers[`derived.${q.optionsFrom}.${key}`] = value;
        }
      }
      resolvedProviders.add(q.optionsFrom);
    }

    // INPUT-04: skipSingleOption auto-selects a sole option without prompting.
    if (options !== undefined && q.skipSingleOption === true && options.length === 1) {
      answers[q.name] = options[0].id;
      pos++;
      continue;
    }

    // INPUT-15 / INV-7: a multiSelect resolves to the string[] of selected ids
    // through the dedicated multi-pick face, so the typed list never collapses
    // into a delimiter-joined scalar. Every other kind records a scalar string.
    if (q.type === "multiSelect") {
      const picked = await port.ui.askMulti(q, options, history.length + 1);
      if (picked.isErr()) {
        return err(picked.error);
      }
      if (picked.value.kind === "back") {
        const restore = history.pop();
        if (restore === undefined) {
          return err(walkCancelled());
        }
        answers = restore.answers;
        pos = restore.pos;
        continue;
      }
      history.push({ pos, answers: { ...answers } });
      answers[q.name] = picked.value.value;
      pos++;
      continue;
    }

    // Render the question through the surface-neutral driver.
    const asked = await port.ui.ask(q, options, history.length + 1);
    if (asked.isErr()) {
      return err(asked.error);
    }
    if (asked.value.kind === "back") {
      const restore = history.pop();
      if (restore === undefined) {
        return err(walkCancelled());
      }
      answers = restore.answers;
      pos = restore.pos;
      continue;
    }
    const value = asked.value.value;

    // INPUT-10: validation — the shorthand `"uri"` normalizes to `{ use: "uri" }`
    // and resolves in the validator registry; a failure is a UserError naming
    // the question (the surface re-shows the prompt interactively).
    if (q.validation !== undefined) {
      const validatorName = typeof q.validation === "string" ? q.validation : q.validation.use;
      const validator = port.validator(validatorName);
      if (validator === undefined) {
        return err(
          systemError(
            INPUT_UNKNOWN_VALIDATOR,
            `validation '${validatorName}' on question '${q.name}' is not a registered validator`
          )
        );
      }
      const message = validator(value);
      if (message !== undefined) {
        return err(
          new UserError({
            source: SOURCE,
            name: INPUT_VALIDATION_FAILED,
            message: `'${q.name}': ${message}`,
          })
        );
      }
    }

    history.push({ pos, answers: { ...answers } });
    answers[q.name] = value;
    pos++;
  }

  return ok(answers);
}

/**
 * The evaluator scope for a question's guards: every declared option id seeded
 * `NULL_VALUE` (declared-but-unanswered → a `== null` presence test), overlaid
 * with the answers resolved so far. A truly undeclared id stays absent, so a
 * typo is still an `EXPR_UNDECLARED_IDENTIFIER` rather than a silent null.
 */
function buildScope(declared: string[], answers: Answers): Scope {
  const scope: Scope = {};
  for (const id of declared) {
    scope[id] = NULL_VALUE;
  }
  for (const [key, value] of Object.entries(answers)) {
    // INV-7: a multiSelect answer (string[]) is off the scalar grammar — it
    // reaches render vars and step `with`, but never the expression scope, so an
    // unanswered scalar discriminator stays NULL_VALUE rather than an array.
    if (Array.isArray(value)) {
      continue;
    }
    scope[key] = value;
  }
  return scope;
}

/**
 * Resolve `optionsFromParams` to a flat `{ name: value }` via the shared evaluator
 * (the same `{from}` / `{expr}` forms as a `condition`). A `derived.<id>.<key>`
 * param naming a provider that has not resolved yet is a forward reference
 * (spec INPUT-08 / INV-4) — providers resolve in declaration order.
 */
function resolveParams(
  optionsFromParams: Record<string, ConditionNode> | undefined,
  scope: Scope,
  resolvedProviders: Set<string>,
  port: CollectInputsPort
): Result<Record<string, string>, FxError> {
  const params: Record<string, string> = {};
  if (optionsFromParams === undefined) {
    return ok(params);
  }
  for (const [key, node] of Object.entries(optionsFromParams)) {
    if ("from" in node && node.from.startsWith("derived.")) {
      const producer = node.from.split(".")[1];
      if (!resolvedProviders.has(producer)) {
        return err(
          systemError(
            INPUT_FORWARD_DERIVED_REFERENCE,
            `param '${key}' references '${node.from}' before provider '${producer}' resolves`
          )
        );
      }
    }
    const r = port.evaluate(node, scope);
    if (r.isErr()) {
      return err(r.error);
    }
    params[key] = typeof r.value === "string" ? r.value : String(r.value);
  }
  return ok(params);
}

/** A stable provider-cache key: params serialized with sorted keys (INV-5 normalize). */
function stableStringify(params: Record<string, string>): string {
  const sorted: Record<string, string> = {};
  for (const key of Object.keys(params).sort()) {
    sorted[key] = params[key];
  }
  return JSON.stringify(sorted);
}

function systemError(name: string, message: string): SystemError {
  return new SystemError({ source: SOURCE, name, message });
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
