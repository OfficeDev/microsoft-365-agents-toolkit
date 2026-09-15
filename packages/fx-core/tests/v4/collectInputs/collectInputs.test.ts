// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError, UserError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { assert } from "vitest";
import {
  ConditionNode,
  ExpressionRuntimePort,
  WhitelistFn,
  evaluateExpression,
} from "../../../src/v4/expression/evaluateExpression";
import {
  Asked,
  CollectInputsPort,
  INPUT_BOTH_OPTION_SOURCES,
  INPUT_FORWARD_DERIVED_REFERENCE,
  INPUT_PROVIDER_DERIVED_SCHEMA_VIOLATION,
  INPUT_PROVIDER_FAILED,
  INPUT_VALIDATION_FAILED,
  INPUT_UNKNOWN_PROVIDER,
  INPUT_UNKNOWN_VALIDATOR,
  INPUT_WALK_CANCELLED,
  OptionItem,
  OptionsSource,
  OptionsProvider,
  PromptUI,
  QuestionSpec,
  ResolvedOptions,
  Validator,
  collectInputs,
  walkInputs,
} from "../../../src/v4/collectInputs/collectInputs";

/**
 * Tests for docs/03-specs/operations/scaffolding/collect-inputs.md.
 * One `it` per INPUT-* acceptance-criteria row. v4-isolated (no v3 import).
 *
 * The `evaluate` port face is backed by the real shared evaluator over an
 * in-memory ExpressionRuntimePort, so conditions exercise the real grammar.
 */

// --- in-memory fakes of the narrow CollectInputsPort faces ---

/** The pure expression port: a small whitelist + a configurable feature-flag map. */
class ExprPort implements ExpressionRuntimePort {
  private readonly flagMap: Record<string, boolean>;
  constructor(flagMap: Record<string, boolean> = {}) {
    this.flagMap = flagMap;
  }
  functions(name: string): WhitelistFn | undefined {
    const table: Record<string, WhitelistFn> = {
      safeUpper: (s) => (s ?? "").toUpperCase(),
      safeLower: (s) => (s ?? "").toLowerCase(),
    };
    return table[name];
  }
  flags(name: string): boolean {
    return this.flagMap[name] ?? false;
  }
}

/** A scripted prompt driver: returns pre-programmed answers, recording every ask. */
class ScriptedUI implements PromptUI {
  asked: string[] = [];
  lastOptions: Record<string, OptionItem[] | undefined> = {};
  lastValidations: Record<
    string,
    ((value: string) => string | undefined | Promise<string | undefined>) | undefined
  > = {};
  private readonly script: Record<string, string>;
  private readonly multiScript: Record<string, string[]>;
  constructor(script: Record<string, string>, multiScript: Record<string, string[]> = {}) {
    this.script = script;
    this.multiScript = multiScript;
  }
  async ask(
    question: QuestionSpec,
    options: OptionsSource | undefined,
    _step?: number,
    validation?: (value: string) => string | undefined | Promise<string | undefined>
  ): Promise<Result<Asked<string>, FxError>> {
    this.asked.push(question.name);
    this.lastValidations[question.name] = validation;
    const resolvedOptions = await resolveTestOptions(options);
    this.lastOptions[question.name] = resolvedOptions;
    if (question.skipSingleOption === true && resolvedOptions?.length === 1) {
      return ok({ kind: "skip", value: optionId(resolvedOptions[0]) });
    }
    if (question.name in this.script) {
      const value = this.script[question.name];
      const message = await validation?.(value);
      if (message !== undefined) {
        return err(
          new UserError({
            source: "Test",
            name: INPUT_VALIDATION_FAILED,
            message: `'${question.name}': ${message}`,
          })
        );
      }
      return ok({ kind: "value", value });
    }
    return err(
      new UserError({
        source: "Test",
        name: "NoScriptedAnswer",
        message: `no scripted answer for '${question.name}'`,
      })
    );
  }
  async askMulti(
    question: QuestionSpec,
    options: OptionsSource | undefined
  ): Promise<Result<Asked<string[]>, FxError>> {
    this.asked.push(question.name);
    const resolvedOptions = await resolveTestOptions(options);
    this.lastOptions[question.name] = resolvedOptions;
    if (question.skipSingleOption === true && resolvedOptions?.length === 1) {
      return ok({ kind: "skip", value: [optionId(resolvedOptions[0])] });
    }
    if (question.name in this.multiScript) {
      return ok({ kind: "value", value: this.multiScript[question.name] });
    }
    return err(
      new UserError({
        source: "Test",
        name: "NoScriptedAnswer",
        message: `no scripted multi-answer for '${question.name}'`,
      })
    );
  }
}

function optionId(option: OptionItem): string {
  return option.id;
}

async function resolveTestOptions(
  options: OptionsSource | undefined
): Promise<OptionItem[] | undefined> {
  if (options === undefined || Array.isArray(options)) {
    return options;
  }
  return (await options()).options;
}

/** A no-scripted-answer error for the sequence-driven driver. */
function noScripted(name: string): FxError {
  return new UserError({ source: "Test", name: "NoScriptedAnswer", message: name });
}

/** One scripted reply for the sequenced driver: a scalar value, a multi value, a surface skip, or a host back. */
type SeqResponse =
  | { kind: "value"; value: string }
  | { kind: "multi"; value: string[] }
  | { kind: "skip"; value: string }
  | { kind: "back" };

/**
 * A sequence-driven prompt driver: it answers each ask / askMulti from an ordered
 * script (independent of the question name, so a question re-asked after a `back`
 * can get a different answer) and records each call's name + the host `step`.
 */
class SequencedPromptUI implements PromptUI {
  calls: { name: string; step?: number }[] = [];
  private cursor = 0;
  constructor(private readonly responses: SeqResponse[]) {}
  ask(
    question: QuestionSpec,
    _options: OptionsSource | undefined,
    step?: number
  ): Promise<Result<Asked<string>, FxError>> {
    this.calls.push({ name: question.name, step });
    const response = this.responses[this.cursor++];
    if (response === undefined || response.kind === "multi") {
      return Promise.resolve(err(noScripted(question.name)));
    }
    if (response.kind === "back") {
      return Promise.resolve(ok({ kind: "back" }));
    }
    if (response.kind === "skip") {
      return Promise.resolve(ok({ kind: "skip", value: response.value }));
    }
    return Promise.resolve(ok({ kind: "value", value: response.value }));
  }
  askMulti(
    question: QuestionSpec,
    _options: OptionsSource | undefined,
    step?: number
  ): Promise<Result<Asked<string[]>, FxError>> {
    this.calls.push({ name: question.name, step });
    const response = this.responses[this.cursor++];
    if (response === undefined || response.kind === "value" || response.kind === "skip") {
      return Promise.resolve(err(noScripted(question.name)));
    }
    if (response.kind === "back") {
      return Promise.resolve(ok({ kind: "back" }));
    }
    return Promise.resolve(ok({ kind: "value", value: response.value }));
  }
}

/** An in-memory options provider: records its fetch count and the params it saw. */
class FakeProvider implements OptionsProvider {
  fetchCount = 0;
  lastParams: Record<string, string> | undefined;
  derivedSchema?: string[];
  private readonly result: ResolvedOptions;
  constructor(result: ResolvedOptions, derivedSchema?: string[]) {
    this.result = result;
    this.derivedSchema = derivedSchema;
  }
  fetch(params: Record<string, string>): Promise<ResolvedOptions> {
    this.fetchCount++;
    this.lastParams = params;
    return Promise.resolve(this.result);
  }
}

/** The `"uri"` validator: an error message for a non-URI, `undefined` when valid. */
const uriValidator: Validator = (value) => {
  try {
    void new URL(value);
    return undefined;
  } catch {
    return "must be a valid URI";
  }
};

function makePort(opts: {
  ui: PromptUI;
  providers?: Record<string, OptionsProvider>;
  validators?: Record<string, Validator>;
  exprPort?: ExpressionRuntimePort;
}): CollectInputsPort {
  const exprPort = opts.exprPort ?? new ExprPort();
  return {
    ui: opts.ui,
    optionsProvider: (id) => opts.providers?.[id],
    validator: (name) => opts.validators?.[name],
    evaluate: (node, scope) => evaluateExpression(node, scope, exprPort),
  };
}

describe("collectInputs (v4)", () => {
  for (const providerId of ["catalog", "catalog.remote"]) {
    it(`INPUT-37: structured references consume accepted ${providerId} outputs`, async () => {
      const producer = new FakeProvider(
        { options: [{ id: "one" }], derived: { context: "exact value" } },
        ["context"]
      );
      const consumer = new FakeProvider({ options: [{ id: "two" }] });
      const result = await collectInputs(
        [
          { name: "producer", type: "singleSelect", optionsFrom: providerId },
          {
            name: "consumer",
            type: "singleSelect",
            optionsFrom: "consumer",
            optionsFromParams: {
              context: { from: `derived.${providerId}.context` },
              singleton: { anyOf: [{ from: `derived.${providerId}.context` }] },
              nested: {
                anyOf: [{ anyOf: [{ anyOf: [{ from: `derived.${providerId}.context` }] }] }],
              },
            },
          },
        ],
        {},
        {},
        makePort({
          ui: new ScriptedUI({ producer: "one", consumer: "two" }),
          providers: { [providerId]: producer, consumer },
        })
      );
      assert.equal(result._unsafeUnwrap().consumer, "two");
      assert.deepEqual(consumer.lastParams, {
        context: "exact value",
        singleton: "exact value",
        nested: "exact value",
      });
    });
  }

  for (const nesting of ["singleton", "deep", "deep-short-circuit"]) {
    for (const reference of [
      "derived.unknown.context",
      "derived.catalog.remote.context",
      "derived.catalog.undeclared",
    ]) {
      for (const prefilled of [false, true]) {
        it(`INPUT-37: ${nesting} anyOf rejects ${reference} with ${prefilled ? "forged" : "absent"} answers`, async () => {
          const producer = new FakeProvider(
            { options: [{ id: "one" }], derived: { context: "accepted" } },
            ["context"]
          );
          const future = new FakeProvider(
            { options: [{ id: "one" }], derived: { context: "future" } },
            ["context"]
          );
          const consumer = new FakeProvider({ options: [{ id: "two" }] });
          const node: ConditionNode =
            nesting === "singleton"
              ? { anyOf: [{ from: reference }] }
              : { anyOf: [{ anyOf: [{ anyOf: [{ from: reference }] }] }] };
          if (nesting === "deep-short-circuit") {
            node.anyOf.unshift({ equals: { producer: "one" } });
          }
          const result = await collectInputs(
            [
              { name: "producer", type: "singleSelect", optionsFrom: "catalog" },
              {
                name: "consumer",
                type: "singleSelect",
                optionsFrom: "consumer",
                optionsFromParams: { context: node },
              },
              { name: "future", type: "singleSelect", optionsFrom: "catalog.remote" },
            ],
            {},
            prefilled ? { [reference]: "forged" } : {},
            makePort({
              ui: new ScriptedUI({ producer: "one", consumer: "two", future: "one" }),
              providers: { catalog: producer, "catalog.remote": future, consumer },
            })
          );
          assert.equal(consumer.fetchCount, 0);
          assert.equal(future.fetchCount, 0);
          assert.equal(result._unsafeUnwrapErr().name, INPUT_FORWARD_DERIVED_REFERENCE);
        });
      }
    }
  }

  for (const reference of ["derived.unknown.context", "derived.catalog.undeclared"]) {
    it(`INPUT-37: rejects ${reference} despite a matching prefill`, async () => {
      const producer = new FakeProvider(
        { options: [{ id: "one" }], derived: { context: "accepted" } },
        ["context"]
      );
      const consumer = new FakeProvider({ options: [{ id: "two" }] });
      const result = await collectInputs(
        [
          { name: "producer", type: "singleSelect", optionsFrom: "catalog" },
          {
            name: "consumer",
            type: "singleSelect",
            optionsFrom: "consumer",
            optionsFromParams: { context: { from: reference } },
          },
        ],
        {},
        { [reference]: "forged" },
        makePort({
          ui: new ScriptedUI({ producer: "one", consumer: "two" }),
          providers: { catalog: producer, consumer },
        })
      );
      assert.equal(result._unsafeUnwrapErr().name, INPUT_FORWARD_DERIVED_REFERENCE);
      assert.equal(consumer.fetchCount, 0);
    });
  }

  it("INPUT-37: a resolved prefix provider and prefill cannot authorize a future dotted provider", async () => {
    const prefix = new FakeProvider({ options: [{ id: "one" }] });
    const producer = new FakeProvider(
      { options: [{ id: "one" }], derived: { context: "accepted" } },
      ["context"]
    );
    const consumer = new FakeProvider({ options: [{ id: "two" }] });
    const result = await collectInputs(
      [
        { name: "prefix", type: "singleSelect", optionsFrom: "catalog" },
        {
          name: "consumer",
          type: "singleSelect",
          optionsFrom: "consumer",
          optionsFromParams: { context: { from: "derived.catalog.remote.context" } },
        },
        { name: "producer", type: "singleSelect", optionsFrom: "catalog.remote" },
      ],
      {},
      { "derived.catalog.remote.context": "forged" },
      makePort({
        ui: new ScriptedUI({ prefix: "one", consumer: "two", producer: "one" }),
        providers: { catalog: prefix, "catalog.remote": producer, consumer },
      })
    );
    assert.equal(result._unsafeUnwrapErr().name, INPUT_FORWARD_DERIVED_REFERENCE);
    assert.equal(producer.fetchCount, 0);
    assert.equal(consumer.fetchCount, 0);
  });

  it("INPUT-37: a declared output missing from a resumed snapshot is unavailable", async () => {
    const questions: QuestionSpec[] = [
      { name: "producer", type: "singleSelect", optionsFrom: "catalog" },
      {
        name: "consumer",
        type: "singleSelect",
        optionsFrom: "consumer",
        optionsFromParams: { context: { from: "derived.catalog.context" } },
      },
    ];
    const producer = new FakeProvider(
      { options: [{ id: "one" }], derived: { context: "accepted" } },
      ["context"]
    );
    const consumer = new FakeProvider({ options: [{ id: "two" }] });
    const port = makePort({
      ui: new ScriptedUI({ producer: "one", consumer: "two" }),
      providers: { catalog: producer, consumer },
    });
    const first = (await walkInputs(questions, {}, {}, port))._unsafeUnwrap();
    if (first.kind !== "done") assert.fail("expected completed walk");
    const history = first.history.map((entry) => {
      const answers = { ...entry.answers };
      delete answers["derived.catalog.context"];
      return { ...entry, answers };
    });
    const resumed = await walkInputs(
      questions,
      { properties: { "derived.catalog.context": {} } },
      {},
      port,
      { resume: { history } }
    );
    assert.equal(resumed._unsafeUnwrapErr().name, INPUT_FORWARD_DERIVED_REFERENCE);
    assert.equal(consumer.fetchCount, 1);
  });

  it("INPUT-38: Back snapshots precede a provider's own derived merge and retain the per-walk cache", async () => {
    const provider = new FakeProvider(
      { options: [{ id: "one" }], derived: { context: "accepted" } },
      ["context"]
    );
    const observed: (string | string[] | undefined)[] = [];
    const ui = new SequencedPromptUI([
      { kind: "value", value: "one" },
      { kind: "back" },
      { kind: "value", value: "one" },
      { kind: "value", value: "done" },
    ]);
    const result = await walkInputs(
      [
        { name: "producer", type: "singleSelect", optionsFrom: "catalog", validation: "observe" },
        { name: "finish", type: "text" },
      ],
      {},
      {},
      makePort({
        ui,
        providers: { catalog: provider },
        validators: {
          observe: (_value, answers) => {
            observed.push(answers["derived.catalog.context"]);
            return undefined;
          },
        },
      })
    );
    const outcome = result._unsafeUnwrap();
    if (outcome.kind !== "done") assert.fail("expected completed walk");
    assert.deepEqual(observed, [undefined, undefined]);
    assert.notProperty(outcome.history[0].answers, "derived.catalog.context");
    assert.deepEqual(outcome.history[0].resolvedProviders, []);
    assert.deepEqual(outcome.history[1].resolvedProviders, ["catalog"]);
    assert.equal(outcome.answers["derived.catalog.context"], "accepted");
    assert.equal(provider.fetchCount, 1);
  });

  it("INPUT-38: Back recomputes changed params but reuses identical params within one walk", async () => {
    const fetched: string[] = [];
    const provider: OptionsProvider = {
      derivedSchema: ["context"],
      fetch: (params) => {
        fetched.push(params.source);
        return { options: [{ id: "one" }], derived: { context: params.source } };
      },
    };
    const ui = new SequencedPromptUI([
      { kind: "value", value: "old" },
      { kind: "skip", value: "one" },
      { kind: "back" },
      { kind: "value", value: "old" },
      { kind: "skip", value: "one" },
      { kind: "back" },
      { kind: "value", value: "new" },
      { kind: "skip", value: "one" },
      { kind: "value", value: "done" },
    ]);
    const result = await walkInputs(
      [
        { name: "source", type: "text" },
        {
          name: "producer",
          type: "singleSelect",
          optionsFrom: "catalog.remote",
          optionsFromParams: { source: { from: "source" } },
        },
        { name: "finish", type: "text" },
      ],
      {},
      {},
      makePort({ ui, providers: { "catalog.remote": provider } })
    );
    const outcome = result._unsafeUnwrap();
    if (outcome.kind !== "done") assert.fail("expected completed walk");
    assert.deepEqual(fetched, ["old", "new"]);
    assert.equal(outcome.answers["derived.catalog.remote.context"], "new");
    assert.notProperty(outcome.history[0].answers, "derived.catalog.remote.context");
    assert.deepEqual(outcome.history[0].resolvedProviders, []);
    assert.deepEqual(outcome.history[1].resolvedProviders, ["catalog.remote"]);
    assert.equal(outcome.promptCount, 2);
  });

  for (const consumerEnabled of [false, true]) {
    it(`INPUT-38: Back conditionally omits a former producer (consumer enabled: ${consumerEnabled})`, async () => {
      const producer = new FakeProvider(
        { options: [{ id: "one" }], derived: { context: "accepted" } },
        ["context"]
      );
      const consumer = new FakeProvider({ options: [{ id: "two" }] });
      const ui = new SequencedPromptUI([
        { kind: "value", value: "on" },
        { kind: "skip", value: "one" },
        { kind: "back" },
        { kind: "value", value: "off" },
        { kind: "value", value: "done" },
      ]);
      const result = await walkInputs(
        [
          { name: "mode", type: "text" },
          {
            name: "producer",
            type: "singleSelect",
            optionsFrom: "catalog",
            condition: { equals: { mode: "on" } },
          },
          {
            name: "consumer",
            type: "singleSelect",
            condition: { expr: "mode == 'off' && featureFlag('CONSUME')" },
            optionsFrom: "consumer",
            optionsFromParams: { context: { from: "derived.catalog.context" } },
          },
          { name: "finish", type: "text" },
        ],
        {},
        {},
        makePort({
          ui,
          providers: { catalog: producer, consumer },
          exprPort: new ExprPort({ CONSUME: consumerEnabled }),
        })
      );
      if (consumerEnabled) {
        assert.equal(result._unsafeUnwrapErr().name, INPUT_FORWARD_DERIVED_REFERENCE);
      } else {
        const outcome = result._unsafeUnwrap();
        if (outcome.kind !== "done") assert.fail("expected completed walk");
        assert.deepEqual(outcome.answers, { mode: "off", finish: "done" });
        assert.notProperty(outcome.history[1].answers, "derived.catalog.context");
        assert.deepEqual(outcome.history[1].resolvedProviders, []);
      }
      assert.equal(producer.fetchCount, 1);
      assert.equal(consumer.fetchCount, 0);
      assert.deepEqual(
        ui.calls.map((call) => call.name),
        consumerEnabled
          ? ["mode", "producer", "finish", "mode"]
          : ["mode", "producer", "finish", "mode", "finish"]
      );
    });
  }

  it("INPUT-38/INPUT-39: resume retains earlier outputs, then Back recomputes from changed params", async () => {
    const producerParams: string[] = [];
    const consumerParams: string[] = [];
    const producer: OptionsProvider = {
      derivedSchema: ["context"],
      fetch: (params) => {
        producerParams.push(params.source);
        return { options: [{ id: "one" }], derived: { context: params.source } };
      },
    };
    const consumer: OptionsProvider = {
      derivedSchema: ["result"],
      fetch: (params) => {
        consumerParams.push(params.context);
        return { options: [{ id: "two" }], derived: { result: params.context } };
      },
    };
    const questions: QuestionSpec[] = [
      { name: "source", type: "text" },
      {
        name: "producer",
        type: "singleSelect",
        optionsFrom: "catalog.remote",
        optionsFromParams: { source: { from: "source" } },
      },
      {
        name: "consumer",
        type: "singleSelect",
        optionsFrom: "consumer",
        optionsFromParams: { context: { from: "derived.catalog.remote.context" } },
      },
    ];
    const providers = { "catalog.remote": producer, consumer };
    const first = (
      await walkInputs(
        questions,
        {},
        {},
        makePort({
          ui: new ScriptedUI({ source: "old", producer: "one", consumer: "two" }),
          providers,
        })
      )
    )._unsafeUnwrap();
    if (first.kind !== "done") assert.fail("expected completed walk");
    assert.deepEqual(
      first.history.map((entry) => entry.resolvedProviders),
      [[], [], ["catalog.remote"]]
    );
    const retained = (
      await walkInputs(
        questions,
        {},
        {},
        makePort({
          ui: new ScriptedUI({ consumer: "two" }),
          providers,
        }),
        { resume: { history: first.history } }
      )
    )._unsafeUnwrap();
    if (retained.kind !== "done") assert.fail("expected completed resume");
    assert.equal(retained.answers["derived.consumer.result"], "old");
    assert.deepEqual(
      retained.history.map((entry) => entry.resolvedProviders),
      [[], [], ["catalog.remote"]]
    );
    assert.deepEqual(producerParams, ["old"]);
    assert.deepEqual(consumerParams, ["old", "old"]);
    const ui = new SequencedPromptUI([
      { kind: "back" },
      { kind: "back" },
      { kind: "value", value: "new" },
      { kind: "value", value: "one" },
      { kind: "value", value: "two" },
    ]);
    const resumed = (
      await walkInputs(questions, {}, {}, makePort({ ui, providers }), {
        resume: { history: retained.history },
      })
    )._unsafeUnwrap();
    if (resumed.kind !== "done") assert.fail("expected completed resume");
    assert.equal(resumed.answers["derived.catalog.remote.context"], "new");
    assert.equal(resumed.answers["derived.consumer.result"], "new");
    assert.notProperty(resumed.history[1].answers, "derived.catalog.remote.context");
    assert.notProperty(resumed.history[2].answers, "derived.consumer.result");
    assert.deepEqual(
      resumed.history.map((entry) => entry.resolvedProviders),
      [[], [], ["catalog.remote"]]
    );
    assert.equal(retained.answers["derived.consumer.result"], "old");
    assert.deepEqual(
      retained.history.map((entry) => entry.resolvedProviders),
      [[], [], ["catalog.remote"]]
    );
    assert.deepEqual(producerParams, ["old", "new"]);
    assert.deepEqual(consumerParams, ["old", "old", "new"]);
    assert.deepEqual(
      ui.calls.map((call) => call.name),
      ["consumer", "producer", "source", "producer", "consumer"]
    );
  });

  for (const type of ["singleSelect", "multiSelect"] satisfies QuestionSpec["type"][]) {
    for (const mode of ["prefill", "default", "prompt"]) {
      it(`INPUT-35: ${type} ${mode} shares derived merging and membership validation`, async () => {
        const value = type === "multiSelect" ? ["one"] : "one";
        const question: QuestionSpec = { name: "choice", type, optionsFrom: "synthetic" };
        if (mode === "default") question.default = value;
        const provider = new FakeProvider(
          { options: [{ id: "one" }], derived: { context: "data" } },
          ["context"]
        );
        const ui = new ScriptedUI({ choice: "one" }, { choice: ["one"] });
        const result = await walkInputs(
          [question],
          {},
          mode === "prefill"
            ? { choice: value }
            : mode === "default"
              ? { nonInteractive: "true" }
              : {},
          makePort({ ui, providers: { synthetic: provider } })
        );
        const outcome = result._unsafeUnwrap();
        if (outcome.kind !== "done") assert.fail("expected completed walk");
        assert.deepEqual(outcome.answers.choice, value);
        assert.equal(outcome.answers["derived.synthetic.context"], "data");
        assert.equal(provider.fetchCount, 1);
        assert.lengthOf(outcome.history, mode === "prompt" ? 1 : 0);
        provider.fetchCount = 0;
        const invalid = type === "multiSelect" ? ["missing"] : "missing";
        const rejected = await collectInputs(
          [{ ...question, default: invalid }],
          {},
          mode === "prefill"
            ? { choice: invalid }
            : mode === "default"
              ? { nonInteractive: "true" }
              : {},
          makePort({
            ui: new ScriptedUI({ choice: "missing" }, { choice: ["missing"] }),
            providers: { synthetic: provider },
          })
        );
        assert.equal(rejected._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
      });
    }
  }

  it("INPUT-36: static singleton multi-select preserves array shape and no history", async () => {
    const result = await walkInputs(
      [
        {
          name: "choice",
          type: "multiSelect",
          staticOptions: [{ id: "one" }],
          skipSingleOption: true,
        },
      ],
      {},
      {},
      makePort({ ui: new ScriptedUI({}) })
    );
    const outcome = result._unsafeUnwrap();
    if (outcome.kind !== "done") assert.fail("expected completed walk");
    assert.deepEqual(outcome.answers.choice, ["one"]);
    assert.isEmpty(outcome.history);
  });

  it("INPUT-36: static singleton auto-selection does not bypass scalar validation", async () => {
    const result = await collectInputs(
      [
        {
          name: "choice",
          type: "singleSelect",
          staticOptions: [{ id: "one" }],
          skipSingleOption: true,
          validation: "reject",
        },
      ],
      {},
      {},
      makePort({ ui: new ScriptedUI({}), validators: { reject: () => "invalid" } })
    );
    assert.equal(result._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
  });

  it("INPUT-01: a question whose condition is false is skipped whole", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "mcpServerType",
        type: "singleSelect",
        staticOptions: [{ id: "local" }, { id: "remote" }],
      },
      { name: "mcpServerUrl", type: "text", condition: { expr: "mcpServerType == 'remote'" } },
    ];
    const ui = new ScriptedUI({ mcpServerType: "local" });
    const res = await collectInputs(
      questions,
      { properties: { mcpServerType: {}, mcpServerUrl: {} } },
      {},
      makePort({ ui })
    );
    assert.isTrue(res.isOk());
    assert.strictEqual(res._unsafeUnwrap().mcpServerType, "local");
    assert.notProperty(res._unsafeUnwrap(), "mcpServerUrl");
    assert.notInclude(ui.asked, "mcpServerUrl");
  });

  it("INPUT-02: an option-level condition hides only that option, not the question", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "authType",
        type: "singleSelect",
        staticOptions: [
          { id: "oauth" },
          {
            id: "oauth-dynamic",
            condition: {
              expr: "featureFlag('TEAMSFX_MCP_FOR_DA_DT') && featureFlag('TEAMSFX_MCP_FOR_DA_DCR')",
            },
          },
          { id: "entra-sso" },
          { id: "none" },
        ],
      },
    ];
    const ui = new ScriptedUI({ authType: "none" });
    // flags off → the oauth-dynamic option is hidden, the question is still asked
    const res = await collectInputs(
      questions,
      { properties: { authType: {} } },
      {},
      makePort({ ui, exprPort: new ExprPort({}) })
    );
    assert.isTrue(res.isOk());
    assert.strictEqual(res._unsafeUnwrap().authType, "none");
    assert.include(ui.asked, "authType");
    assert.deepStrictEqual(
      (ui.lastOptions.authType ?? []).map((o) => o.id),
      ["oauth", "entra-sso", "none"]
    );
  });

  it("INPUT-03: a question declaring both staticOptions and optionsFrom is rejected", async () => {
    const questions: QuestionSpec[] = [
      { name: "x", type: "singleSelect", staticOptions: [{ id: "a" }], optionsFrom: "p" },
    ];
    const res = await collectInputs(
      questions,
      { properties: { x: {} } },
      {},
      makePort({ ui: new ScriptedUI({}) })
    );
    assert.isTrue(res.isErr());
    const e = res._unsafeUnwrapErr();
    assert.instanceOf(e, SystemError);
    assert.strictEqual(e.name, INPUT_BOTH_OPTION_SOURCES);
  });

  it("INPUT-04: skipSingleOption auto-selects a sole provider option through the UI", async () => {
    const provider = new FakeProvider({ options: [{ id: "remote" }] });
    const questions: QuestionSpec[] = [
      {
        name: "mcpServerType",
        type: "singleSelect",
        optionsFrom: "mcp.serverTypes",
        skipSingleOption: true,
      },
    ];
    const ui = new ScriptedUI({});
    const res = await collectInputs(
      questions,
      { properties: { mcpServerType: {} } },
      {},
      makePort({ ui, providers: { "mcp.serverTypes": provider } })
    );
    assert.strictEqual(res._unsafeUnwrap().mcpServerType, "remote");
    assert.include(ui.asked, "mcpServerType");
  });

  it("INPUT-05: optionsFrom invokes the named provider through the port", async () => {
    const provider = new FakeProvider({ options: [{ id: "local" }, { id: "remote" }] });
    const questions: QuestionSpec[] = [
      { name: "mcpServerType", type: "singleSelect", optionsFrom: "mcp.serverTypes" },
    ];
    const ui = new ScriptedUI({ mcpServerType: "remote" });
    const res = await collectInputs(
      questions,
      { properties: { mcpServerType: {} } },
      {},
      makePort({ ui, providers: { "mcp.serverTypes": provider } })
    );
    assert.strictEqual(provider.fetchCount, 1);
    assert.strictEqual(res._unsafeUnwrap().mcpServerType, "remote");
    assert.deepStrictEqual(
      (ui.lastOptions.mcpServerType ?? []).map((o) => o.id),
      ["local", "remote"]
    );
  });

  it("INPUT-23: a prefilled provider answer still resolves derived values without prompting", async () => {
    const provider = new FakeProvider(
      {
        options: [{ id: "local" }, { id: "remote" }],
        derived: { catalog: '{"github":{"command":"gh","args":["mcp"]}}' },
      },
      ["catalog"]
    );
    const questions: QuestionSpec[] = [
      { name: "mcpServerType", type: "singleSelect", optionsFrom: "mcp.serverTypes" },
    ];
    const ui = new ScriptedUI({});

    const res = await collectInputs(
      questions,
      { properties: { mcpServerType: {} } },
      { mcpServerType: "local" },
      makePort({ ui, providers: { "mcp.serverTypes": provider } })
    );

    assert.deepStrictEqual(res._unsafeUnwrap(), {
      mcpServerType: "local",
      "derived.mcp.serverTypes.catalog": '{"github":{"command":"gh","args":["mcp"]}}',
    });
    assert.strictEqual(provider.fetchCount, 1);
    assert.deepStrictEqual(ui.asked, []);
  });

  it("INPUT-26: a prefilled singleSelect must name a visible static option", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "authType",
        type: "singleSelect",
        staticOptions: [{ id: "none" }, { id: "oauth" }, { id: "entra-sso" }],
      },
    ];
    const ui = new ScriptedUI({});

    const res = await collectInputs(
      questions,
      { properties: { authType: { enum: ["none", "oauth", "entra-sso"] } } },
      { authType: "bogus" },
      makePort({ ui })
    );

    assert.isTrue(res.isErr());
    assert.instanceOf(res._unsafeUnwrapErr(), UserError);
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "authType");
    assert.deepStrictEqual(ui.asked, []);
  });

  it("INPUT-27: a prefilled provider multiSelect rejects an unavailable option", async () => {
    const provider = new FakeProvider({ options: [{ id: "alpha" }, { id: "beta" }] });
    const questions: QuestionSpec[] = [
      {
        name: "selectedLocalServers",
        type: "multiSelect",
        optionsFrom: "mcp.localServers",
      },
    ];

    const res = await collectInputs(
      questions,
      { properties: { selectedLocalServers: { type: "array" } } },
      { selectedLocalServers: ["alpha", "missing"] },
      makePort({
        ui: new ScriptedUI({}),
        providers: { "mcp.localServers": provider },
      })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "selectedLocalServers");
    assert.strictEqual(provider.fetchCount, 1);
  });

  it("INPUT-28: a non-interactive default must name a visible option", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "authType",
        type: "singleSelect",
        staticOptions: [{ id: "none" }, { id: "oauth" }],
        default: "bogus",
      },
    ];

    const res = await collectInputs(
      questions,
      { properties: { authType: { enum: ["none", "oauth"] } } },
      { nonInteractive: "true" },
      makePort({ ui: new ScriptedUI({}) })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "authType");
  });

  it("INPUT-29: a prefilled answer still runs its named validator", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "mcpServerUrl",
        type: "text",
        condition: { expr: "mcpServerUrl == null" },
        validation: "uri",
      },
    ];

    const res = await collectInputs(
      questions,
      { properties: { mcpServerUrl: { type: "string", format: "uri" } } },
      { mcpServerUrl: "not a uri" },
      makePort({ ui: new ScriptedUI({}), validators: { uri: uriValidator } })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "mcpServerUrl");
  });

  it("INPUT-29: a prefilled duplicate-name answer runs the active branch validator", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "apiSpecLocation",
        type: "text",
        condition: { expr: "openApiSpecType == 'enter-url'" },
        validation: "uri",
      },
      {
        name: "apiSpecLocation",
        type: "singleFile",
        condition: { expr: "openApiSpecType == 'open-file'" },
      },
    ];

    const res = await collectInputs(
      questions,
      { properties: { openApiSpecType: {}, apiSpecLocation: {} } },
      { openApiSpecType: "enter-url", apiSpecLocation: "not a uri" },
      makePort({ ui: new ScriptedUI({}), validators: { uri: uriValidator } })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "apiSpecLocation");
  });

  it("INPUT-29: an inactive duplicate URL branch does not reject a prefilled file path", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "apiSpecLocation",
        type: "text",
        condition: { expr: "openApiSpecType == 'enter-url'" },
        validation: "uri",
      },
      {
        name: "apiSpecLocation",
        type: "singleFile",
        condition: { expr: "openApiSpecType == 'open-file'" },
      },
    ];

    const res = await collectInputs(
      questions,
      { properties: { openApiSpecType: {}, apiSpecLocation: {} } },
      { apiSpecLocation: "specs/openapi.yaml" },
      makePort({ ui: new ScriptedUI({}), validators: { uri: uriValidator } })
    );

    assert.isTrue(res.isOk(), res.isErr() ? res.error.message : "expected ok");
    assert.strictEqual(res._unsafeUnwrap().apiSpecLocation, "specs/openapi.yaml");
  });

  it("INPUT-30: a non-interactive default still runs its named validator", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "mcpServerUrl",
        type: "text",
        default: "not a uri",
        validation: "uri",
      },
    ];

    const res = await collectInputs(
      questions,
      { properties: { mcpServerUrl: { type: "string", format: "uri" } } },
      { nonInteractive: "true" },
      makePort({ ui: new ScriptedUI({}), validators: { uri: uriValidator } })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "mcpServerUrl");
  });

  it("INPUT-31: post-prompt validation remains authoritative", async () => {
    const questions: QuestionSpec[] = [{ name: "mcpServerUrl", type: "text", validation: "uri" }];

    const res = await collectInputs(
      questions,
      { properties: { mcpServerUrl: { type: "string", format: "uri" } } },
      {},
      makePort({
        ui: new SequencedPromptUI([{ kind: "value", value: "not a uri" }]),
        validators: { uri: uriValidator },
      })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "mcpServerUrl");
  });

  it("INPUT-32: a prompted singleSelect must name a visible static option", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "authType",
        type: "singleSelect",
        staticOptions: [{ id: "none" }, { id: "oauth" }],
      },
    ];

    const res = await collectInputs(
      questions,
      { properties: { authType: { enum: ["none", "oauth"] } } },
      {},
      makePort({ ui: new SequencedPromptUI([{ kind: "value", value: "bogus" }]) })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "authType");
  });

  it("INPUT-33: a prompted provider multiSelect rejects an unavailable option", async () => {
    const provider = new FakeProvider({ options: [{ id: "alpha" }, { id: "beta" }] });
    const questions: QuestionSpec[] = [
      {
        name: "selectedLocalServers",
        type: "multiSelect",
        optionsFrom: "mcp.localServers",
      },
    ];

    const res = await collectInputs(
      questions,
      { properties: { selectedLocalServers: { type: "array" } } },
      {},
      makePort({
        ui: new SequencedPromptUI([{ kind: "multi", value: ["alpha", "missing"] }]),
        providers: { "mcp.localServers": provider },
      })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "selectedLocalServers");
  });

  it("a prompted static multiSelect rejects an unavailable option", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "selectedServers",
        type: "multiSelect",
        staticOptions: [{ id: "alpha" }, { id: "beta" }],
      },
    ];

    const res = await collectInputs(
      questions,
      { properties: { selectedServers: { type: "array" } } },
      {},
      makePort({ ui: new SequencedPromptUI([{ kind: "multi", value: ["missing"] }]) })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "selectedServers");
  });

  it("INPUT-34: a provider-backed non-interactive default must be available", async () => {
    const provider = new FakeProvider({ options: [{ id: "remote" }] });
    const questions: QuestionSpec[] = [
      {
        name: "mcpServerType",
        type: "singleSelect",
        optionsFrom: "mcp.serverTypes",
        default: "local",
      },
    ];

    const res = await collectInputs(
      questions,
      { properties: { mcpServerType: { enum: ["local", "remote"] } } },
      { nonInteractive: "true" },
      makePort({
        ui: new ScriptedUI({}),
        providers: { "mcp.serverTypes": provider },
      })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "mcpServerType");
    assert.strictEqual(provider.fetchCount, 1);

    const validProvider = new FakeProvider(
      { options: [{ id: "remote" }], derived: { catalog: "available" } },
      ["catalog"]
    );
    questions[0].default = "remote";
    const valid = await collectInputs(
      questions,
      { properties: { mcpServerType: { enum: ["local", "remote"] } } },
      { nonInteractive: "true" },
      makePort({
        ui: new ScriptedUI({}),
        providers: { "mcp.serverTypes": validProvider },
      })
    );

    assert.isTrue(valid.isOk());
    assert.deepStrictEqual(valid._unsafeUnwrap(), {
      nonInteractive: "true",
      mcpServerType: "remote",
      "derived.mcp.serverTypes.catalog": "available",
    });
  });

  it("INPUT-02: a prefilled static option propagates a visibility expression error", async () => {
    const res = await collectInputs(
      [
        {
          name: "authType",
          type: "singleSelect",
          staticOptions: [{ id: "oauth", condition: { expr: "unknownPredicate()" } }],
        },
      ],
      { properties: { authType: {} } },
      { authType: "oauth" },
      makePort({ ui: new ScriptedUI({}) })
    );

    assert.isTrue(res.isErr());
    assert.instanceOf(res._unsafeUnwrapErr(), SystemError);
  });

  it("INPUT-05: a prefilled provider question rejects an unknown provider", async () => {
    const res = await collectInputs(
      [{ name: "server", type: "singleSelect", optionsFrom: "missing.provider" }],
      { properties: { server: {} } },
      { server: "remote" },
      makePort({ ui: new ScriptedUI({}) })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_UNKNOWN_PROVIDER);
  });

  it("INPUT-05: a prefilled provider exception is wrapped as a provider failure", async () => {
    const provider: OptionsProvider = {
      fetch: async () => Promise.reject(new Error("provider exploded")),
    };
    const res = await collectInputs(
      [{ name: "server", type: "singleSelect", optionsFrom: "mcp.servers" }],
      { properties: { server: {} } },
      { server: "remote" },
      makePort({ ui: new ScriptedUI({}), providers: { "mcp.servers": provider } })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_PROVIDER_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "provider exploded");
  });

  it("INPUT-06: a prefilled provider question propagates a parameter expression error", async () => {
    const provider = new FakeProvider({ options: [{ id: "remote" }] });
    const res = await collectInputs(
      [
        {
          name: "server",
          type: "singleSelect",
          optionsFrom: "mcp.servers",
          optionsFromParams: { source: { expr: "unknownSource()" } },
        },
      ],
      { properties: { server: {} } },
      { server: "remote" },
      makePort({ ui: new ScriptedUI({}), providers: { "mcp.servers": provider } })
    );

    assert.isTrue(res.isErr());
    assert.instanceOf(res._unsafeUnwrapErr(), SystemError);
    assert.strictEqual(provider.fetchCount, 0);
  });

  it("INPUT-25: a prefilled provider question rejects undeclared derived data", async () => {
    const provider = new FakeProvider(
      { options: [{ id: "remote" }], derived: { undeclared: "value" } },
      ["catalog"]
    );
    const res = await collectInputs(
      [{ name: "server", type: "singleSelect", optionsFrom: "mcp.servers" }],
      { properties: { server: {} } },
      { server: "remote" },
      makePort({ ui: new ScriptedUI({}), providers: { "mcp.servers": provider } })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_PROVIDER_DERIVED_SCHEMA_VIOLATION);
    assert.include(res._unsafeUnwrapErr().message, "undeclared");
  });

  it("INPUT-25: a non-interactive provider default rejects undeclared derived data", async () => {
    const provider = new FakeProvider(
      { options: [{ id: "remote" }], derived: { undeclared: "value" } },
      ["catalog"]
    );
    const res = await collectInputs(
      [
        {
          name: "server",
          type: "singleSelect",
          optionsFrom: "mcp.servers",
          default: "remote",
        },
      ],
      { properties: { server: {} } },
      { nonInteractive: "true" },
      makePort({ ui: new ScriptedUI({}), providers: { "mcp.servers": provider } })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_PROVIDER_DERIVED_SCHEMA_VIOLATION);
    assert.include(res._unsafeUnwrapErr().message, "undeclared");
  });

  it("INPUT-26: a prefilled singleSelect rejects a list-shaped answer", async () => {
    const res = await collectInputs(
      [
        {
          name: "authType",
          type: "singleSelect",
          staticOptions: [{ id: "oauth" }],
        },
      ],
      { properties: { authType: {} } },
      { authType: ["oauth"] },
      makePort({ ui: new ScriptedUI({}) })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "invalid answer type");
  });

  it("INPUT-29: a prefilled scalar rejects an unknown validator", async () => {
    const res = await collectInputs(
      [{ name: "serverUrl", type: "text", validation: "missing" }],
      { properties: { serverUrl: {} } },
      { serverUrl: "https://example.com/mcp" },
      makePort({ ui: new ScriptedUI({}) })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_UNKNOWN_VALIDATOR);
  });

  it("INPUT-26: option membership is ignored for a non-select question", async () => {
    const res = await collectInputs(
      [{ name: "label", type: "text", staticOptions: [{ id: "unused" }] }],
      { properties: { label: {} } },
      { label: "free text" },
      makePort({ ui: new ScriptedUI({}) })
    );

    assert.deepStrictEqual(res._unsafeUnwrap(), { label: "free text" });
  });

  it("INPUT-02: a prompted static option propagates a visibility expression error", async () => {
    const res = await collectInputs(
      [
        {
          name: "authType",
          type: "singleSelect",
          staticOptions: [{ id: "oauth", condition: { expr: "unknownPredicate()" } }],
        },
      ],
      { properties: { authType: {} } },
      {},
      makePort({ ui: new ScriptedUI({ authType: "oauth" }) })
    );

    assert.isTrue(res.isErr());
    assert.instanceOf(res._unsafeUnwrapErr(), SystemError);
  });

  it("INPUT-06: optionsFromParams close over an answer via the shared evaluator", async () => {
    const provider = new FakeProvider({ options: [{ id: "op1" }] });
    const questions: QuestionSpec[] = [
      { name: "apiSpecLocation", type: "text" },
      {
        name: "apiOperation",
        type: "singleSelect",
        optionsFrom: "openapi.operations",
        optionsFromParams: { specLocation: { from: "apiSpecLocation" } },
        skipSingleOption: true,
      },
    ];
    const ui = new ScriptedUI({ apiSpecLocation: "https://contoso.example/openapi.yaml" });
    const res = await collectInputs(
      questions,
      { properties: { apiSpecLocation: {}, apiOperation: {} } },
      {},
      makePort({ ui, providers: { "openapi.operations": provider } })
    );
    assert.isTrue(res.isOk());
    assert.deepStrictEqual(provider.lastParams, {
      specLocation: "https://contoso.example/openapi.yaml",
    });
  });

  it("INPUT-07: provider derived merges under the reserved derived.<id>.<key> namespace", async () => {
    const provider = new FakeProvider(
      { options: [{ id: "remote" }], derived: { apiAuthData: "bearer" } },
      ["apiAuthData"]
    );
    const questions: QuestionSpec[] = [
      {
        name: "mcpServerType",
        type: "singleSelect",
        optionsFrom: "mcp.serverTypes",
        skipSingleOption: true,
      },
    ];
    const res = await collectInputs(
      questions,
      { properties: { mcpServerType: {} } },
      {},
      makePort({ ui: new ScriptedUI({}), providers: { "mcp.serverTypes": provider } })
    );
    assert.strictEqual(res._unsafeUnwrap()["derived.mcp.serverTypes.apiAuthData"], "bearer");
  });

  it("INPUT-25: provider derived rejects a key absent from derivedSchema", async () => {
    const provider = new FakeProvider(
      { options: [{ id: "remote" }], derived: { undeclared: "value" } },
      ["catalog"]
    );
    const questions: QuestionSpec[] = [
      {
        name: "mcpServerType",
        type: "singleSelect",
        optionsFrom: "mcp.serverTypes",
        skipSingleOption: true,
      },
    ];

    const res = await collectInputs(
      questions,
      { properties: { mcpServerType: {} } },
      {},
      makePort({ ui: new ScriptedUI({}), providers: { "mcp.serverTypes": provider } })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, "InputProviderDerivedSchemaViolation");
    assert.include(res._unsafeUnwrapErr().message, "undeclared");
  });

  it("INPUT-25: provider derived rejects a missing derivedSchema key", async () => {
    const provider = new FakeProvider({ options: [{ id: "remote" }] }, ["catalog"]);
    const questions: QuestionSpec[] = [
      {
        name: "mcpServerType",
        type: "singleSelect",
        optionsFrom: "mcp.serverTypes",
        skipSingleOption: true,
      },
    ];

    const res = await collectInputs(
      questions,
      { properties: { mcpServerType: {} } },
      {},
      makePort({ ui: new ScriptedUI({}), providers: { "mcp.serverTypes": provider } })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, "InputProviderDerivedSchemaViolation");
    assert.include(res._unsafeUnwrapErr().message, "catalog");
  });

  it("INPUT-08: a forward derived.<id>.<key> reference is rejected", async () => {
    const early = new FakeProvider({ options: [{ id: "x" }] });
    const late = new FakeProvider({ options: [{ id: "y" }], derived: { key: "v" } }, ["key"]);
    const questions: QuestionSpec[] = [
      // 'early' reads derived.late.key, but 'late' is declared after → forward reference
      {
        name: "q1",
        type: "singleSelect",
        optionsFrom: "early",
        optionsFromParams: { p: { from: "derived.late.key" } },
        skipSingleOption: true,
      },
      { name: "q2", type: "singleSelect", optionsFrom: "late", skipSingleOption: true },
    ];
    const res = await collectInputs(
      questions,
      { properties: { q1: {}, q2: {} } },
      {},
      makePort({ ui: new ScriptedUI({}), providers: { early, late } })
    );
    assert.isTrue(res.isErr());
    const e = res._unsafeUnwrapErr();
    assert.instanceOf(e, SystemError);
    assert.strictEqual(e.name, INPUT_FORWARD_DERIVED_REFERENCE);
  });

  it("INPUT-09: a provider resolves once per (providerId, params) within a run", async () => {
    const provider = new FakeProvider({ options: [{ id: "remote" }] });
    const questions: QuestionSpec[] = [
      { name: "a", type: "singleSelect", optionsFrom: "mcp.serverTypes", skipSingleOption: true },
      { name: "b", type: "singleSelect", optionsFrom: "mcp.serverTypes", skipSingleOption: true },
    ];
    const res = await collectInputs(
      questions,
      { properties: { a: {}, b: {} } },
      {},
      makePort({ ui: new ScriptedUI({}), providers: { "mcp.serverTypes": provider } })
    );
    assert.isTrue(res.isOk());
    // the second invocation hits the session cache — no re-fetch
    assert.strictEqual(provider.fetchCount, 1);
  });

  it("INPUT-10: a failed validation is a UserError naming the question", async () => {
    const questions: QuestionSpec[] = [{ name: "mcpServerUrl", type: "text", validation: "uri" }];
    const ui = new ScriptedUI({ mcpServerUrl: "not a uri" });
    const res = await collectInputs(
      questions,
      { properties: { mcpServerUrl: {} } },
      {},
      makePort({ ui, validators: { uri: uriValidator } })
    );
    assert.isFunction(ui.lastValidations.mcpServerUrl);
    assert.isTrue(res.isErr());
    const e = res._unsafeUnwrapErr();
    assert.instanceOf(e, UserError);
    assert.strictEqual(e.name, INPUT_VALIDATION_FAILED);
    assert.include(e.message, "mcpServerUrl");
  });

  it("INPUT-10: an unknown question validator is rejected before prompting", async () => {
    const questions: QuestionSpec[] = [
      { name: "mcpServerUrl", type: "text", validation: "missing" },
    ];
    const ui = new ScriptedUI({ mcpServerUrl: "https://api.example.com/mcp" });

    const res = await collectInputs(
      questions,
      { properties: { mcpServerUrl: {} } },
      {},
      makePort({ ui })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_UNKNOWN_VALIDATOR);
    assert.deepStrictEqual(ui.asked, []);
  });

  it("INPUT-10: an unknown input-box validator is rejected before prompting", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "apiSpecLocation",
        type: "singleFileOrText",
        inputOptionItem: { id: "input" },
        inputBoxConfig: { name: "input-api-spec-url", validation: "missing" },
      },
    ];
    const ui = new ScriptedUI({ apiSpecLocation: "https://example.com/openapi.yaml" });

    const res = await collectInputs(
      questions,
      { properties: { apiSpecLocation: {} } },
      {},
      makePort({ ui })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_UNKNOWN_VALIDATOR);
    assert.deepStrictEqual(ui.asked, []);
  });

  it("INPUT-10: scalar provider errors after prompting are returned as input errors", async () => {
    const providerError = new UserError({
      source: "Test",
      name: "ProviderUserError",
      message: "bad",
    });
    const provider: OptionsProvider = { fetch: async () => Promise.reject(providerError) };
    const questions: QuestionSpec[] = [
      { name: "openApiSpec", type: "singleSelect", optionsFrom: "openapi.search" },
    ];

    const res = await collectInputs(
      questions,
      { properties: { openApiSpec: {} } },
      {},
      makePort({
        ui: new SequencedPromptUI([{ kind: "value", value: "https://example.com/openapi.yaml" }]),
        providers: { "openapi.search": provider },
      })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, "ProviderUserError");
  });

  it("INPUT-10: multi provider exceptions after prompting are wrapped as provider failures", async () => {
    const provider: OptionsProvider = {
      fetch: async () => Promise.reject(new Error("provider exploded")),
    };
    const questions: QuestionSpec[] = [
      { name: "apiOperations", type: "multiSelect", optionsFrom: "openapi.operations" },
    ];

    const res = await collectInputs(
      questions,
      { properties: { apiOperations: {} } },
      {},
      makePort({
        ui: new SequencedPromptUI([{ kind: "multi", value: ["GET /repairs"] }]),
        providers: { "openapi.operations": provider },
      })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_PROVIDER_FAILED);
  });

  it("INPUT-10: non-Error provider exceptions are wrapped as provider failures", async () => {
    const provider: OptionsProvider = {
      fetch: async () => Promise.reject("provider exploded"),
    };
    const questions: QuestionSpec[] = [
      { name: "apiOperations", type: "multiSelect", optionsFrom: "openapi.operations" },
    ];

    const res = await collectInputs(
      questions,
      { properties: { apiOperations: {} } },
      {},
      makePort({
        ui: new SequencedPromptUI([{ kind: "multi", value: ["GET /repairs"] }]),
        providers: { "openapi.operations": provider },
      })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_PROVIDER_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "provider exploded");
  });

  it("INPUT-11: machine-state (odr.exe) gating is the provider, never a condition predicate", async () => {
    // odr absent → the provider yields only 'remote'; no condition probes the machine.
    const odrAbsent = new FakeProvider({ options: [{ id: "remote" }] });
    const questions: QuestionSpec[] = [
      {
        name: "mcpServerType",
        type: "singleSelect",
        optionsFrom: "mcp.serverTypes",
        skipSingleOption: true,
      },
    ];
    const res = await collectInputs(
      questions,
      { properties: { mcpServerType: {} } },
      {},
      makePort({ ui: new ScriptedUI({}), providers: { "mcp.serverTypes": odrAbsent } })
    );
    assert.strictEqual(res._unsafeUnwrap().mcpServerType, "remote");
    assert.strictEqual(odrAbsent.fetchCount, 1);
  });

  it("INPUT-12: an entry.params pre-fill skips the question (condition false) and uses the value", async () => {
    // modify add-mcp-server conformance fixture: condition `mcpServerUrl == null`.
    const questions: QuestionSpec[] = [
      {
        name: "mcpServerUrl",
        type: "text",
        condition: { expr: "mcpServerUrl == null" },
        validation: "uri",
      },
    ];
    const url = "https://api.github.com/mcp";
    // pre-filled: the supplied value is used, the question is not prompted
    const uiPre = new ScriptedUI({});
    const resPre = await collectInputs(
      questions,
      { properties: { mcpServerUrl: {} } },
      { mcpServerUrl: url },
      makePort({ ui: uiPre, validators: { uri: uriValidator } })
    );
    assert.isTrue(resPre.isOk());
    assert.strictEqual(resPre._unsafeUnwrap().mcpServerUrl, url);
    assert.notInclude(uiPre.asked, "mcpServerUrl");
    // not pre-filled: the unanswered declared id is null → the question IS asked
    const uiAsk = new ScriptedUI({ mcpServerUrl: url });
    const resAsk = await collectInputs(
      questions,
      { properties: { mcpServerUrl: {} } },
      {},
      makePort({ ui: uiAsk, validators: { uri: uriValidator } })
    );
    assert.isTrue(resAsk.isOk());
    assert.include(uiAsk.asked, "mcpServerUrl");
    assert.strictEqual(resAsk._unsafeUnwrap().mcpServerUrl, url);
  });

  it("INPUT-13: caller-provided language question behaves like any other singleSelect", async () => {
    const questions: QuestionSpec[] = [
      { name: "first", type: "singleSelect", staticOptions: [{ id: "a" }, { id: "b" }] },
      {
        name: "language",
        type: "singleSelect",
        staticOptions: [
          { id: "typescript", label: "TypeScript" },
          { id: "javascript", label: "JavaScript" },
          { id: "python", label: "Python" },
        ],
      },
    ];
    const uiMulti = new ScriptedUI({ first: "a", language: "typescript" });
    const resMulti = await collectInputs(
      questions,
      { properties: { first: {}, language: {} } },
      {},
      makePort({ ui: uiMulti })
    );
    assert.strictEqual(resMulti._unsafeUnwrap().first, "a");
    assert.strictEqual(resMulti._unsafeUnwrap().language, "typescript");
    assert.deepStrictEqual(uiMulti.asked, ["first", "language"]);
    // the language options carry proper-cased display labels (mirroring v3's LanguageOptionMap),
    // not the raw lowercase ids
    assert.deepStrictEqual(uiMulti.lastOptions.language, [
      { id: "typescript", label: "TypeScript" },
      { id: "javascript", label: "JavaScript" },
      { id: "python", label: "Python" },
    ]);

    // A caller whose descriptor languages are ['common'] simply does not add a language question.
    const uiCommon = new ScriptedUI({});
    const resCommon = await collectInputs([], {}, {}, makePort({ ui: uiCommon }));
    assert.notInclude(uiCommon.asked, "language");
    assert.notProperty(resCommon._unsafeUnwrap(), "language");
  });

  it("INPUT-13: a pre-filled language skips the caller-provided language question", async () => {
    const ui = new ScriptedUI({});
    const res = await collectInputs(
      [
        {
          name: "language",
          type: "singleSelect",
          staticOptions: [{ id: "typescript" }, { id: "javascript" }],
        },
      ],
      { properties: { language: {} } },
      { language: "javascript" },
      makePort({ ui })
    );

    assert.isTrue(res.isOk());
    assert.strictEqual(res._unsafeUnwrap().language, "javascript");
    assert.notInclude(ui.asked, "language");
  });

  it("INPUT-14: identical inputs collect identical answers", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "mcpServerType",
        type: "singleSelect",
        optionsFrom: "mcp.serverTypes",
        skipSingleOption: true,
      },
      { name: "authType", type: "singleSelect", staticOptions: [{ id: "none" }, { id: "oauth" }] },
    ];
    const build = (provider: OptionsProvider): CollectInputsPort =>
      makePort({
        ui: new ScriptedUI({ authType: "none" }),
        providers: { "mcp.serverTypes": provider },
      });
    const a = await collectInputs(
      questions,
      { properties: { mcpServerType: {}, authType: {} } },
      {},
      build(new FakeProvider({ options: [{ id: "remote" }] }))
    );
    const b = await collectInputs(
      questions,
      { properties: { mcpServerType: {}, authType: {} } },
      {},
      build(new FakeProvider({ options: [{ id: "remote" }] }))
    );
    assert.deepStrictEqual(a._unsafeUnwrap(), b._unsafeUnwrap());
  });

  it("INPUT-15: a multiSelect question records the selected ids as a string[]", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "selectedLocalServers",
        type: "multiSelect",
        staticOptions: [{ id: "alpha" }, { id: "beta" }, { id: "gamma" }],
      },
    ];
    const ui = new ScriptedUI({}, { selectedLocalServers: ["alpha", "gamma"] });
    const res = await collectInputs(
      questions,
      { properties: { selectedLocalServers: {} } },
      {},
      makePort({ ui })
    );
    assert.isTrue(res.isOk());
    // INV-7: the multi-pick face yields the string[] of selected ids, order-preserving
    assert.deepStrictEqual(res._unsafeUnwrap().selectedLocalServers, ["alpha", "gamma"]);
    assert.include(ui.asked, "selectedLocalServers");
  });

  it("INPUT-16: a back re-asks the previous prompted question, discarding the stale answer", async () => {
    const questions: QuestionSpec[] = [
      { name: "first", type: "singleSelect", staticOptions: [{ id: "a" }, { id: "b" }] },
      { name: "second", type: "singleSelect", staticOptions: [{ id: "x" }, { id: "y" }] },
    ];
    // first→a, second→back (re-asks first), first→b, second→x
    const ui = new SequencedPromptUI([
      { kind: "value", value: "a" },
      { kind: "back" },
      { kind: "value", value: "b" },
      { kind: "value", value: "x" },
    ]);
    const res = await collectInputs(
      questions,
      { properties: { first: {}, second: {} } },
      {},
      makePort({ ui })
    );
    assert.isTrue(res.isOk());
    // the stale first=a is discarded; the re-picked first=b wins
    assert.deepStrictEqual(res._unsafeUnwrap(), { first: "b", second: "x" });
    // back re-asks first, so the call order is first, second, first, second
    assert.deepStrictEqual(
      ui.calls.map((c) => c.name),
      ["first", "second", "first", "second"]
    );
    // the first prompt is step 1 (no Back button); the second is step 2
    assert.deepStrictEqual(
      ui.calls.map((c) => c.step),
      [1, 2, 1, 2]
    );
  });

  it("INPUT-17: a back from a caller-provided language question crosses into the previous question", async () => {
    const questions: QuestionSpec[] = [
      { name: "first", type: "singleSelect", staticOptions: [{ id: "a" }, { id: "b" }] },
      {
        name: "language",
        type: "singleSelect",
        staticOptions: [{ id: "typescript" }, { id: "javascript" }],
      },
    ];
    // first→a, language→back (re-asks Q2), first→b, language→javascript
    const ui = new SequencedPromptUI([
      { kind: "value", value: "a" },
      { kind: "back" },
      { kind: "value", value: "b" },
      { kind: "value", value: "javascript" },
    ]);
    const res = await collectInputs(
      questions,
      { properties: { first: {}, language: {} } },
      {},
      makePort({ ui })
    );
    assert.isTrue(res.isOk());
    // the stale first=a is discarded; the re-picked first=b wins
    assert.deepStrictEqual(res._unsafeUnwrap(), { first: "b", language: "javascript" });
    assert.deepStrictEqual(
      ui.calls.map((c) => c.name),
      ["first", "language", "first", "language"]
    );
    assert.deepStrictEqual(
      ui.calls.map((c) => c.step),
      [1, 2, 1, 2]
    );
  });

  it("INPUT-18: a back at the very first prompt cancels the walk", async () => {
    const questions: QuestionSpec[] = [
      { name: "first", type: "singleSelect", staticOptions: [{ id: "a" }, { id: "b" }] },
    ];
    const ui = new SequencedPromptUI([{ kind: "back" }]);
    const res = await collectInputs(questions, { properties: { first: {} } }, {}, makePort({ ui }));
    assert.isTrue(res.isErr());
    const e = res._unsafeUnwrapErr();
    assert.instanceOf(e, UserError);
    assert.strictEqual(e.name, INPUT_WALK_CANCELLED);
    // only the first prompt was shown, at step 1 (so the host showed no Back button)
    assert.deepStrictEqual(ui.calls, [{ name: "first", step: 1 }]);
  });

  it("INPUT-19: a back at a multiSelect re-asks the previous question; the multi value is discarded", async () => {
    const questions: QuestionSpec[] = [
      { name: "first", type: "singleSelect", staticOptions: [{ id: "a" }, { id: "b" }] },
      {
        name: "servers",
        type: "multiSelect",
        staticOptions: [{ id: "x" }, { id: "y" }, { id: "z" }],
      },
    ];
    // first→a, servers→back (re-asks first), first→b, servers→[x,z]
    const ui = new SequencedPromptUI([
      { kind: "value", value: "a" },
      { kind: "back" },
      { kind: "value", value: "b" },
      { kind: "multi", value: ["x", "z"] },
    ]);
    const res = await collectInputs(
      questions,
      { properties: { first: {}, servers: {} } },
      {},
      makePort({ ui })
    );
    assert.isTrue(res.isOk());
    assert.deepStrictEqual(res._unsafeUnwrap(), { first: "b", servers: ["x", "z"] });
    assert.deepStrictEqual(
      ui.calls.map((c) => c.name),
      ["first", "servers", "first", "servers"]
    );
  });

  it("INPUT-20: a baseStep offset continues the step numbering so the first prompt shows a Back button", async () => {
    const questions: QuestionSpec[] = [
      { name: "first", type: "singleSelect", staticOptions: [{ id: "a" }, { id: "b" }] },
      { name: "second", type: "singleSelect", staticOptions: [{ id: "x" }, { id: "y" }] },
    ];
    const ui = new SequencedPromptUI([
      { kind: "value", value: "a" },
      { kind: "value", value: "x" },
    ]);
    const res = await walkInputs(
      questions,
      { properties: { first: {}, second: {} } },
      {},
      makePort({ ui }),
      { baseStep: 3 }
    );
    assert.isTrue(res.isOk());
    const outcome = res._unsafeUnwrap();
    assert.strictEqual(outcome.kind, "done");
    if (outcome.kind === "done") {
      assert.deepStrictEqual(outcome.answers, { first: "a", second: "x" });
      assert.strictEqual(outcome.promptCount, 2);
    }
    // baseStep 3 → first prompt is step 4 (Back button shown), second is step 5
    assert.deepStrictEqual(
      ui.calls.map((c) => c.step),
      [4, 5]
    );
  });

  it("INPUT-21: with backable set, a back at the first prompt returns a typed back outcome instead of cancelling", async () => {
    const questions: QuestionSpec[] = [
      { name: "first", type: "singleSelect", staticOptions: [{ id: "a" }, { id: "b" }] },
    ];
    const ui = new SequencedPromptUI([{ kind: "back" }]);
    const res = await walkInputs(questions, { properties: { first: {} } }, {}, makePort({ ui }), {
      baseStep: 2,
      backable: true,
    });
    assert.isTrue(res.isOk(), res.isErr() ? `${res.error.name}: ${res.error.message}` : "ok");
    assert.strictEqual(res._unsafeUnwrap().kind, "back");
    // default (backable false) still cancels — INPUT-18 unchanged
    const uiCancel = new SequencedPromptUI([{ kind: "back" }]);
    const cancelled = await walkInputs(
      questions,
      { properties: { first: {} } },
      {},
      makePort({ ui: uiCancel })
    );
    assert.isTrue(cancelled.isErr());
    assert.strictEqual(cancelled._unsafeUnwrapErr().name, INPUT_WALK_CANCELLED);
  });

  it("INPUT-22: resuming a prior walk's history re-asks the last prompted question and preserves back", async () => {
    const questions: QuestionSpec[] = [
      { name: "first", type: "singleSelect", staticOptions: [{ id: "a" }, { id: "b" }] },
      { name: "second", type: "singleSelect", staticOptions: [{ id: "x" }, { id: "y" }] },
    ];
    // First, run to done to capture the walk history.
    const uiFirst = new SequencedPromptUI([
      { kind: "value", value: "a" },
      { kind: "value", value: "x" },
    ]);
    const firstRun = await walkInputs(
      questions,
      { properties: { first: {}, second: {} } },
      {},
      makePort({ ui: uiFirst })
    );
    assert.isTrue(firstRun.isOk());
    const done = firstRun._unsafeUnwrap();
    assert.strictEqual(done.kind, "done");
    if (done.kind !== "done") {
      return;
    }
    // Resume: re-asks 'second' (the last prompted question), then a back crosses into
    // 'first' (the retained history is intact), then re-forward.
    const uiResume = new SequencedPromptUI([
      { kind: "back" },
      { kind: "value", value: "b" },
      { kind: "value", value: "y" },
    ]);
    const resumed = await walkInputs(
      questions,
      { properties: { first: {}, second: {} } },
      {},
      makePort({ ui: uiResume }),
      { resume: { history: done.history } }
    );
    assert.isTrue(resumed.isOk());
    const resumedOutcome = resumed._unsafeUnwrap();
    assert.strictEqual(resumedOutcome.kind, "done");
    if (resumedOutcome.kind === "done") {
      assert.deepStrictEqual(resumedOutcome.answers, { first: "b", second: "y" });
    }
    assert.deepStrictEqual(
      uiResume.calls.map((c) => c.name),
      ["second", "first", "second"]
    );
    // 'second' re-asked at step 2 (history retains 'first'); back crosses to 'first' at step 1
    assert.deepStrictEqual(
      uiResume.calls.map((c) => c.step),
      [2, 1, 2]
    );
  });

  it("INPUT-24: a surface-skipped question is back-transparent (pushes no history)", async () => {
    const questions: QuestionSpec[] = [
      { name: "skipped", type: "singleSelect", staticOptions: [{ id: "auto" }] },
      { name: "second", type: "singleSelect", staticOptions: [{ id: "x" }, { id: "y" }] },
    ];
    // 'skipped' auto-skips (records its answer, but is not a back-stop); a back at 'second'
    // crosses straight over it into an empty history and cancels — it does NOT re-ask 'skipped'.
    const ui = new SequencedPromptUI([{ kind: "skip", value: "auto" }, { kind: "back" }]);
    const res = await collectInputs(
      questions,
      { properties: { skipped: {}, second: {} } },
      {},
      makePort({ ui })
    );
    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_WALK_CANCELLED);
    // both prompts are step 1: 'skipped' pushed no history, so 'second' is also the first step.
    assert.deepStrictEqual(ui.calls, [
      { name: "skipped", step: 1 },
      { name: "second", step: 1 },
    ]);
  });

  it("INPUT-21b: a multiSelect first prompt returns a typed back when backable", async () => {
    const questions: QuestionSpec[] = [
      { name: "servers", type: "multiSelect", staticOptions: [{ id: "x" }, { id: "y" }] },
    ];
    const ui = new SequencedPromptUI([{ kind: "back" }]);
    const res = await walkInputs(questions, { properties: { servers: {} } }, {}, makePort({ ui }), {
      backable: true,
    });
    assert.isTrue(res.isOk());
    assert.strictEqual(res._unsafeUnwrap().kind, "back");
  });

  it("INPUT-22b: resuming an empty history cancels, or hands a typed back when backable", async () => {
    const questions: QuestionSpec[] = [
      { name: "first", type: "singleSelect", staticOptions: [{ id: "a" }] },
    ];
    const cancelled = await walkInputs(
      questions,
      { properties: { first: {} } },
      {},
      makePort({ ui: new SequencedPromptUI([]) }),
      { resume: { history: [] } }
    );
    assert.isTrue(cancelled.isErr());
    assert.strictEqual(cancelled._unsafeUnwrapErr().name, INPUT_WALK_CANCELLED);
    const back = await walkInputs(
      questions,
      { properties: { first: {} } },
      {},
      makePort({ ui: new SequencedPromptUI([]) }),
      { resume: { history: [] }, backable: true }
    );
    assert.isTrue(back.isOk());
    assert.strictEqual(back._unsafeUnwrap().kind, "back");
  });

  it("INPUT-33: a non-interactive multiSelect array default is applied as the answer", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "officeAddinHosts",
        type: "multiSelect",
        staticOptions: [{ id: "word" }, { id: "excel" }, { id: "outlook" }],
        default: ["word", "excel"],
      },
    ];

    const res = await collectInputs(
      questions,
      { properties: { officeAddinHosts: { type: "array" } } },
      { nonInteractive: "true" },
      makePort({ ui: new ScriptedUI({}) })
    );

    assert.isTrue(res.isOk(), res.isErr() ? res.error.message : "expected ok");
    assert.deepStrictEqual(res._unsafeUnwrap().officeAddinHosts, ["word", "excel"]);
  });

  it("INPUT-34: a non-interactive multiSelect array default rejects an unavailable option", async () => {
    const questions: QuestionSpec[] = [
      {
        name: "officeAddinHosts",
        type: "multiSelect",
        staticOptions: [{ id: "word" }, { id: "excel" }],
        default: ["word", "bogus"],
      },
    ];

    const res = await collectInputs(
      questions,
      { properties: { officeAddinHosts: { type: "array" } } },
      { nonInteractive: "true" },
      makePort({ ui: new ScriptedUI({}) })
    );

    assert.isTrue(res.isErr());
    assert.strictEqual(res._unsafeUnwrapErr().name, INPUT_VALIDATION_FAILED);
    assert.include(res._unsafeUnwrapErr().message, "officeAddinHosts");
  });
});
