// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SystemError } from "@microsoft/teamsfx-api";
import {
  STEP_REGISTER_PLUGIN_MANIFEST,
  createDaActionRegisterPluginManifestStep,
  daActionRegisterPluginManifest,
} from "../../../../src/v4/runtime/steps/daAction";
import {
  Pipeline,
  StepContext,
  runScaffoldPipeline,
} from "../../../../src/v4/pipeline/runScaffoldPipeline";
import {
  STEP_REGISTRY,
  buildPipelinePort,
  createStepRegistry,
} from "../../../../src/v4/runtime/runtimeRegistry";
import { assert } from "vitest";
import { err, ok } from "neverthrow";
import { createInMemoryRuntime } from "../../../../src/v4/runtime/inMemoryRuntime";
import {
  DaManifestService,
  daManifestService,
} from "../../../../src/v4/runtime/services/daManifestService";
import { STEP_SET_SENSITIVITY_LABEL } from "../../../../src/v4/runtime/steps/daSensitivity";

/** A minimal in-memory `StepContext` whose read/write share one file map. */
function makeCtx(initial: Record<string, string> = {}): {
  ctx: StepContext;
  files: Map<string, Buffer>;
} {
  const runtime = createInMemoryRuntime();
  for (const [path, body] of Object.entries(initial)) {
    runtime.files.set(path, Buffer.from(body, "utf8"));
  }
  const ctx: StepContext = {
    read: runtime.port.read,
    write: runtime.port.write,
    writeEnvironment: runtime.port.writeEnvironment,
  };
  return { ctx, files: runtime.files };
}

function text(files: Map<string, Buffer>, path: string): string {
  return files.get(path)?.toString("utf8") ?? "";
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isRecordArray(value: unknown): value is Record<string, unknown>[] {
  return Array.isArray(value) && value.every(isRecord);
}

function readJsonObject(files: Map<string, Buffer>, path: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(text(files, path));
  assert.isTrue(isRecord(parsed));
  return parsed;
}

function actions(manifest: Record<string, unknown>): Record<string, unknown>[] {
  const value = manifest.actions;
  assert.isTrue(isRecordArray(value));
  return value;
}

describe("da-action steps (v4)", () => {
  describe(STEP_REGISTER_PLUGIN_MANIFEST, () => {
    it("OWN-01: passes the exact generic execution context and resolved paths to the injected service", async () => {
      const runtime = createInMemoryRuntime();
      runtime.files.set(
        "appPackage/manifest.json",
        Buffer.from(JSON.stringify({ declarativeAgents: [{ file: "declarativeAgent.json" }] }))
      );
      runtime.files.set("appPackage/declarativeAgent.json", Buffer.from("{}"));
      let executionContext: StepContext | undefined;
      let serviceContext: Pick<StepContext, "read" | "write"> | undefined;
      const registrations: [string, string][] = [];
      const registry = createStepRegistry(undefined, {
        registerDeclarativeAgentAction(io, teamsManifestPath, pluginManifestPath) {
          serviceContext = io;
          registrations.push([teamsManifestPath, pluginManifestPath]);
          return ok(undefined);
        },
        setSensitivityLabel: () => ok(undefined),
      });
      const step = registry.get(STEP_REGISTER_PLUGIN_MANIFEST);
      assert.isDefined(step);
      const port = buildPipelinePort(
        runtime.exprPort,
        runtime.port,
        runtime.port.writeEnvironment,
        new Map([
          [
            STEP_REGISTER_PLUGIN_MANIFEST,
            {
              validateParams: step.validateParams,
              apply(params, ctx) {
                executionContext = ctx;
                return step.apply(params, ctx);
              },
            },
          ],
        ])
      );

      const result = await runScaffoldPipeline(
        {
          pipeline: "default",
          steps: [
            {
              step: STEP_REGISTER_PLUGIN_MANIFEST,
              with: {
                teamsManifestPath: "{{packagePath}}/manifest.json",
                pluginManifestPath: "{{packagePath}}/ai-plugin-{{actionId}}.json",
              },
            },
          ],
        },
        [],
        { packagePath: "appPackage", actionId: "github" },
        { path: "/out", existing: [] },
        port
      );

      assert.isTrue(result.isOk());
      assert.deepEqual(registrations, [
        ["appPackage/manifest.json", "appPackage/ai-plugin-github.json"],
      ]);
      assert.strictEqual(serviceContext, executionContext);
      assert.isDefined(executionContext);
      assert.notProperty(executionContext, "manifestWrapper");
      assert.notProperty(port, "manifestWrapper");
    });

    it("is registered in the v4 step registry", () => {
      const step = STEP_REGISTRY.get(STEP_REGISTER_PLUGIN_MANIFEST);
      assert.isDefined(step);
      assert.isUndefined(
        step.validateParams({
          teamsManifestPath: "appPackage/manifest.json",
          pluginManifestPath: "appPackage/ai-plugin.json",
        })
      );
    });

    it("propagates an injected manifest service error unchanged", async () => {
      const failure = err(
        new SystemError({
          source: "Scaffold",
          name: "ManifestMutationUnavailable",
          message: "the current runtime does not provide manifest mutation",
        })
      );
      const step = createDaActionRegisterPluginManifestStep({
        registerDeclarativeAgentAction: () => failure,
        setSensitivityLabel: () => ok(undefined),
      });
      const { ctx } = makeCtx();
      const result = await step.apply(
        {
          teamsManifestPath: "appPackage/manifest.json",
          pluginManifestPath: "appPackage/ai-plugin.json",
        },
        ctx
      );

      assert.strictEqual(result, failure);
    });

    it("validateParams: passes when teamsManifestPath/pluginManifestPath are strings", () => {
      assert.isUndefined(
        daActionRegisterPluginManifest.validateParams({
          teamsManifestPath: "appPackage/manifest.json",
          pluginManifestPath: "appPackage/ai-plugin-apigithubc.json",
        })
      );
    });

    it("AC-12: delegates path-aware mutation to the manifest service without reading JSON", async () => {
      const registrations: [string, string][] = [];
      const service: DaManifestService = {
        registerDeclarativeAgentAction: (io, teamsManifestPath, pluginManifestPath) => {
          assert.strictEqual(io, ctx);
          registrations.push([teamsManifestPath, pluginManifestPath]);
          return ok(undefined);
        },
        setSensitivityLabel: () => ok(undefined),
      };
      const ctx: StepContext = {
        read: () => {
          throw new Error("the step must not parse manifests directly");
        },
        write: () => {
          throw new Error("the step must not write manifests directly");
        },
        writeEnvironment: () => Promise.resolve(ok(undefined)),
      };

      const res = await createDaActionRegisterPluginManifestStep(service).apply(
        {
          teamsManifestPath: "appPackage/manifest.json",
          pluginManifestPath: "appPackage/ai-plugin-apigithubc.json",
        },
        ctx
      );

      assert.isTrue(res.isOk(), res.isErr() ? res.error.message : "expected ok");
      assert.deepEqual(registrations, [
        ["appPackage/manifest.json", "appPackage/ai-plugin-apigithubc.json"],
      ]);
    });

    it("SCN-ADD-MCP-04: derives the DA manifest path and registers the plugin manifest", async () => {
      const { ctx, files } = makeCtx({
        "appPackage/manifest.json": JSON.stringify({
          declarativeAgents: [{ file: "declarativeAgent.json" }],
        }),
        "appPackage/declarativeAgent.json": JSON.stringify({ name: "Agent" }),
      });

      const res = await daActionRegisterPluginManifest.apply(
        {
          teamsManifestPath: "appPackage/manifest.json",
          pluginManifestPath: "appPackage/ai-plugin-apigithubc.json",
        },
        ctx
      );

      assert.isTrue(res.isOk(), res.isErr() ? res.error.message : "expected ok");
      const manifest = readJsonObject(files, "appPackage/declarativeAgent.json");
      assert.deepInclude(actions(manifest), {
        id: "apigithubc",
        file: "ai-plugin-apigithubc.json",
      });
    });

    it("SCN-ADD-MCP-05: upserts by pluginManifestPath so a re-run does not duplicate the action", async () => {
      const { ctx, files } = makeCtx({
        "appPackage/manifest.json": JSON.stringify({
          declarativeAgents: [{ file: "declarativeAgent.json" }],
        }),
        "appPackage/declarativeAgent.json": JSON.stringify({
          actions: [{ id: "apigithubc", file: "ai-plugin-apigithubc.json" }],
        }),
      });

      await daActionRegisterPluginManifest.apply(
        {
          teamsManifestPath: "appPackage/manifest.json",
          pluginManifestPath: "appPackage/ai-plugin-apigithubc.json",
        },
        ctx
      );

      const manifest = readJsonObject(files, "appPackage/declarativeAgent.json");
      assert.lengthOf(actions(manifest), 1);
    });

    it("errors when the Teams manifest does not point at a DA manifest", async () => {
      const { ctx } = makeCtx({ "appPackage/manifest.json": JSON.stringify({}) });
      const res = await daActionRegisterPluginManifest.apply(
        {
          teamsManifestPath: "appPackage/manifest.json",
          pluginManifestPath: "appPackage/ai-plugin-apigithubc.json",
        },
        ctx
      );
      assert.isTrue(res.isErr());
      assert.instanceOf(res._unsafeUnwrapErr(), SystemError);
      assert.strictEqual(res._unsafeUnwrapErr().name, "DaActionManifestFileMissing");
    });

    it.each([
      {
        name: "missing Teams manifest",
        initial: {},
        errorName: "DaActionTeamsManifestMissing",
      },
      {
        name: "invalid Teams manifest",
        initial: { "appPackage/manifest.json": "{" },
        errorName: "DaActionTeamsManifestInvalid",
      },
      {
        name: "missing declarative agent manifest",
        initial: {
          "appPackage/manifest.json": JSON.stringify({
            declarativeAgents: [{ file: "declarativeAgent.json" }],
          }),
        },
        errorName: "DaActionManifestMissing",
      },
      {
        name: "invalid declarative agent manifest",
        initial: {
          "appPackage/manifest.json": JSON.stringify({
            declarativeAgents: [{ file: "declarativeAgent.json" }],
          }),
          "appPackage/declarativeAgent.json": "{",
        },
        errorName: "DaActionManifestInvalid",
      },
    ])("returns a distinct error for $name", async ({ initial, errorName }) => {
      const { ctx } = makeCtx(initial);

      const result = await daActionRegisterPluginManifest.apply(
        {
          teamsManifestPath: "appPackage/manifest.json",
          pluginManifestPath: "appPackage/ai-plugin.json",
        },
        ctx
      );

      assert.isTrue(result.isErr());
      assert.strictEqual(result._unsafeUnwrapErr().name, errorName);
    });

    it("returns a distinct error when the declarative agent manifest cannot be written", () => {
      const runtime = createInMemoryRuntime();
      const port = buildPipelinePort(
        runtime.exprPort,
        {
          writeNew: runtime.port.writeNew,
          read: (filePath): Buffer | undefined => {
            if (filePath === "appPackage/manifest.json") {
              return Buffer.from(
                JSON.stringify({ declarativeAgents: [{ file: "declarativeAgent.json" }] })
              );
            }
            if (filePath === "appPackage/declarativeAgent.json") {
              return Buffer.from(JSON.stringify({ name: "Agent" }));
            }
            return undefined;
          },
          write: (): void => {
            throw new Error("write failed at C:\\secret\\project");
          },
        },
        runtime.port.writeEnvironment
      );

      const result = daManifestService.registerDeclarativeAgentAction(
        port,
        "appPackage/manifest.json",
        "appPackage/ai-plugin.json"
      );

      assert.isTrue(result.isErr());
      assert.strictEqual(result._unsafeUnwrapErr().name, "DaActionManifestWriteFailed");
      assert.notInclude(result._unsafeUnwrapErr().message, "C:\\secret\\project");
    });

    it.each([
      ["appPackage/manifest.json", "DaActionTeamsManifestReadFailed"],
      ["appPackage/declarativeAgent.json", "DaActionManifestReadFailed"],
    ])("preserves the read error for %s", async (failedPath, errorName) => {
      const { ctx } = makeCtx({
        "appPackage/manifest.json": JSON.stringify({
          declarativeAgents: [{ file: "declarativeAgent.json" }],
        }),
      });
      const result = await daActionRegisterPluginManifest.apply(
        {
          teamsManifestPath: "appPackage/manifest.json",
          pluginManifestPath: "appPackage/ai-plugin.json",
        },
        {
          ...ctx,
          read(filePath) {
            if (filePath === failedPath) {
              throw new Error("read failed at C:\\secret\\project");
            }
            return ctx.read(filePath);
          },
        }
      );

      assert.strictEqual(result._unsafeUnwrapErr().name, errorName);
      assert.notInclude(result._unsafeUnwrapErr().message, "C:\\secret\\project");
    });

    it("OWN-02: real DA steps preserve paths and upserts without writes crossing runtimes", async () => {
      const registry = createStepRegistry({ resolveId: async () => "general-label-id" });
      const first = createInMemoryRuntime(undefined, registry);
      const second = createInMemoryRuntime(undefined, registry);
      for (const { runtime, name } of [
        { runtime: first, name: "First" },
        { runtime: second, name: "Second" },
      ]) {
        runtime.files.set(
          "appPackage/manifest.json",
          Buffer.from(
            JSON.stringify({
              declarativeAgents: [{ file: "agents/primary.json" }, { file: "ignored.json" }],
            })
          )
        );
        runtime.files.set("appPackage/agents/primary.json", Buffer.from(JSON.stringify({ name })));
        runtime.files.set("appPackage/ignored.json", Buffer.from("{}"));
      }
      const pipeline: Pipeline = {
        pipeline: "default",
        steps: [
          {
            step: STEP_REGISTER_PLUGIN_MANIFEST,
            with: {
              teamsManifestPath: "appPackage/manifest.json",
              pluginManifestPath: "appPackage/plugins/{{pluginName}}.json",
            },
          },
          {
            step: STEP_SET_SENSITIVITY_LABEL,
            with: { manifestPath: "appPackage/agents/primary.json" },
          },
        ],
      };
      const secondBefore = new Map(second.files);
      const firstResult = await runScaffoldPipeline(
        pipeline,
        [],
        { pluginName: "ai-plugin-first" },
        { path: "/first", existing: [] },
        first.port
      );
      assert.isTrue(firstResult.isOk());
      assert.deepEqual(second.files, secondBefore);
      const firstAfter = new Map(first.files);
      const secondResult = await runScaffoldPipeline(
        pipeline,
        [],
        { pluginName: "lookup" },
        { path: "/second", existing: [] },
        second.port
      );
      assert.isTrue(secondResult.isOk());
      assert.deepEqual(first.files, firstAfter);
      const repeated = await runScaffoldPipeline(
        pipeline,
        [],
        { pluginName: "ai-plugin-first" },
        { path: "/first", existing: [] },
        first.port
      );
      assert.isTrue(repeated.isOk());
      assert.deepEqual(first.files, firstAfter);
      const firstManifest = readJsonObject(first.files, "appPackage/agents/primary.json");
      const secondManifest = readJsonObject(second.files, "appPackage/agents/primary.json");
      assert.deepEqual(actions(firstManifest), [
        { id: "first", file: "../plugins/ai-plugin-first.json" },
      ]);
      assert.deepEqual(actions(secondManifest), [{ id: "lookup", file: "../plugins/lookup.json" }]);
      assert.deepEqual(firstManifest.sensitivity_label, { id: "general-label-id" });
      assert.deepEqual(secondManifest.sensitivity_label, { id: "general-label-id" });
      assert.strictEqual(firstManifest.name, "First");
      assert.strictEqual(secondManifest.name, "Second");
      assert.strictEqual(text(first.files, "appPackage/ignored.json"), "{}");
      assert.strictEqual(text(second.files, "appPackage/ignored.json"), "{}");
    });
  });
});
