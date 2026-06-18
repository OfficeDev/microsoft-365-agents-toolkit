// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import * as fs from "fs";
import * as path from "path";
import { UserError } from "@microsoft/teamsfx-api";
import { TemplateFileEntry } from "../../../src/v4/model/dataModel";
import { REQUIRE_EMPTY_TARGET } from "../../../src/v4/pipeline/runScaffoldPipeline";
import { createInMemoryRuntime } from "../../../src/v4/runtime/inMemoryRuntime";
import { ScaffoldRequest, scaffold } from "../../../src/v4/runtime/scaffold";

/**
 * T3 scenario tier (ADR-0018): the whole `da/api-plugin-from-existing-api`
 * create package — the declarative agent with an action generated from an
 * existing OpenAPI description document — scaffolded under `InMemoryRuntime`.
 *
 * Spec: docs/03-specs/scenarios/da/create-api-plugin-from-existing-api.md
 * (SCN-CREATE-APIPLUGIN-OPENAPI-01..09)
 */

const PKG_DIR = path.resolve(
  __dirname,
  "../../../../../templates/v4/create/da/api-plugin-from-existing-api"
);
const SPEC_PATH = path.resolve(__dirname, "fixtures/repairs-openapi.yaml");
const APIKEY_SPEC_PATH = path.resolve(__dirname, "fixtures/repairs-openapi-apikey.yaml");

const EXPECTED_RENDER_FILES = [
  ".gitignore",
  ".vscode/extensions.json",
  ".vscode/launch.json",
  ".vscode/settings.json",
  ".vscode/tasks.json",
  "README.md",
  "appPackage/color.png",
  "appPackage/declarativeAgent.json",
  "appPackage/instruction.txt",
  "appPackage/manifest.json",
  "appPackage/outline.png",
  "env/.env.dev",
  "env/.env.local",
  "evals/prompts.json",
  "m365agents.local.yml",
  "m365agents.yml",
];

const descriptor: unknown = JSON.parse(
  fs.readFileSync(path.join(PKG_DIR, "descriptor.json"), "utf8")
);
const pipeline: unknown = JSON.parse(
  fs.readFileSync(path.join(PKG_DIR, "pipeline.json"), "utf8")
);

function loadContent(): TemplateFileEntry[] {
  const root = path.join(PKG_DIR, "content");
  const entries: TemplateFileEntry[] = [];
  const walk = (dir: string): void => {
    for (const name of fs.readdirSync(dir)) {
      const full = path.join(dir, name);
      if (fs.statSync(full).isDirectory()) {
        walk(full);
      } else {
        entries.push({
          path: path.relative(root, full).replace(/\\/g, "/"),
          data: fs.readFileSync(full),
        });
      }
    }
  };
  walk(root);
  return entries;
}

const content = loadContent();

async function run(options: { existing?: string[]; specPath?: string } = {}) {
  const runtime = createInMemoryRuntime();
  const request: ScaffoldRequest = {
    descriptor,
    pipeline,
    content,
    answers: { apiSpecLocation: options.specPath ?? SPEC_PATH, apiOperations: ["GET /repairs"] },
    callerFloor: { appName: "MyAgent", language: "common" },
    targetDir: { path: "/out", existing: options.existing ?? [] },
  };
  const result = await scaffold(request, runtime);
  assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
  return { files: runtime.files, outcome: result._unsafeUnwrap() };
}

function text(files: Map<string, Buffer>, filePath: string): string {
  const buf = files.get(filePath);
  assert.isDefined(buf, `expected '${filePath}' to be written`);
  return (buf ?? Buffer.from("")).toString("utf8");
}

function readJson(files: Map<string, Buffer>, filePath: string): any {
  return JSON.parse(text(files, filePath));
}

describe("SCN-DA-CREATE-API-PLUGIN-FROM-EXISTING-API (v4, T3 InMemoryRuntime)", () => {
  it("SCN-CREATE-APIPLUGIN-OPENAPI-01: the render phase writes exactly the common file set", async () => {
    const { outcome } = await run();
    assert.deepStrictEqual([...outcome.written].sort(), [...EXPECTED_RENDER_FILES].sort());
    assert.isEmpty(outcome.skipped);
  });

  it("SCN-CREATE-APIPLUGIN-OPENAPI-02: the OpenAPI post-render step generates the plugin manifest and filtered spec", async () => {
    const { files } = await run();
    assert.isTrue(files.has("appPackage/ai-plugin.json"));
    assert.isTrue(files.has("appPackage/apiSpecificationFile/openapi.yaml"));
    const plugin = readJson(files, "appPackage/ai-plugin.json");
    assert.lengthOf(plugin.runtimes, 1);
    assert.strictEqual(plugin.runtimes[0].type, "OpenApi");
    assert.strictEqual(plugin.runtimes[0].auth.type, "None");
    assert.strictEqual(plugin.runtimes[0].spec.url, "apiSpecificationFile/openapi.yaml");
  });

  it("SCN-CREATE-APIPLUGIN-OPENAPI-03: declarativeAgent.json is updated with the generated action", async () => {
    const { files } = await run();
    const agent = readJson(files, "appPackage/declarativeAgent.json");
    assert.strictEqual(agent.name, "MyAgent");
    assert.deepStrictEqual(agent.actions, [{ id: "action_1", file: "ai-plugin.json" }]);
  });

  it("SCN-CREATE-APIPLUGIN-OPENAPI-09: OpenAPI summaries are propagated to conversation starters", async () => {
    const { files } = await run();
    const agent = readJson(files, "appPackage/declarativeAgent.json");
    assert.deepStrictEqual(agent.conversation_starters, [{ text: "List repairs" }]);
  });

  it("SCN-CREATE-APIPLUGIN-OPENAPI-04: manifest.json preserves the declarative agent wiring and env refs", async () => {
    const { files } = await run();
    const manifest = readJson(files, "appPackage/manifest.json");
    assert.strictEqual(manifest.manifestVersion, "1.28");
    assert.strictEqual(manifest.id, "${{TEAMS_APP_ID}}");
    assert.deepStrictEqual(manifest.copilotAgents.declarativeAgents[0], {
      id: "declarativeAgent",
      file: "declarativeAgent.json",
    });
  });

  it("SCN-CREATE-APIPLUGIN-OPENAPI-05: the only post-render step is openapi/generate-plugin-files after require-empty-target", async () => {
    const { outcome } = await run();
    assert.deepStrictEqual(outcome.stepsRun, ["require-empty-target", "openapi/generate-plugin-files"]);
    assert.isEmpty(outcome.stepsSkipped);
  });

  it("SCN-CREATE-APIPLUGIN-OPENAPI-06: a non-empty target fails require-empty-target first and writes nothing", async () => {
    const runtime = createInMemoryRuntime();
    const request: ScaffoldRequest = {
      descriptor,
      pipeline,
      content,
      answers: { apiSpecLocation: SPEC_PATH, apiOperations: ["GET /repairs"] },
      callerFloor: { appName: "MyAgent", language: "common" },
      targetDir: { path: "/out", existing: ["appPackage/manifest.json"] },
    };
    const result = await scaffold(request, runtime);
    assert.isTrue(result.isErr());
    const error = result._unsafeUnwrapErr();
    assert.instanceOf(error, UserError);
    assert.strictEqual(error.name, REQUIRE_EMPTY_TARGET);
    assert.strictEqual(runtime.files.size, 0);
  });

  it("SCN-CREATE-APIPLUGIN-OPENAPI-07: identical inputs are deterministic", async () => {
    const first = await run();
    const second = await run();
    assert.deepStrictEqual([...first.outcome.written].sort(), [...second.outcome.written].sort());
    assert.strictEqual(
      text(first.files, "appPackage/ai-plugin.json"),
      text(second.files, "appPackage/ai-plugin.json")
    );
  });

  it("SCN-CREATE-APIPLUGIN-OPENAPI-08: selected API-key operations inject API-key registration actions into yml", async () => {
    const { files } = await run({ specPath: APIKEY_SPEC_PATH });
    const yml = text(files, "m365agents.yml");
    const localYml = text(files, "m365agents.local.yml");
    for (const content of [yml, localYml]) {
      assert.include(content, "  - uses: apiKey/register");
      assert.include(content, "      name: ApiKeyAuth");
      assert.include(content, "      apiSpecPath: ./appPackage/apiSpecificationFile/openapi.yaml");
      assert.include(content, "      registrationId:");
    }
  });
});