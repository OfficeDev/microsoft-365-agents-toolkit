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
 * T3 scenario tier (ADR-0018): the whole `da/api-plugin-from-scratch` create
 * package — the declarative agent with a brand-new (no-auth) API plugin action —
 * scaffolded under `InMemoryRuntime`, asserting the vertical contract for each
 * declared language.
 *
 * Spec: docs/03-specs/scenarios/da/create-api-plugin-from-scratch.md
 * (SCN-CREATE-APIPLUGIN-01..09)
 *
 * Each `it("SCN-CREATE-APIPLUGIN-0X")` maps 1:1 to a scenario AC row. This is the
 * first **language-partitioned** v4 package, so the suite also locks the
 * `select-language-content` axis: the same descriptor / pipeline scaffolds the
 * TypeScript or JavaScript subtree off the Q0 `language`, prefix stripped. The
 * action is pre-baked into `repairDeclarativeAgent.json` (the new-API-from-scratch
 * path), so — like basic DA — the only pipeline step is `require-empty-target`.
 * No v3 symbol participates.
 */

const PKG_DIR = path.resolve(
  __dirname,
  "../../../../../templates/v4/create/da/api-plugin-from-scratch"
);

/** The TypeScript backend file set a scaffold writes (`.tpl` stripped, language prefix stripped). */
const EXPECTED_TS_FILES = [
  ".funcignore",
  ".gitignore",
  ".vscode/extensions.json",
  ".vscode/launch.json",
  ".vscode/settings.json",
  ".vscode/tasks.json",
  "README.md",
  "appPackage/adaptiveCards/listRepairs.data.json",
  "appPackage/adaptiveCards/listRepairs.json",
  "appPackage/ai-plugin.json",
  "appPackage/apiSpecificationFile/repair.yml",
  "appPackage/color.png",
  "appPackage/instruction.txt",
  "appPackage/manifest.json",
  "appPackage/outline.png",
  "appPackage/repairDeclarativeAgent.json",
  "env/.env.dev",
  "env/.env.dev.user",
  "env/.env.local",
  "env/.env.local.user",
  "host.json",
  "infra/azure.bicep",
  "infra/azure.parameters.json",
  "local.settings.json",
  "m365agents.local.yml",
  "m365agents.yml",
  "package.json",
  "src/functions/repairs.ts",
  "src/repairsData.json",
  "tsconfig.json",
];

/** The JavaScript backend file set — the TS set minus `tsconfig.json`, with `repairs.js` for `repairs.ts`. */
const EXPECTED_JS_FILES = EXPECTED_TS_FILES.filter(
  (f) => f !== "tsconfig.json" && f !== "src/functions/repairs.ts"
).concat("src/functions/repairs.js");

const descriptor: unknown = JSON.parse(
  fs.readFileSync(path.join(PKG_DIR, "descriptor.json"), "utf8")
);
const pipeline: unknown = JSON.parse(fs.readFileSync(path.join(PKG_DIR, "pipeline.json"), "utf8"));

/** Load the package's `content/**` as the opened-entry list (forward-slash paths, raw bytes). */
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

interface RunOptions {
  existing?: string[];
}

/** Scaffold the package for one language against a fresh in-memory runtime (no answers). */
async function run(
  language: string,
  options: RunOptions = {}
): Promise<{ files: Map<string, Buffer>; outcome: ReturnType<typeof unwrapOutcome> }> {
  const runtime = createInMemoryRuntime();
  const request: ScaffoldRequest = {
    descriptor,
    pipeline,
    content,
    answers: {},
    callerFloor: { appName: "MyAgent", language },
    targetDir: { path: "/out", existing: options.existing ?? [] },
  };
  const result = await scaffold(request, runtime);
  return { files: runtime.files, outcome: unwrapOutcome(result) };
}

/** Narrow a successful scaffold result to its outcome (failing the test otherwise). */
function unwrapOutcome(result: Awaited<ReturnType<typeof scaffold>>) {
  assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
  return result._unsafeUnwrap();
}

function text(files: Map<string, Buffer>, filePath: string): string {
  const buf = files.get(filePath);
  assert.isDefined(buf, `expected '${filePath}' to be written`);
  return (buf ?? Buffer.from("")).toString("utf8");
}

/** Parse a written JSON file (returns `any` — a test navigating rendered output). */
function readJson(files: Map<string, Buffer>, filePath: string): any {
  return JSON.parse(text(files, filePath));
}

describe("SCN-DA-CREATE-API-PLUGIN-FROM-SCRATCH (v4, T3 InMemoryRuntime)", () => {
  it("SCN-CREATE-APIPLUGIN-01: the render phase writes exactly the TypeScript file set (empty target)", async () => {
    const { outcome } = await run("typescript");
    assert.deepStrictEqual([...outcome.written].sort(), [...EXPECTED_TS_FILES].sort());
    assert.isEmpty(outcome.skipped);
  });

  it("SCN-CREATE-APIPLUGIN-02: repairDeclarativeAgent.json wires the pre-baked action and omits the sensitivity label", async () => {
    const { files } = await run("typescript");
    const agent = readJson(files, "appPackage/repairDeclarativeAgent.json");
    assert.strictEqual(agent.name, "MyAgent${{APP_NAME_SUFFIX}}");
    assert.strictEqual(agent.instructions, "$[file('instruction.txt')]");
    assert.lengthOf(agent.actions, 1);
    assert.deepStrictEqual(agent.actions[0], { id: "repairPlugin", file: "ai-plugin.json" });
    // the sensitivity-label feature is off by default, so the section is omitted.
    assert.notProperty(agent, "sensitivity_label");
  });

  it("SCN-CREATE-APIPLUGIN-03: ai-plugin.json is the no-auth OpenApi runtime over the bundled spec", async () => {
    const { files } = await run("typescript");
    const plugin = readJson(files, "appPackage/ai-plugin.json");
    assert.strictEqual(plugin.namespace, "repairs");
    assert.strictEqual(plugin.name_for_human, "MyAgent${{APP_NAME_SUFFIX}}");
    assert.lengthOf(plugin.runtimes, 1);
    const runtime = plugin.runtimes[0];
    assert.strictEqual(runtime.type, "OpenApi");
    assert.strictEqual(runtime.auth.type, "None");
    assert.strictEqual(runtime.spec.url, "apiSpecificationFile/repair.yml");
  });

  it("SCN-CREATE-APIPLUGIN-04: manifest.json wires the single declarative agent and preserves the env refs", async () => {
    const { files } = await run("typescript");
    const manifest = readJson(files, "appPackage/manifest.json");
    assert.strictEqual(manifest.manifestVersion, "1.28");
    // the env-var refs survive render verbatim (provision resolves them later).
    assert.strictEqual(manifest.id, "${{TEAMS_APP_ID}}");
    assert.strictEqual(manifest.name.short, "MyAgent${{APP_NAME_SUFFIX}}");
    const agents = manifest.copilotAgents.declarativeAgents;
    assert.lengthOf(agents, 1);
    assert.deepStrictEqual(agents[0], {
      id: "repairDeclarativeAgent",
      file: "repairDeclarativeAgent.json",
    });
  });

  it("SCN-CREATE-APIPLUGIN-05: language 'typescript' narrows to the typescript subtree, prefix stripped", async () => {
    const { files, outcome } = await run("typescript");
    // every written path is project-root-relative — the language prefix is gone.
    for (const written of outcome.written) {
      assert.isFalse(
        written.startsWith("typescript/") || written.startsWith("javascript/"),
        `written path '${written}' still carries a language prefix`
      );
    }
    assert.isTrue(files.has("src/functions/repairs.ts"));
    assert.isTrue(files.has("tsconfig.json"));
    assert.isFalse(files.has("src/functions/repairs.js"));
    // the derived package.json name proves the SafeProjectNameLowerCase producer.
    assert.strictEqual(readJson(files, "package.json").name, "myagent");
  });

  it("SCN-CREATE-APIPLUGIN-06: language 'javascript' writes the javascript subtree with the same rendered shapes", async () => {
    const { files, outcome } = await run("javascript");
    assert.deepStrictEqual([...outcome.written].sort(), [...EXPECTED_JS_FILES].sort());
    assert.isTrue(files.has("src/functions/repairs.js"));
    assert.isFalse(files.has("tsconfig.json"));
    assert.isFalse(files.has("src/functions/repairs.ts"));
    // the rendered agent / plugin / manifest shapes hold identically for JS.
    assert.deepStrictEqual(readJson(files, "appPackage/repairDeclarativeAgent.json").actions[0], {
      id: "repairPlugin",
      file: "ai-plugin.json",
    });
    assert.strictEqual(readJson(files, "appPackage/ai-plugin.json").runtimes[0].auth.type, "None");
    assert.strictEqual(readJson(files, "package.json").name, "myagent");
  });

  it("SCN-CREATE-APIPLUGIN-07: the only pipeline step is require-empty-target; no post-render injection runs", async () => {
    const { outcome } = await run("typescript");
    assert.deepStrictEqual(outcome.stepsRun, ["require-empty-target"]);
    assert.isEmpty(outcome.stepsSkipped);
  });

  it("SCN-CREATE-APIPLUGIN-08: a non-empty target fails require-empty-target first and writes nothing", async () => {
    const runtime = createInMemoryRuntime();
    const request: ScaffoldRequest = {
      descriptor,
      pipeline,
      content,
      answers: {},
      callerFloor: { appName: "MyAgent", language: "typescript" },
      targetDir: { path: "/out", existing: ["appPackage/manifest.json"] },
    };
    const result = await scaffold(request, runtime);
    assert.isTrue(result.isErr());
    const error = result._unsafeUnwrapErr();
    assert.instanceOf(error, UserError);
    assert.strictEqual(error.name, REQUIRE_EMPTY_TARGET);
    assert.strictEqual(runtime.files.size, 0);
  });

  it("SCN-CREATE-APIPLUGIN-09: an identical re-run is deterministic (same written set and rendered agent)", async () => {
    const first = await run("typescript");
    const second = await run("typescript");
    assert.deepStrictEqual([...first.outcome.written].sort(), [...second.outcome.written].sort());
    assert.strictEqual(
      readJson(first.files, "appPackage/repairDeclarativeAgent.json").name,
      readJson(second.files, "appPackage/repairDeclarativeAgent.json").name
    );
  });
});
