// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import { UserError } from "@microsoft/teamsfx-api";
import { REQUIRE_EMPTY_TARGET } from "../../../src/v4/pipeline/runScaffoldPipeline";
import { createInMemoryRuntime } from "../../../src/v4/runtime/inMemoryRuntime";
import { scaffold } from "../../../src/v4/runtime/scaffold";
import {
  loadV4Package,
  readJsonObject,
  recordProperty,
  runV4Package,
  text,
  V4ScenarioOutcome,
} from "./helpers/scenarioHarness";

const EXPECTED_FILES = [
  ".gitignore",
  ".vscode/extensions.json",
  ".vscode/launch.json",
  ".vscode/settings.json",
  ".vscode/tasks.json",
  "AGENTS.md",
  "README.md",
  "appPackage/adaptiveCards/searchIssues.json",
  "appPackage/color.png",
  "appPackage/manifest.json",
  "appPackage/outline.png",
  "assets/image.png",
  "env/.env.dev",
  "env/.env.local",
  "evals/prompts.json",
  "m365agents.local.yml",
  "m365agents.yml",
  "package.json",
  "scripts/generate-env.js",
  "src/agent/actions/github.tsp",
  "src/agent/env.tsp",
  "src/agent/main.tsp",
  "src/agent/prompts/instructions.tsp",
  "tspconfig.yaml",
];

const templatePackage = loadV4Package("create", "da/typespec");

interface RunOptions {
  existing?: string[];
}

async function run(
  options: RunOptions = {}
): Promise<{ files: Map<string, Buffer>; outcome: V4ScenarioOutcome }> {
  return runV4Package(templatePackage, {
    callerFloor: { appName: "My Agent", language: "common" },
    existing: options.existing,
  });
}

describe("SCN-DA-CREATE-TYPESPEC (v4, T3 InMemoryRuntime)", () => {
  it("SCN-CREATE-TYPESPEC-01: the render phase writes exactly the TypeSpec DA file set (empty target)", async () => {
    const { outcome } = await run();
    assert.deepStrictEqual([...outcome.written].sort(), EXPECTED_FILES);
    assert.isEmpty(outcome.skipped);
  });

  it("SCN-CREATE-TYPESPEC-02: package.json renders the safe package name and TypeSpec dependencies", async () => {
    const { files } = await run();
    const pkg = readJsonObject(files, "package.json");
    const devDependencies = recordProperty(pkg, "devDependencies");
    assert.strictEqual(pkg.name, "myagent");
    assert.property(devDependencies, "@microsoft/typespec-m365-copilot");
    assert.property(devDependencies, "@typespec/compiler");
  });

  it("SCN-CREATE-TYPESPEC-03: main.tsp renders display name and safe namespace separately", async () => {
    const { files } = await run();
    const main = text(files, "src/agent/main.tsp");
    assert.include(main, 'import "@microsoft/typespec-m365-copilot";');
    assert.include(main, '  "My Agent",');
    assert.include(main, "namespace myagent {");
    assert.notInclude(main, "namespace My Agent");
  });

  it("SCN-CREATE-TYPESPEC-04: manifest.json preserves env refs and leaves TypeSpec output to compile", async () => {
    const { files } = await run();
    const manifest = readJsonObject(files, "appPackage/manifest.json");
    const name = recordProperty(manifest, "name");
    assert.strictEqual(manifest.manifestVersion, "1.28");
    assert.strictEqual(manifest.id, "${{TEAMS_APP_ID}}");
    assert.strictEqual(name.short, "My Agent${{APP_NAME_SUFFIX}}");
    assert.notProperty(manifest, "copilotAgents");
  });

  it("SCN-CREATE-TYPESPEC-05: project yaml includes TypeSpec compile and scaffold runs only require-empty-target", async () => {
    const { files, outcome } = await run();
    const yml = text(files, "m365agents.yml");
    assert.deepStrictEqual(outcome.stepsRun, ["require-empty-target"]);
    assert.isEmpty(outcome.stepsSkipped);
    assert.include(yml, "uses: cli/runNpmCommand");
    assert.include(yml, "uses: typeSpec/compile");
    assert.include(yml, "manifestPath: ./appPackage/manifest.json");
    assert.notInclude(yml, "oauth/register");
    assert.notInclude(yml, "pluginManifestPath");
  });

  it("SCN-CREATE-TYPESPEC-06: a non-empty target fails require-empty-target first and writes nothing", async () => {
    const runtime = createInMemoryRuntime();
    const request = {
      descriptor: templatePackage.descriptor,
      pipeline: templatePackage.pipeline,
      content: templatePackage.content,
      answers: {},
      callerFloor: { appName: "My Agent", language: "common" },
      targetDir: { path: "/out", existing: ["appPackage/manifest.json"] },
    };
    const result = await scaffold(request, runtime);
    assert.isTrue(result.isErr());
    const error = result._unsafeUnwrapErr();
    assert.instanceOf(error, UserError);
    assert.strictEqual(error.name, REQUIRE_EMPTY_TARGET);
    assert.strictEqual(runtime.files.size, 0);
  });

  it("SCN-CREATE-TYPESPEC-07: an identical re-run is deterministic (same written set and rendered namespace)", async () => {
    const first = await run();
    const second = await run();
    assert.deepStrictEqual([...first.outcome.written].sort(), [...second.outcome.written].sort());
    assert.strictEqual(
      text(first.files, "src/agent/main.tsp"),
      text(second.files, "src/agent/main.tsp")
    );
  });
});
