// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import { UserError } from "@microsoft/teamsfx-api";
import { REQUIRE_EMPTY_TARGET } from "../../../src/v4/pipeline/runScaffoldPipeline";
import { createInMemoryRuntime } from "../../../src/v4/runtime/inMemoryRuntime";
import { scaffold } from "../../../src/v4/runtime/scaffold";
import {
  isRecordArray,
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
  "README.md",
  "appPackage/color.png",
  "appPackage/declarativeAgent.json",
  "appPackage/instruction.txt",
  "appPackage/manifest.json",
  "appPackage/outline.png",
  "appPackage/skills/hello-atk/SKILL.md",
  "env/.env.dev",
  "env/.env.local",
  "evals/prompts.json",
  "m365agents.local.yml",
  "m365agents.yml",
];

const templatePackage = loadV4Package("create", "da/skill");

interface RunOptions {
  existing?: string[];
}

async function run(
  options: RunOptions = {}
): Promise<{ files: Map<string, Buffer>; outcome: V4ScenarioOutcome }> {
  return runV4Package(templatePackage, {
    callerFloor: { appName: "MyAgent", language: "common" },
    existing: options.existing,
  });
}

describe("SCN-DA-CREATE-SKILL (v4, T3 InMemoryRuntime)", () => {
  it("SCN-CREATE-SKILL-01: the render phase writes exactly the skill-DA file set (empty target)", async () => {
    const { outcome } = await run();
    assert.deepStrictEqual([...outcome.written].sort(), EXPECTED_FILES);
    assert.isEmpty(outcome.skipped);
  });

  it("SCN-CREATE-SKILL-02: declarativeAgent.json renders the hello-atk agent skill", async () => {
    const { files } = await run();
    const agent = readJsonObject(files, "appPackage/declarativeAgent.json");
    assert.strictEqual(agent.version, "v1.8");
    assert.strictEqual(agent.name, "MyAgent${{APP_NAME_SUFFIX}}");
    assert.strictEqual(agent.instructions, "$[file('instruction.txt')]");
    assert.deepStrictEqual(agent.agent_skills, [{ folder: "skills/hello-atk" }]);
    assert.notProperty(agent, "capabilities");
    assert.notProperty(agent, "sensitivity_label");
  });

  it("SCN-CREATE-SKILL-03: manifest.json wires the single declarative agent and preserves the env refs", async () => {
    const { files } = await run();
    const manifest = readJsonObject(files, "appPackage/manifest.json");
    const name = recordProperty(manifest, "name");
    const copilotAgents = recordProperty(manifest, "copilotAgents");
    const agents = copilotAgents.declarativeAgents;
    assert.strictEqual(manifest.manifestVersion, "1.28");
    assert.strictEqual(manifest.id, "${{TEAMS_APP_ID}}");
    assert.strictEqual(name.short, "MyAgent${{APP_NAME_SUFFIX}}");
    assert.isTrue(isRecordArray(agents));
    assert.lengthOf(agents, 1);
    assert.deepStrictEqual(agents[0], { id: "declarativeAgent", file: "declarativeAgent.json" });
  });

  it("SCN-CREATE-SKILL-04: the hello-atk skill file carries the expected skill metadata", async () => {
    const { files } = await run();
    const skill = text(files, "appPackage/skills/hello-atk/SKILL.md");
    assert.include(skill, "name: hello-atk");
    assert.include(skill, "fun fact about declarative agents");
  });

  it("SCN-CREATE-SKILL-05: the only pipeline step is require-empty-target; no action wiring is injected", async () => {
    const { files, outcome } = await run();
    const yml = text(files, "m365agents.yml");
    assert.deepStrictEqual(outcome.stepsRun, ["require-empty-target"]);
    assert.isEmpty(outcome.stepsSkipped);
    assert.include(yml, "version: v1.12");
    assert.notInclude(yml, "oauth/register");
    assert.notInclude(yml, "pluginManifestPath");
  });

  it("SCN-CREATE-SKILL-06: a non-empty target fails require-empty-target first and writes nothing", async () => {
    const runtime = createInMemoryRuntime();
    const request = {
      descriptor: templatePackage.descriptor,
      pipeline: templatePackage.pipeline,
      content: templatePackage.content,
      answers: {},
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

  it("SCN-CREATE-SKILL-07: an identical re-run is deterministic (same written set and rendered skill)", async () => {
    const first = await run();
    const second = await run();
    assert.deepStrictEqual([...first.outcome.written].sort(), [...second.outcome.written].sort());
    assert.deepStrictEqual(
      readJsonObject(first.files, "appPackage/declarativeAgent.json").agent_skills,
      readJsonObject(second.files, "appPackage/declarativeAgent.json").agent_skills
    );
  });
});
