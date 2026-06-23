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
  recordArrayProperty,
  recordProperty,
  runV4Package,
  text,
  V4ScenarioOutcome,
} from "./helpers/scenarioHarness";

const EXPECTED_FILES = [
  ".funcignore",
  ".gitignore",
  ".vscode/extensions.json",
  ".vscode/launch.json",
  ".vscode/settings.json",
  ".vscode/tasks.json",
  "README.md",
  "aad.manifest.json",
  "api.http",
  "appPackage/color.png",
  "appPackage/declarativeAgent.json",
  "appPackage/instruction.txt",
  "appPackage/manifest.json",
  "appPackage/outline.png",
  "assets/add-cc-with-da.png",
  "assets/copilot-results.png",
  "env/.env.dev",
  "env/.env.local",
  "host.json",
  "infra/azure.bicep",
  "infra/azure.parameters.json",
  "m365agents.local.yml",
  "m365agents.yml",
  "package.json",
  "scripts/Clear-Issues.ps1",
  "scripts/Init-Issues.ps1",
  "scripts/admin-consent.js",
  "scripts/issues.json",
  "src/config.ts",
  "src/connection.ts",
  "src/custom/getAclFromItem.ts",
  "src/custom/getAllItemsFromAPI.ts",
  "src/custom/getExternalItemFromItem.ts",
  "src/functions/connections.ts",
  "src/graphClient.ts",
  "src/ingest.ts",
  "src/longRunningOperationMiddleware.ts",
  "src/models/Config.ts",
  "src/models/Item.ts",
  "src/references/schema.json",
  "src/references/template.json",
  "src/schema.ts",
  "src/services/crawlService.ts",
  "src/services/itemsService.ts",
  "src/utils.ts",
  "tsconfig.json",
];

const templatePackage = loadV4Package("create", "da/graph-connector");

interface RunOptions {
  existing?: string[];
}

async function run(
  options: RunOptions = {}
): Promise<{ files: Map<string, Buffer>; outcome: V4ScenarioOutcome }> {
  return runV4Package(templatePackage, {
    answers: {
      graphConnectorName: "GitHub Issues",
      graphConnectorConnectionId: "githubissues",
    },
    callerFloor: { appName: "My Agent", language: "common" },
    existing: options.existing,
  });
}

describe("SCN-DA-CREATE-GRAPH-CONNECTOR (v4, T3 InMemoryRuntime)", () => {
  it("SCN-CREATE-GC-01: the render phase writes exactly the flattened connector and DA file set", async () => {
    const { outcome } = await run();
    assert.deepStrictEqual([...outcome.written].sort(), EXPECTED_FILES);
    assert.isEmpty(outcome.skipped);
  });

  it("SCN-CREATE-GC-02: declarativeAgent.json renders the GraphConnectors capability", async () => {
    const { files } = await run();
    const agent = readJsonObject(files, "appPackage/declarativeAgent.json");
    const capabilities = recordArrayProperty(agent, "capabilities");
    const graphCapability = capabilities[0];
    const connections = recordArrayProperty(graphCapability, "connections");
    assert.strictEqual(agent.version, "v1.7");
    assert.strictEqual(agent.name, "My Agent${{APP_NAME_SUFFIX}}");
    assert.strictEqual(agent.instructions, "$[file('instruction.txt')]");
    assert.strictEqual(graphCapability.name, "GraphConnectors");
    assert.deepStrictEqual(connections[0], { connection_id: "${{CONNECTOR_ID}}" });
    assert.notProperty(agent, "sensitivity_label");
  });

  it("SCN-CREATE-GC-03: env files render connector id and name from Q2 answers", async () => {
    const { files } = await run();
    const localEnv = text(files, "env/.env.local");
    const devEnv = text(files, "env/.env.dev");
    assert.include(localEnv, "CONNECTOR_ID=githubissues");
    assert.include(localEnv, "CONNECTOR_NAME=GitHub Issues");
    assert.include(devEnv, "CONNECTOR_ID=");
    assert.include(devEnv, "CONNECTOR_NAME=GitHub Issues");
  });

  it("SCN-CREATE-GC-04: package and yaml render the connector project plus DA stages", async () => {
    const { files } = await run();
    const yml = text(files, "m365agents.yml");
    assert.strictEqual(readJsonObject(files, "package.json").name, "myagent");
    assert.include(yml, "uses: teamsApp/create");
    assert.include(yml, "uses: aadApp/create");
    assert.include(yml, "deploymentName: Create-resources-for-gc");
    assert.include(yml, "uses: azureFunctions/zipDeploy");
  });

  it("SCN-CREATE-GC-05: scaffold only runs require-empty-target and no post-render copy step", async () => {
    const { outcome } = await run();
    assert.deepStrictEqual(outcome.stepsRun, ["require-empty-target"]);
    assert.isEmpty(outcome.stepsSkipped);
  });

  it("SCN-CREATE-GC-06: a non-empty target fails require-empty-target first and writes nothing", async () => {
    const runtime = createInMemoryRuntime();
    const result = await scaffold(
      {
        descriptor: templatePackage.descriptor,
        pipeline: templatePackage.pipeline,
        content: templatePackage.content,
        answers: {
          graphConnectorName: "GitHub Issues",
          graphConnectorConnectionId: "githubissues",
        },
        callerFloor: { appName: "My Agent", language: "common" },
        targetDir: { path: "/out", existing: ["appPackage/manifest.json"] },
      },
      runtime
    );
    assert.isTrue(result.isErr());
    const error = result._unsafeUnwrapErr();
    assert.instanceOf(error, UserError);
    assert.strictEqual(error.name, REQUIRE_EMPTY_TARGET);
    assert.strictEqual(runtime.files.size, 0);
  });

  it("SCN-CREATE-GC-07: an identical re-run is deterministic (same written set and connector env)", async () => {
    const first = await run();
    const second = await run();
    assert.deepStrictEqual([...first.outcome.written].sort(), [...second.outcome.written].sort());
    assert.strictEqual(text(first.files, "env/.env.local"), text(second.files, "env/.env.local"));
  });
});
