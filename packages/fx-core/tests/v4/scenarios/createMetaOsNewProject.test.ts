// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import { REQUIRE_EMPTY_TARGET } from "../../../src/v4/pipeline/runScaffoldPipeline";
import { createInMemoryRuntime } from "../../../src/v4/runtime/inMemoryRuntime";
import { scaffold } from "../../../src/v4/runtime/scaffold";
import { assert } from "vitest";
import {
  isRecordArray,
  loadV4Package,
  readJsonObject,
  runV4Package,
  text,
} from "./helpers/scenarioHarness";

/**
 * T3 scenario tier: the `declarative-agent-meta-os-new-project` create package
 * scaffolded under `InMemoryRuntime`.
 *
 * Spec: docs/03-specs/scenarios/da/create-metaos-new-project.md
 * (SCN-CREATE-METAOS-01..05)
 */

const templatePackage = loadV4Package("create", "declarative-agent-meta-os-new-project");
const callerFloor = { appName: "MetaOS Agent", language: "common" };
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/;

function expectedWritten(): string[] {
  return templatePackage.content.map((entry) => entry.path.replace(/\.tpl$/, "")).sort();
}

async function run() {
  return runV4Package(templatePackage, { callerFloor });
}

describe("SCN-DA-CREATE-METAOS-NEW-PROJECT (v4, T3 InMemoryRuntime)", () => {
  it("SCN-CREATE-METAOS-01: the render phase writes the MetaOS new-project file set", async () => {
    const { outcome } = await run();
    assert.deepStrictEqual([...outcome.written].sort(), expectedWritten());
    assert.include(outcome.written, "src/taskpane/taskpane.ts");
    assert.include(outcome.written, "src/commands/commands.ts");
    assert.include(outcome.written, "appPackage/declarativeAgent.json");
    assert.include(outcome.written, "appPackage/alchemy-plugin.json");
  });

  it("SCN-CREATE-METAOS-02: unify-project-id writes the same UUID to manifest and env", async () => {
    const { files } = await run();
    const manifest = readJsonObject(files, "appPackage/manifest.json");
    assert.isString(manifest.id);
    const manifestId = typeof manifest.id === "string" ? manifest.id : "";
    assert.match(manifestId, UUID);
    assert.include(text(files, "env/.env.dev"), `TEAMS_APP_ID=${manifestId}`);
  });

  it("SCN-CREATE-METAOS-03: DA and action manifests preserve the Office action shape", async () => {
    const { files } = await run();
    const agent = readJsonObject(files, "appPackage/declarativeAgent.json");
    assert.isTrue(isRecordArray(agent.actions));
    assert.deepInclude(agent.actions, { id: "alchemyPlugin", file: "alchemy-plugin.json" });
    const action = readJsonObject(files, "appPackage/alchemy-plugin.json");
    assert.strictEqual(action.schema_version, "v2.4");
    assert.strictEqual(action.namespace, "AddInFunctions");
    assert.isTrue(isRecordArray(action.functions));
    const names = action.functions.map((item) => item.name);
    assert.sameMembers(names, ["addfooter", "fillcolor", "addtexttoslide"]);
  });

  it("SCN-CREATE-METAOS-04: pipeline steps run in order", async () => {
    const { outcome } = await run();
    assert.deepStrictEqual(outcome.stepsRun, ["require-empty-target", "metaos/unify-project-id"]);
    assert.isEmpty(outcome.stepsSkipped);
  });

  it("SCN-CREATE-METAOS-05: a non-empty target fails require-empty-target first", async () => {
    const runtime = createInMemoryRuntime();
    const result = await scaffold(
      {
        descriptor: templatePackage.descriptor,
        pipeline: templatePackage.pipeline,
        content: templatePackage.content,
        answers: {},
        callerFloor,
        targetDir: { path: "/out", existing: ["README.md"] },
      },
      runtime
    );
    assert.isTrue(result.isErr());
    const error = result._unsafeUnwrapErr();
    assert.instanceOf(error, UserError);
    assert.strictEqual(error.name, REQUIRE_EMPTY_TARGET);
    assert.strictEqual(runtime.files.size, 0);
  });
});
