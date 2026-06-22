// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import * as fs from "fs";
import * as path from "path";
import { TemplateFileEntry } from "../../../src/v4/model/dataModel";
import { createInMemoryRuntime } from "../../../src/v4/runtime/inMemoryRuntime";
import { ScaffoldRequest, scaffold } from "../../../src/v4/runtime/scaffold";

/**
 * T3 scenario tier: the `modify/add-mcp-server` package applied to an existing
 * DA project under `InMemoryRuntime`.
 *
 * Spec: docs/03-specs/scenarios/da/add-mcp-server.md (SCN-ADD-MCP-01..09)
 */

const PKG_DIR = path.resolve(__dirname, "../../../../../templates/v4/modify/add-mcp-server");

const MCP_SERVER_URL = "https://api.github.com/mcp";
const NAMESPACE = "apigithubc";
const PLUGIN_PATH = `appPackage/ai-plugin-${NAMESPACE}.json`;
const TEAMS_MANIFEST_PATH = "appPackage/manifest.json";
const DA_MANIFEST_PATH = "appPackage/declarativeAgent.json";
const YML_PATH = "m365agents.yml";
const ENV_PATH = "env/.env.dev";
const AUTH_REF = "${{MCP_DA_AUTH_ID_APIGITHUBC}}";
const AUTH_ENV_VAR = "MCP_DA_AUTH_ID_APIGITHUBC";

const descriptor: unknown = JSON.parse(
  fs.readFileSync(path.join(PKG_DIR, "descriptor.json"), "utf8")
);
const pipeline: unknown = JSON.parse(fs.readFileSync(path.join(PKG_DIR, "pipeline.json"), "utf8"));

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

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isRecordArray(value: unknown): value is Record<string, unknown>[] {
  return Array.isArray(value) && value.every(isRecord);
}

function text(files: Map<string, Buffer>, filePath: string): string {
  const buf = files.get(filePath);
  assert.isDefined(buf, `expected '${filePath}' to exist`);
  return (buf ?? Buffer.from("", "utf8")).toString("utf8");
}

function readJsonObject(files: Map<string, Buffer>, filePath: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(text(files, filePath));
  assert.isTrue(isRecord(parsed));
  return parsed;
}

function actions(manifest: Record<string, unknown>): Record<string, unknown>[] {
  const value = manifest.actions;
  assert.isTrue(isRecordArray(value));
  return value;
}

function runtimes(plugin: Record<string, unknown>): Record<string, unknown>[] {
  const value = plugin.runtimes;
  assert.isTrue(isRecordArray(value));
  return value;
}

function auth(runtime: Record<string, unknown>): Record<string, unknown> {
  const value = runtime.auth;
  assert.isTrue(isRecord(value));
  return value;
}

interface RunOptions {
  authType?: string;
}

async function run(options: RunOptions = {}): Promise<{
  files: Map<string, Buffer>;
  outcome: Awaited<ReturnType<typeof unwrapOutcome>>;
}> {
  const authType = options.authType ?? "none";
  const runtime = createInMemoryRuntime();
  runtime.files.set(
    TEAMS_MANIFEST_PATH,
    Buffer.from(
      JSON.stringify({
        copilotAgents: {
          declarativeAgents: [{ id: "declarativeAgent", file: "declarativeAgent.json" }],
        },
      }),
      "utf8"
    )
  );
  runtime.files.set(
    DA_MANIFEST_PATH,
    Buffer.from(JSON.stringify({ name: "Existing Agent" }), "utf8")
  );
  runtime.files.set(YML_PATH, Buffer.from(["version: v1.12", "provision:"].join("\n"), "utf8"));
  runtime.files.set(ENV_PATH, Buffer.from("TEAMSFX_ENV=dev\n", "utf8"));

  const request: ScaffoldRequest = {
    descriptor,
    pipeline,
    content,
    answers: { mcpServerUrl: MCP_SERVER_URL, teamsManifestPath: TEAMS_MANIFEST_PATH, authType },
    callerFloor: { appName: "Existing Agent", language: "common" },
    targetDir: {
      path: "/project",
      existing: [TEAMS_MANIFEST_PATH, DA_MANIFEST_PATH, YML_PATH, ENV_PATH],
    },
  };
  const result = await scaffold(request, runtime);
  return { files: runtime.files, outcome: unwrapOutcome(result) };
}

function unwrapOutcome(result: Awaited<ReturnType<typeof scaffold>>) {
  assert.isTrue(result.isOk(), result.isErr() ? result.error.message : "expected ok");
  return result._unsafeUnwrap();
}

describe("SCN-DA-ADD-MCP-ACTION-TO-DA (v4, T3 InMemoryRuntime)", () => {
  it("SCN-ADD-MCP-01: writes only the dynamic plugin manifest in the render phase", async () => {
    const { outcome } = await run();
    assert.deepStrictEqual(outcome.written, [PLUGIN_PATH]);
  });

  it("SCN-ADD-MCP-02/03: renders namespace and RemoteMCPServer dynamic discovery", async () => {
    const { files } = await run();
    const plugin = readJsonObject(files, PLUGIN_PATH);
    assert.strictEqual(plugin.namespace, NAMESPACE);
    const runtime = runtimes(plugin)[0];
    assert.strictEqual(runtime.type, "RemoteMCPServer");
    const spec = runtime.spec;
    assert.isTrue(isRecord(spec));
    assert.strictEqual(spec.url, MCP_SERVER_URL);
    assert.strictEqual(spec.enable_dynamic_discovery, true);
  });

  it("SCN-ADD-MCP-04: registers the rendered plugin in the existing DA manifest", async () => {
    const { files, outcome } = await run();
    assert.include(outcome.stepsRun, "da-action/register-plugin-manifest");
    const daManifest = readJsonObject(files, DA_MANIFEST_PATH);
    assert.deepInclude(actions(daManifest), { id: NAMESPACE, file: `ai-plugin-${NAMESPACE}.json` });
  });

  it("SCN-ADD-MCP-06: authType=oauth renders OAuthPluginVault and injects oauth/register", async () => {
    const { files, outcome } = await run({ authType: "oauth" });
    const runtime = runtimes(readJsonObject(files, PLUGIN_PATH))[0];
    assert.strictEqual(auth(runtime).type, "OAuthPluginVault");
    assert.strictEqual(auth(runtime).reference_id, AUTH_REF);
    assert.include(outcome.stepsRun, "mcp-auth/inject-yml-action");
    assert.include(text(files, YML_PATH), "oauth/register");
  });

  it("SCN-ADD-MCP-07: authType oauth/entra-sso persists MCP_DA_AUTH_ID_<NS>", async () => {
    for (const authType of ["oauth", "entra-sso"]) {
      const { files, outcome } = await run({ authType });
      assert.include(outcome.stepsRun, "mcp-auth/persist-credential-env");
      assert.include(text(files, ENV_PATH), `${AUTH_ENV_VAR}=`);
    }
  });

  it("SCN-ADD-MCP-08: authType=none renders auth None and skips auth wiring steps", async () => {
    const { files, outcome } = await run({ authType: "none" });
    const runtime = runtimes(readJsonObject(files, PLUGIN_PATH))[0];
    assert.strictEqual(auth(runtime).type, "None");
    assert.include(outcome.stepsSkipped, "mcp-auth/inject-yml-action");
    assert.include(outcome.stepsSkipped, "mcp-auth/persist-credential-env");
    assert.notInclude(text(files, ENV_PATH), "MCP_DA_AUTH_ID_");
  });

  it("SCN-ADD-MCP-05: a same-URL re-run skips render collision and does not duplicate the action", async () => {
    const first = await run();
    const runtime = createInMemoryRuntime();
    for (const [filePath, body] of first.files.entries()) {
      runtime.files.set(filePath, body);
    }
    const result = await scaffold(
      {
        descriptor,
        pipeline,
        content,
        answers: {
          mcpServerUrl: MCP_SERVER_URL,
          teamsManifestPath: TEAMS_MANIFEST_PATH,
          authType: "none",
        },
        callerFloor: { appName: "Existing Agent", language: "common" },
        targetDir: {
          path: "/project",
          existing: [TEAMS_MANIFEST_PATH, DA_MANIFEST_PATH, PLUGIN_PATH],
        },
      },
      runtime
    );

    const outcome = unwrapOutcome(result);
    assert.deepStrictEqual(outcome.written, []);
    assert.deepStrictEqual(
      outcome.skipped.map((item) => item.path),
      [PLUGIN_PATH]
    );
    const daManifest = readJsonObject(runtime.files, DA_MANIFEST_PATH);
    assert.lengthOf(actions(daManifest), 1);
  });
});
