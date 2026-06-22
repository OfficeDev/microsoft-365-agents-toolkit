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
 * T3 scenario tier (ADR-0018): the whole `da/api-plugin-from-scratch-oauth`
 * create package — the declarative agent with a brand-new, **OAuth- / Microsoft-
 * Entra-protected** API plugin action — scaffolded under `InMemoryRuntime`,
 * asserting the vertical contract for each declared language.
 *
 * Spec: docs/03-specs/scenarios/da/create-api-plugin-from-scratch-oauth.md
 * (SCN-CREATE-APIPLUGIN-OAUTH-01..11)
 *
 * One package serves **two** action sources — `microsoft-entra` and `oauth` —
 * distinguished by the `apiAuth` selector dimension. The descriptor's `replaceMap`
 * carries a single conditional render var, `{ MicrosoftEntra: when apiAuth ==
 * 'microsoft-entra' }`; the `{{#MicrosoftEntra}}` / `{{^MicrosoftEntra}}` sections
 * (ai-plugin.json, aad.manifest.json, m365agents*.yml, README) then select the
 * Entra vs. generic-OAuth wiring. The file **set** is identical for both sources;
 * only the rendered content of those sections differs. No v3 symbol participates.
 */

const PKG_DIR = path.resolve(
  __dirname,
  "../../../../../templates/v4/create/da/api-plugin-from-scratch-oauth"
);

/** The TypeScript backend file set a scaffold writes (`.tpl` stripped, language prefix stripped). */
const EXPECTED_TS_FILES = [
  ".funcignore",
  ".gitignore",
  ".tours/custom-token-validation-without-using-Easy-Auth.tour",
  ".vscode/extensions.json",
  ".vscode/launch.json",
  ".vscode/settings.json",
  ".vscode/tasks.json",
  "aad.manifest.json",
  "appPackage/adaptiveCards/listRepairs.data.json",
  "appPackage/adaptiveCards/listRepairs.json",
  "appPackage/ai-plugin.json",
  "appPackage/ai-plugin.local.json",
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
  "README.md",
  "src/functions/middleware/authMiddleware.ts",
  "src/functions/middleware/config.ts",
  "src/functions/middleware/tokenCacheWrapper.ts",
  "src/functions/middleware/tokenValidator.ts",
  "src/functions/middleware/utils.ts",
  "src/functions/repairs.ts",
  "src/repairsData.json",
  "tsconfig.json",
];

/**
 * The JavaScript backend file set — the TS set minus `tsconfig.json`, with every
 * `src/functions/**.ts` (the middleware + `repairs.ts`) re-suffixed `.js`.
 */
const EXPECTED_JS_FILES = EXPECTED_TS_FILES.filter((f) => f !== "tsconfig.json").map((f) =>
  f.startsWith("src/functions/") && f.endsWith(".ts") ? f.replace(/\.ts$/, ".js") : f
);

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

/**
 * Scaffold the package for one language + `apiAuth` source against a fresh
 * in-memory runtime. `apiAuth` ('microsoft-entra' | 'oauth') is the selector
 * dimension the descriptor's `MicrosoftEntra` render var keys off.
 */
async function run(
  language: string,
  apiAuth: string,
  options: RunOptions = {}
): Promise<{ files: Map<string, Buffer>; outcome: ReturnType<typeof unwrapOutcome> }> {
  const runtime = createInMemoryRuntime();
  const request: ScaffoldRequest = {
    descriptor,
    pipeline,
    content,
    answers: { apiAuth },
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

describe("SCN-DA-CREATE-API-PLUGIN-FROM-SCRATCH-OAUTH (v4, T3 InMemoryRuntime)", () => {
  it("SCN-CREATE-APIPLUGIN-OAUTH-01: the render phase writes exactly the TypeScript file set (incl. the auth middleware + aad.manifest.json)", async () => {
    const { outcome } = await run("typescript", "microsoft-entra");
    assert.deepStrictEqual([...outcome.written].sort(), [...EXPECTED_TS_FILES].sort());
    assert.isEmpty(outcome.skipped);
  });

  it("SCN-CREATE-APIPLUGIN-OAUTH-02: repairDeclarativeAgent.json wires the pre-baked action and omits the sensitivity label", async () => {
    const { files } = await run("typescript", "microsoft-entra");
    const agent = readJson(files, "appPackage/repairDeclarativeAgent.json");
    assert.strictEqual(agent.name, "MyAgent${{APP_NAME_SUFFIX}}");
    assert.strictEqual(agent.instructions, "$[file('instruction.txt')]");
    assert.lengthOf(agent.actions, 1);
    assert.deepStrictEqual(agent.actions[0], { id: "repairPlugin", file: "ai-plugin.json" });
    assert.notProperty(agent, "sensitivity_label");
  });

  it("SCN-CREATE-APIPLUGIN-OAUTH-03: apiAuth 'microsoft-entra' renders the OAuthPluginVault runtime with the Entra reference_id", async () => {
    const { files } = await run("typescript", "microsoft-entra");
    const plugin = readJson(files, "appPackage/ai-plugin.json");
    assert.strictEqual(plugin.namespace, "repairs");
    assert.lengthOf(plugin.runtimes, 1);
    const runtime = plugin.runtimes[0];
    assert.strictEqual(runtime.type, "OpenApi");
    assert.strictEqual(runtime.auth.type, "OAuthPluginVault");
    // the {{#MicrosoftEntra}} branch selects the Entra auth-code configuration ref.
    assert.strictEqual(runtime.auth.reference_id, "${{AADAUTHCODE_CONFIGURATION_ID}}");
    assert.strictEqual(runtime.spec.url, "apiSpecificationFile/repair.yml");
  });

  it("SCN-CREATE-APIPLUGIN-OAUTH-04: apiAuth 'oauth' renders the OAuthPluginVault runtime with the generic OAuth reference_id", async () => {
    const { files } = await run("typescript", "oauth");
    const runtime = readJson(files, "appPackage/ai-plugin.json").runtimes[0];
    assert.strictEqual(runtime.auth.type, "OAuthPluginVault");
    // MicrosoftEntra is absent (falsy) -> the {{^MicrosoftEntra}} branch selects the generic OAuth ref.
    assert.strictEqual(runtime.auth.reference_id, "${{OAUTH2AUTHCODE_CONFIGURATION_ID}}");
  });

  it("SCN-CREATE-APIPLUGIN-OAUTH-05: ai-plugin.local.json is the no-auth local runtime (both sources)", async () => {
    for (const apiAuth of ["microsoft-entra", "oauth"]) {
      const { files } = await run("typescript", apiAuth);
      const runtime = readJson(files, "appPackage/ai-plugin.local.json").runtimes[0];
      assert.strictEqual(runtime.auth.type, "None");
      assert.strictEqual(runtime.spec.url, "apiSpecificationFile/repair.local.yml");
    }
  });

  it("SCN-CREATE-APIPLUGIN-OAUTH-06: manifest.json wires the single declarative agent and preserves the env refs", async () => {
    const { files } = await run("typescript", "microsoft-entra");
    const manifest = readJson(files, "appPackage/manifest.json");
    assert.strictEqual(manifest.manifestVersion, "1.28");
    assert.strictEqual(manifest.id, "${{TEAMS_APP_ID}}");
    assert.strictEqual(manifest.name.short, "MyAgent${{APP_NAME_SUFFIX}}");
    const agents = manifest.copilotAgents.declarativeAgents;
    assert.lengthOf(agents, 1);
    assert.deepStrictEqual(agents[0], {
      id: "repairDeclarativeAgent",
      file: "repairDeclarativeAgent.json",
    });
  });

  it("SCN-CREATE-APIPLUGIN-OAUTH-07: language 'typescript' narrows to the typescript subtree, prefix stripped", async () => {
    const { files, outcome } = await run("typescript", "microsoft-entra");
    for (const written of outcome.written) {
      assert.isFalse(
        written.startsWith("typescript/") || written.startsWith("javascript/"),
        `written path '${written}' still carries a language prefix`
      );
    }
    assert.isTrue(files.has("src/functions/repairs.ts"));
    assert.isTrue(files.has("src/functions/middleware/authMiddleware.ts"));
    assert.isTrue(files.has("tsconfig.json"));
    assert.isFalse(files.has("src/functions/repairs.js"));
    // The TS oauth package.json hardcodes its name (a v3 template quirk preserved
    // verbatim by the 0-diff migration); the SafeProjectNameLowerCase producer is
    // exercised by the JS variant instead (SCN-CREATE-APIPLUGIN-OAUTH-08).
    assert.strictEqual(readJson(files, "package.json").name, "apipluginoauth");
  });

  it("SCN-CREATE-APIPLUGIN-OAUTH-08: language 'javascript' writes the javascript subtree with the same rendered auth shape", async () => {
    const { files, outcome } = await run("javascript", "microsoft-entra");
    assert.deepStrictEqual([...outcome.written].sort(), [...EXPECTED_JS_FILES].sort());
    assert.isTrue(files.has("src/functions/repairs.js"));
    assert.isTrue(files.has("src/functions/middleware/authMiddleware.js"));
    assert.isFalse(files.has("tsconfig.json"));
    assert.isFalse(files.has("src/functions/repairs.ts"));
    assert.strictEqual(
      readJson(files, "appPackage/ai-plugin.json").runtimes[0].auth.reference_id,
      "${{AADAUTHCODE_CONFIGURATION_ID}}"
    );
    assert.strictEqual(readJson(files, "package.json").name, "myagent");
  });

  it("SCN-CREATE-APIPLUGIN-OAUTH-09: the only pipeline step is require-empty-target; no post-render injection runs", async () => {
    const { outcome } = await run("typescript", "oauth");
    assert.deepStrictEqual(outcome.stepsRun, ["require-empty-target"]);
    assert.isEmpty(outcome.stepsSkipped);
  });

  it("SCN-CREATE-APIPLUGIN-OAUTH-10: a non-empty target fails require-empty-target first and writes nothing", async () => {
    const runtime = createInMemoryRuntime();
    const request: ScaffoldRequest = {
      descriptor,
      pipeline,
      content,
      answers: { apiAuth: "microsoft-entra" },
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

  it("SCN-CREATE-APIPLUGIN-OAUTH-11: the two sources differ only in the rendered auth ref — identical file set, distinct reference_id", async () => {
    const entra = await run("typescript", "microsoft-entra");
    const oauth = await run("typescript", "oauth");
    // identical written set: apiAuth changes content, not the emitted file list.
    assert.deepStrictEqual([...entra.outcome.written].sort(), [...oauth.outcome.written].sort());
    // but the rendered remote-plugin auth ref diverges on the MicrosoftEntra section.
    assert.notStrictEqual(
      readJson(entra.files, "appPackage/ai-plugin.json").runtimes[0].auth.reference_id,
      readJson(oauth.files, "appPackage/ai-plugin.json").runtimes[0].auth.reference_id
    );
  });
});
