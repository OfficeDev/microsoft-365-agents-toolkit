// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs";
import path from "path";
import { load } from "js-yaml";
import { assert, afterEach, beforeEach, vi } from "vitest";
import { featureFlagManager } from "../../../src/common/featureFlags";
import { renderTemplateFileData } from "../../../src/component/generator/utils";
import {
  isRecord,
  loadV4Package,
  recordArrayProperty,
  recordProperty,
  runV4Package,
  text,
} from "../../v4/scenarios/helpers/scenarioHarness";

const templatesRoot = path.resolve(__dirname, "../../../../../templates");
const scratchTemplates = [
  "declarative-agent-with-action-from-scratch",
  "declarative-agent-with-action-from-scratch-bearer",
  "declarative-agent-with-action-from-scratch-oauth",
];
const legacyTemplates = [
  ...["ts", "js"].flatMap((language) =>
    scratchTemplates.map((template) => path.join("vsc", language, template))
  ),
  ...scratchTemplates.map((template) => path.join("vs", "csharp", template)),
  path.join("vs", "csharp", "declarative-agent-with-action-from-existing-api"),
  path.join("vsc", "common", "declarative-agent-with-action-from-existing-api"),
  path.join("vsc", "common", "declarative-agent-typespec"),
  path.join("vsc", "common", "declarative-agent-with-action-from-mcp"),
];

function publishAction(yml: string, stage = "provision"): Record<string, unknown> {
  const lifecycle: unknown = load(yml);
  if (!isRecord(lifecycle)) {
    assert.fail("Expected a lifecycle object");
  }
  const action = recordArrayProperty(lifecycle, stage).find(
    (step) => step.uses === "copilotAgent/publish"
  );
  if (!action) {
    assert.fail(`Missing copilotAgent/publish in ${stage}`);
  }
  return action;
}

function assertRemote(files: Map<string, Buffer>): void {
  assert.match(text(files, "env/.env.dev"), /^AGENT_SCOPE=shared\r?$/m);
  const action = publishAction(text(files, "m365agents.yml"));
  assert.strictEqual(recordProperty(action, "with").scope, "${{AGENT_SCOPE}}");
}

function assertLocal(files: Map<string, Buffer>): void {
  const action = publishAction(text(files, "m365agents.local.yml"));
  const scope = recordProperty(action, "with").scope;
  assert.isTrue(scope === undefined || scope === "personal");
  assert.notProperty(recordProperty(action, "writeToEnvironmentFile"), "shareLink");
}

function assertShareLink(files: Map<string, Buffer>): void {
  const action = publishAction(text(files, "m365agents.yml"));
  assert.strictEqual(recordProperty(action, "writeToEnvironmentFile").shareLink, "SHARE_LINK");
}

function renderLegacy(template: string): Map<string, Buffer> {
  const root = path.join(templatesRoot, template);
  const files = new Map<string, Buffer>();
  const env = fs.readdirSync(path.join(root, "env"));
  const names = [
    "m365agents.yml.tpl",
    ...env
      .filter((name) => /^\.env\.(dev|local)(\.tpl)?$/.test(name))
      .map((name) => path.join("env", name)),
  ];
  if (supportsLocal(template)) {
    names.push("m365agents.local.yml.tpl");
  }
  for (const name of names) {
    const rendered = renderTemplateFileData(name, fs.readFileSync(path.join(root, name)), {
      appName: "ScopeAgent",
      ProjectName: "ScopeAgent",
      DeclarativeCopilot: "true",
      ApiKey: false,
      OAuth: false,
      MicrosoftEntra: false,
      PlaceProjectFileInSolutionDir: false,
    });
    files.set(
      name.replace(/\\/g, "/").replace(/\.tpl$/, ""),
      Buffer.isBuffer(rendered) ? rendered : Buffer.from(rendered)
    );
  }
  return files;
}

function supportsLocal(template: string): boolean {
  return (
    !template.endsWith("declarative-agent-with-action-from-mcp") &&
    template !== path.join("vs", "csharp", "declarative-agent-with-action-from-existing-api")
  );
}

describe("Advanced DA scope defaults (#16867)", () => {
  afterEach(() => vi.restoreAllMocks());

  for (const template of legacyTemplates) {
    describe(template, () => {
      it("AC-01: remote provisioning explicitly uses shared scope", () => {
        assertRemote(renderLegacy(template));
      });

      if (supportsLocal(template)) {
        it("AC-02: local debugging stays personal without inheriting dev scope", () => {
          assertLocal(renderLegacy(template));
        });
      }

      it("AC-06: remote provisioning persists the share link", () => {
        assertShareLink(renderLegacy(template));
      });
    });
  }

  const v4Templates = [
    "da/api-plugin-from-scratch",
    "da/api-plugin-from-scratch-bearer",
    "da/api-plugin-from-scratch-oauth",
    "da/api-plugin-from-existing-api",
    "da/typespec",
    "da/mcp-server",
    "da/mcp-server-static",
  ];
  for (const template of v4Templates) {
    const templatePackage = loadV4Package("create", template);
    const languages = template.includes("from-scratch") ? ["typescript", "javascript"] : ["common"];
    for (const language of languages) {
      describe(`v4 ${template} (${language})`, () => {
        beforeEach(() => {
          vi.spyOn(featureFlagManager, "getBooleanValue").mockReturnValue(false);
        });

        async function run() {
          const specPath = path.resolve(
            __dirname,
            "../../v4/scenarios/fixtures/repairs-openapi.yaml"
          );
          return runV4Package(templatePackage, {
            callerFloor: { appName: "ScopeAgent", language },
            answers: {
              surface: "cli",
              apiAuth: "oauth",
              apiSpecLocation: specPath,
              "derived.openapi.operations.apiSpecLocation": specPath,
              apiOperations: ["GET /repairs"],
              mcpServerType: "remote",
              mcpServerUrl: "https://example.com/mcp",
              authType: "none",
              mcpToolsJson: JSON.stringify({
                tools: [{ name: "search", description: "Search", inputSchema: { type: "object" } }],
              }),
              selectedMcpTools: ["search"],
            },
          });
        }

        it("AC-01: scaffolded remote provisioning explicitly uses shared scope", async () => {
          assertRemote((await run()).files);
        });

        if (!template.includes("mcp-server")) {
          it("AC-02: scaffolded local debugging stays personal independently of dev", async () => {
            assertLocal((await run()).files);
          });
        }

        it("AC-06: scaffolded remote provisioning persists the share link", async () => {
          assertShareLink((await run()).files);
        });
      });
    }
  }
});
