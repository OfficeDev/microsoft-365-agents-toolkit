// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import os from "os";
import path from "path";
import { AppManifestUtils, TeamsManifestConverter } from "../src";

function previewManifest(overrides: Record<string, unknown> = {}): string {
  return JSON.stringify({
    $schema:
      "https://developer.microsoft.com/json-schemas/teams/vDevPreview/MicrosoftTeams.schema.json",
    manifestVersion: "devPreview",
    version: "1.0.0",
    id: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
    developer: {
      name: "Test",
      websiteUrl: "https://example.com",
      privacyUrl: "https://example.com/privacy",
      termsOfUseUrl: "https://example.com/terms",
    },
    name: { short: "Test", full: "Test app" },
    description: { short: "Test", full: "Test app" },
    icons: { color: "color.png", outline: "outline.png" },
    accentColor: "#FFFFFF",
    ...overrides,
  });
}

function connector(source: Record<string, unknown>) {
  return { id: "test", displayName: "Test", reusable: true, toolSource: source };
}

function toolSource(kind: string, authorization: Record<string, unknown>) {
  return {
    [kind]: {
      ...(kind === "remoteMcpServer" ? { mcpServerUrl: "https://example.com/mcp" } : {}),
      ...(kind === "localMcpServer" ? { mcpServerIdentifier: "test-server" } : {}),
      authorization,
    },
  };
}

async function validate(json: string): Promise<string[]> {
  return AppManifestUtils.validateAgainstSchema(TeamsManifestConverter.jsonToManifest(json));
}

describe("DevPreview schema snapshot", () => {
  for (const manifestVersion of ["devPreview", "m365DevPreview", "m365devpreview"]) {
    it(`round trips and validates ${manifestVersion} with the packaged schema`, async () => {
      const json = previewManifest({
        manifestVersion,
        agentConnectors: [
          connector(
            toolSource("openApiConnector", {
              type: "ConnectionVault",
              referenceId: "shared_kusto",
            })
          ),
        ],
      });
      const manifest = TeamsManifestConverter.jsonToManifest(json);
      assert.deepEqual(
        JSON.parse(TeamsManifestConverter.manifestToJson(manifest)),
        JSON.parse(json)
      );
      assert.deepEqual(await AppManifestUtils.validateAgainstSchema(manifest), []);
    });

    it(`does not bypass parsing validation for ${manifestVersion}`, () => {
      assert.throws(
        () => TeamsManifestConverter.jsonToManifest(previewManifest({ manifestVersion, id: 42 })),
        /id/
      );
    });

    it(`does not bypass serialization validation for ${manifestVersion}`, () => {
      const manifest = TeamsManifestConverter.jsonToManifest(previewManifest({ manifestVersion }));
      Object.assign(manifest, { id: 42 });
      assert.throws(() => TeamsManifestConverter.manifestToJson(manifest), /id/);
    });
  }

  for (const kind of ["openApiConnector", "remoteMcpServer", "localMcpServer"]) {
    for (const length of [128, 129]) {
      it(`${kind} enforces the ConnectionVault referenceId limit at ${length} characters`, async () => {
        const errors = await validate(
          previewManifest({
            agentConnectors: [
              connector(
                toolSource(kind, { type: "ConnectionVault", referenceId: "a".repeat(length) })
              ),
            ],
          })
        );
        if (length === 128) {
          assert.deepEqual(errors, []);
        } else {
          assert.include(errors.join("\n"), "128");
        }
      });
    }
  }

  for (const authorization of [
    undefined,
    {},
    { type: "ConnectionVault" },
    { referenceId: "shared_kusto" },
    { type: "None", referenceId: "shared_kusto" },
    { type: "ConnectionVault", referenceId: 42 },
    { type: "ConnectionVault", referenceId: "shared_kusto", unknown: "private-value" },
  ]) {
    it(`rejects invalid OpenAPI authorization: ${JSON.stringify(authorization)}`, () => {
      const json = previewManifest({
        agentConnectors: [connector({ openApiConnector: { authorization } })],
      });
      assert.throws(
        () => TeamsManifestConverter.jsonToManifest(json),
        /agentConnectors\[0\]\.toolSource\.openApiConnector\.authorization/
      );
    });
  }

  for (const kind of ["remoteMcpServer", "localMcpServer"]) {
    for (const type of [
      "None",
      "OAuthPluginVault",
      "ApiKeyPluginVault",
      "DynamicClientRegistration",
      "AzureKeyVault",
    ]) {
      it(`preserves ${kind} authorization ${type}`, async () => {
        assert.deepEqual(
          await validate(
            previewManifest({
              agentConnectors: [
                connector(
                  toolSource(kind, {
                    type,
                    ...(type === "None" ? {} : { referenceId: "test-reference" }),
                  })
                ),
              ],
            })
          ),
          []
        );
      });
    }

    it(`requires referenceId for ${kind} ConnectionVault during schema validation`, async () => {
      const errors = await validate(
        previewManifest({
          agentConnectors: [connector(toolSource(kind, { type: "ConnectionVault" }))],
        })
      );
      assert.include(errors.join("\n"), "referenceId");
    });

    it(`requires file when ${kind} declares mcpToolDescription`, () => {
      const source = toolSource(kind, { type: "None" });
      Object.assign(source[kind], { mcpToolDescription: {} });
      assert.throws(
        () =>
          TeamsManifestConverter.jsonToManifest(
            previewManifest({ agentConnectors: [connector(source)] })
          ),
        /mcpToolDescription\.file/
      );
    });
  }

  it("rejects unexpected OpenAPI connector properties", () => {
    assert.throws(
      () =>
        TeamsManifestConverter.jsonToManifest(
          previewManifest({
            agentConnectors: [
              connector({
                openApiConnector: {
                  authorization: { type: "ConnectionVault", referenceId: "shared_kusto" },
                  unknown: "private-value",
                },
              }),
            ],
          })
        ),
      /openApiConnector\.unknown/
    );
  });

  it("reads and validates a connector manifest through file utilities", async () => {
    const folder = await fs.mkdtemp(path.join(os.tmpdir(), "manifest-preview-"));
    try {
      const file = path.join(folder, "manifest.json");
      const json = previewManifest({
        agentConnectors: [
          connector(
            toolSource("openApiConnector", {
              type: "ConnectionVault",
              referenceId: "shared_kusto",
            })
          ),
        ],
      });
      await fs.writeFile(file, json);
      const [manifest, errors] = await AppManifestUtils.readAndValidateTeamsManifest(file);
      assert.deepEqual(errors, []);
      await AppManifestUtils.writeTeamsManifest(file, manifest);
      assert.deepEqual(await fs.readJson(file), JSON.parse(json));
    } finally {
      await fs.remove(folder);
    }
  });

  it("round trips the additional preview declarations and relaxed name requirement", async () => {
    const json = previewManifest({
      name: { short: "Test" },
      copilotAgents: { plugins: [{ id: "plugin", file: "plugin.json" }] },
      localAgents: [{ id: "agent", file: "localagent.json" }],
      localTools: [{ id: "tool", file: "localtool.json" }],
      thirdPartyLicenses: [
        {
          licenseName: "MIT",
          licenseUrl: "https://example.com/license",
          licenseDescription: "Test license",
        },
      ],
    });
    const manifest = TeamsManifestConverter.jsonToManifest(json);
    assert.deepEqual(JSON.parse(TeamsManifestConverter.manifestToJson(manifest)), JSON.parse(json));
    assert.deepEqual(await AppManifestUtils.validateAgainstSchema(manifest), []);
  });

  it("validates copilotExtensions and rejects simultaneous copilotAgents", async () => {
    const copilotExtensions = { plugins: [{ id: "plugin", file: "plugin.json" }] };
    assert.deepEqual(await validate(previewManifest({ copilotExtensions })), []);
    const errors = await validate(
      previewManifest({
        copilotExtensions,
        copilotAgents: { plugins: [{ id: "plugin", file: "plugin.json" }] },
      })
    );
    assert.isNotEmpty(errors);
  });

  it("rejects the removed static tab contentBotId property", () => {
    assert.throws(
      () =>
        TeamsManifestConverter.jsonToManifest(
          previewManifest({
            staticTabs: [
              {
                entityId: "home",
                scopes: ["personal"],
                contentBotId: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
              },
            ],
          })
        ),
      /staticTabs\[0\]\.contentBotId/
    );
  });
});
