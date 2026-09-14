// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "vitest";
import { CURRENT_V4_ENGINE_VERSION } from "../../../src/v4/engineVersion";
import { createDefaultCreateOptionsProviders } from "../../../src/v4/providers/createOptionsProviders";
import { STEP_REGISTRY } from "../../../src/v4/runtime/runtimeRegistry";
import {
  templateCapabilities,
  templateCapabilityFloor,
  templateCapabilityOutputs,
} from "../../../src/v4/validation/capabilityCatalog";
import { createDefaultCreateInputValidators } from "../../../src/v4/validators/createInputValidators";

function sorted(values: Iterable<string>): string[] {
  return [...values].sort();
}

async function loadDeclarations() {
  try {
    return await import("../../../src/v4/capabilities/declarations");
  } catch (error) {
    assert.fail(`Pure capability declarations must load: ${String(error)}`);
  }
}

describe("v4/validation/capabilityCatalog", () => {
  it("API-02/OWN-03: declarations pin literal capability identities, versions and outputs", async () => {
    const { capabilityDeclarations } = await loadDeclarations();

    assert.deepEqual(capabilityDeclarations, {
      step: {
        requireEmptyTarget: { id: "require-empty-target", since: "5.20.0" },
        registerPluginManifest: { id: "da-action/register-plugin-manifest", since: "5.20.0" },
        setSensitivityLabel: { id: "da/set-sensitivity-label", since: "6.11.0" },
        injectYmlAction: { id: "mcp-auth/inject-yml-action", since: "5.20.0" },
        persistCredentialEnv: { id: "mcp-auth/persist-credential-env", since: "5.20.0" },
        materializeLocalServers: { id: "mcp-local/materialize-servers", since: "5.20.0" },
        materializeStaticMcpTools: { id: "mcp-static/materialize-tools", since: "5.20.0" },
        unifyProjectId: { id: "metaos/unify-project-id", since: "5.20.0" },
        upgradeExistingProject: { id: "metaos/upgrade-existing-project", since: "5.20.0" },
        importExistingOfficeAddinProject: {
          id: "officeaddin/import-existing-project",
          since: "5.20.0",
        },
        generateOpenApiPluginFiles: { id: "openapi/generate-plugin-files", since: "5.20.0" },
        generateTeamsAiCustomApiFiles: {
          id: "openapi/generate-teams-ai-custom-api-files",
          since: "5.20.0",
        },
      },
      provider: {
        createLanguages: { id: "create.languages", since: "6.12.0" },
        mcpServerTypes: {
          id: "mcp.serverTypes",
          since: "5.20.0",
          outputs: [{ name: "catalog" }],
        },
        mcpLocalServers: { id: "mcp.localServers", since: "5.20.0" },
        mcpTools: { id: "mcp.tools", since: "5.20.0", outputs: [{ name: "toolsJson" }] },
        openApiSearch: { id: "openapi.search", since: "5.20.0" },
        openApiOperations: {
          id: "openapi.operations",
          since: "5.20.0",
          outputs: [{ name: "apiSpecLocation", since: "6.12.0" }],
        },
        openApiTeamsAiOperations: {
          id: "openapi.teamsAiOperations",
          since: "6.13.0",
          outputs: [{ name: "apiSpecLocation", since: "6.13.0" }],
        },
      },
      validator: {
        uri: { id: "uri", since: "5.20.0" },
        openApiUrl: { id: "openapiUrl", since: "5.20.0" },
        graphConnectorName: { id: "graphConnectorName", since: "5.20.0" },
        graphConnectorConnectionId: { id: "graphConnectorConnectionId", since: "5.20.0" },
        mcpOauthClientIdRequired: { id: "mcp.oauthClientIdRequired", since: "6.11.0" },
        mcpOauthClientSecretRequired: { id: "mcp.oauthClientSecretRequired", since: "6.11.0" },
        mcpEntraClientIdRequired: { id: "mcp.entraClientIdRequired", since: "6.11.0" },
        mcpServerUrl: { id: "mcp.serverUrl", since: "6.11.0" },
      },
    });

    for (const kind of ["step", "provider", "validator"] satisfies Array<
      keyof typeof capabilityDeclarations
    >) {
      const declarations = Object.values(capabilityDeclarations[kind]);
      assert.deepEqual(
        templateCapabilities(kind),
        declarations.map((declaration) => declaration.id)
      );
      for (const declaration of declarations) {
        assert.equal(templateCapabilityFloor(kind, declaration.id), declaration.since);
      }
    }
  });

  it("OWN-03: default bindings and output schemas agree with declarations without registering the guard", async () => {
    const { capabilityDeclarations } = await loadDeclarations();
    const providers = createDefaultCreateOptionsProviders(
      async () => ({ tools: [], requiresAuth: false }),
      async () => []
    );
    const validators = createDefaultCreateInputValidators();

    assert.isFalse(STEP_REGISTRY.has("require-empty-target"));
    assert.deepEqual(
      sorted(STEP_REGISTRY.keys()),
      sorted(
        Object.values(capabilityDeclarations.step)
          .map((declaration) => declaration.id)
          .filter((id) => id !== "require-empty-target")
      )
    );
    assert.deepEqual(
      sorted(Object.keys(providers)),
      sorted(Object.values(capabilityDeclarations.provider).map((declaration) => declaration.id))
    );
    assert.deepEqual(
      sorted(Object.keys(validators)),
      sorted(Object.values(capabilityDeclarations.validator).map((declaration) => declaration.id))
    );
    for (const declaration of Object.values(capabilityDeclarations.provider)) {
      const outputs =
        "outputs" in declaration ? declaration.outputs.map((output) => output.name) : [];
      assert.deepEqual(providers[declaration.id].derivedSchema ?? [], outputs, declaration.id);
      assert.deepEqual(templateCapabilityOutputs("provider", declaration.id), outputs);
    }
  });

  it("OWN-04: output floors override only on consumption and unknown outputs retain the capability floor", () => {
    assert.equal(templateCapabilityFloor("provider", "openapi.operations"), "5.20.0");
    assert.equal(
      templateCapabilityFloor("provider", "openapi.operations", "apiSpecLocation"),
      "6.12.0"
    );
    assert.equal(templateCapabilityFloor("provider", "openapi.operations", "unknown"), "5.20.0");
    assert.equal(templateCapabilityFloor("provider", "openapi.operations", ""), "5.20.0");
    assert.equal(templateCapabilityFloor("provider", "mcp.serverTypes", "catalog"), "5.20.0");
    assert.equal(templateCapabilityFloor("provider", "mcp.tools", "toolsJson"), "5.20.0");
    assert.deepEqual(templateCapabilityOutputs("provider", "openapi.operations"), [
      "apiSpecLocation",
    ]);
    assert.deepEqual(templateCapabilityOutputs("provider", "create.languages"), []);
    for (const kind of ["step", "provider", "validator"] satisfies Array<
      Parameters<typeof templateCapabilityFloor>[0]
    >) {
      assert.isUndefined(templateCapabilityFloor(kind, "unknown"));
      assert.isUndefined(templateCapabilityFloor(kind, "unknown", "apiSpecLocation"));
      assert.isUndefined(templateCapabilityFloor(kind, "toString"));
      assert.deepEqual(templateCapabilityOutputs(kind, "unknown"), []);
    }
  });

  it("API-02: Teams AI listing and source start at engine 6.13 independently of template versions", () => {
    assert.equal(CURRENT_V4_ENGINE_VERSION, "6.13.0");
    assert.equal(templateCapabilityFloor("provider", "openapi.teamsAiOperations"), "6.13.0");
    assert.equal(
      templateCapabilityFloor("provider", "openapi.teamsAiOperations", "apiSpecLocation"),
      "6.13.0"
    );
    assert.deepEqual(templateCapabilityOutputs("provider", "openapi.teamsAiOperations"), [
      "apiSpecLocation",
    ]);
  });

  it("AC-23/24: every runtime step has exactly one source-owned capability floor", () => {
    const runtimeSteps = new Set(["require-empty-target", ...STEP_REGISTRY.keys()]);

    assert.deepEqual(sorted(templateCapabilities("step")), sorted(runtimeSteps));
  });

  it("AC-23/24: every default provider has exactly one source-owned capability floor", () => {
    const providers = createDefaultCreateOptionsProviders(
      async () => ({ tools: [], requiresAuth: false }),
      async () => []
    );

    assert.deepEqual(sorted(templateCapabilities("provider")), sorted(Object.keys(providers)));
  });

  it("INPUT-25: provider derived schemas exactly match source-owned capability outputs", () => {
    const providers = createDefaultCreateOptionsProviders(
      async () => ({ tools: [], requiresAuth: false }),
      async () => []
    );

    for (const [id, provider] of Object.entries(providers)) {
      assert.deepEqual(
        sorted(templateCapabilityOutputs("provider", id)),
        sorted(provider.derivedSchema ?? []),
        id
      );
    }
  });

  it("AC-23/24: every default validator has exactly one source-owned capability floor", () => {
    const validators = createDefaultCreateInputValidators();

    assert.deepEqual(sorted(templateCapabilities("validator")), sorted(Object.keys(validators)));
  });
});
