// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export type CapabilityKind = "step" | "provider" | "validator";

export interface CapabilityDeclaration {
  readonly id: string;
  readonly since: string;
  readonly outputs?: readonly { readonly name: string; readonly since?: string }[];
}

export const capabilityDeclarations = {
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
} satisfies Record<CapabilityKind, Readonly<Record<string, CapabilityDeclaration>>>;
