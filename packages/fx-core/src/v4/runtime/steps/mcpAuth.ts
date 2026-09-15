// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SystemError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { capabilityDeclarations } from "../../capabilities/declarations";
import { RegisteredStep } from "../../pipeline/runScaffoldPipeline";
import { defineStep } from "../../pipeline/defineStep";
import { stringParam, stringArrayParam } from "../../pipeline/stepParams";
import { injectMcpAuthAction, persistMcpAuthRegistrationEnv } from "../../mcp/mcpAuthScaffold";

/** MCP auth post-render steps for the create flow. See create-mcp-server scenario spec. */

const SOURCE = "Scaffold";

/** Engine step name `mcp-auth/inject-yml-action`. */
export const STEP_INJECT_YML_ACTION = capabilityDeclarations.step.injectYmlAction.id;

/** Engine step name `mcp-auth/persist-credential-env`. */
export const STEP_PERSIST_CREDENTIAL_ENV = capabilityDeclarations.step.persistCredentialEnv.id;

function systemError(name: string, message: string): SystemError {
  return new SystemError({ source: SOURCE, name, message });
}

/** Registered step for injecting the shared v4 MCP auth action into `m365agents.yml`. */
export const mcpAuthInjectYmlAction: RegisteredStep = defineStep({
  parse(resolved): Result<Parameters<typeof injectMcpAuthAction>[1], string> {
    const ymlPath = stringParam(resolved, "ymlPath");
    if (ymlPath === undefined) {
      return err("missing string parameter 'ymlPath'");
    }
    const authType = stringParam(resolved, "authType");
    if (authType === undefined) {
      return err("missing string parameter 'authType'");
    }
    const mcpServerUrl = stringParam(resolved, "mcpServerUrl");
    if (mcpServerUrl === undefined) {
      return err("missing string parameter 'mcpServerUrl'");
    }
    const optionalYmlPaths = stringArrayParam(resolved, "optionalYmlPaths");
    if (resolved.optionalYmlPaths !== undefined && optionalYmlPaths === undefined) {
      return err("parameter 'optionalYmlPaths' must be a string array");
    }
    return ok({
      ymlPath,
      authType,
      mcpServerUrl,
      optionalYmlPaths,
      credentialFields: {
        clientId: authType === "oauth" || authType === "entra-sso",
        clientSecret: authType === "oauth",
        scope: Boolean(stringParam(resolved, "oauthScopes")?.trim()),
        apiKey: authType === "bearer-token",
      },
    });
  },
  apply: (params, ctx) => injectMcpAuthAction(ctx, params),
  invalidParams: () =>
    systemError("McpAuthInjectParams", "resolved parameters are not all strings"),
});

/**
 * Registered step for writing the deterministic `MCP_DA_AUTH_ID_<NS>` registration placeholder into
 * `env/.env.dev`.
 */
export const mcpAuthPersistCredentialEnv: RegisteredStep = defineStep({
  parse(resolved): Result<Parameters<typeof persistMcpAuthRegistrationEnv>[1], string> {
    const authType = stringParam(resolved, "authType");
    if (authType === undefined) {
      return err("missing string parameter 'authType'");
    }
    const mcpServerUrl = stringParam(resolved, "mcpServerUrl");
    if (mcpServerUrl === undefined) {
      return err("missing string parameter 'mcpServerUrl'");
    }
    return ok({
      authType,
      mcpServerUrl,
      oauthClientId: stringParam(resolved, "oauthClientId"),
      oauthClientSecret: stringParam(resolved, "oauthClientSecret"),
      oauthScopes: stringParam(resolved, "oauthScopes"),
      entraClientId: stringParam(resolved, "entraClientId"),
      apiKey: stringParam(resolved, "apiKey"),
    });
  },
  apply: (params, ctx) => persistMcpAuthRegistrationEnv(ctx, params),
  invalidParams: () =>
    systemError("McpAuthPersistParams", "resolved parameters are not all strings"),
});
