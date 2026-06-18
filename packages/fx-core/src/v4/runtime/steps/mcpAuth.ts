// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { RegisteredStep, StepContext, StepParams } from "../../pipeline/runScaffoldPipeline";
import { deriveMcpServerName } from "../whitelist";

/**
 * The `mcp-auth/*` pipeline steps (ADR-0017 whitelist; domain-typed names,
 * INV-6) for the `da/mcp-server` create flow. Both read-modify-write a
 * render-phase file through `ctx.read` + `ctx.write`
 * (run-scaffold-pipeline AC-21); neither mutates a manifest, so neither routes
 * through `manifestWrapper`.
 *
 * Spec: docs/03-specs/scenarios/da/create-mcp-server.md
 *       (SCN-CREATE-MCP-05 yml injection, SCN-CREATE-MCP-06 env persistence)
 *
 * v4-owned (INV-7): a self-contained create-flow subset. The impure provision
 * fidelity (probed authorization/token URLs, DCR auto-discovery, the v1.12 →
 * v1.13 schema bump) is deferred — see the package `pipeline.json` comment and
 * scaffolding.backlog.md §1.
 */

const SOURCE = "Scaffold";

/** Engine step name `mcp-auth/inject-yml-action`. */
export const STEP_INJECT_YML_ACTION = "mcp-auth/inject-yml-action";

/** Engine step name `mcp-auth/persist-credential-env`. */
export const STEP_PERSIST_CREDENTIAL_ENV = "mcp-auth/persist-credential-env";

/** The create flow scaffolds the `dev` environment; this is its credential file. */
const CREATE_ENV_FILE = "env/.env.dev";

function systemError(name: string, message: string): SystemError {
  return new SystemError({ source: SOURCE, name, message });
}

/** Read a `with` value as a string, or `undefined` if it is absent / non-string. */
function stringParam(params: StepParams, key: string): string | undefined {
  const value = params[key];
  return typeof value === "string" ? value : undefined;
}

/** The credential env-var / registration id for a server URL (e.g. `MCP_DA_AUTH_ID_APIGITHUBC`). */
function registrationId(serverUrl: string): string {
  return "MCP_DA_AUTH_ID_" + deriveMcpServerName(serverUrl).toUpperCase();
}

/** The registration action a given auth type uses. */
function authActionUses(authType: string): string {
  return authType === "entra-sso" ? "microsoftEntra/register" : "oauth/register";
}

/** Insert an action block as the first entry under the top-level `provision:` key. */
function injectUnderProvision(yml: string, block: string): string {
  const lines = yml.split("\n");
  const index = lines.findIndex((line) => line.trim() === "provision:");
  if (index === -1) {
    return yml + "\nprovision:\n" + block + "\n";
  }
  lines.splice(index + 1, 0, block);
  return lines.join("\n");
}

/** True if the dotenv content already declares `name=…` (idempotency guard). */
function containsEnvVar(envContent: string, name: string): boolean {
  return envContent.split("\n").some((line) => line.startsWith(name + "="));
}

/**
 * `mcp-auth/inject-yml-action` — inject the auth registration action under
 * `provision` in the render-phase `m365agents.yml` (read-modify-write).
 */
export const mcpAuthInjectYmlAction: RegisteredStep = {
  validateParams(resolved: StepParams): string | undefined {
    if (stringParam(resolved, "ymlPath") === undefined) {
      return "missing string parameter 'ymlPath'";
    }
    if (stringParam(resolved, "authType") === undefined) {
      return "missing string parameter 'authType'";
    }
    if (stringParam(resolved, "mcpServerUrl") === undefined) {
      return "missing string parameter 'mcpServerUrl'";
    }
    return undefined;
  },
  apply(resolved: StepParams, ctx: StepContext): Result<void, FxError> {
    const ymlPath = stringParam(resolved, "ymlPath");
    const authType = stringParam(resolved, "authType");
    const serverUrl = stringParam(resolved, "mcpServerUrl");
    if (ymlPath === undefined || authType === undefined || serverUrl === undefined) {
      return err(systemError("McpAuthInjectYmlParams", "resolved parameters are not all strings"));
    }
    const current = ctx.read(ymlPath);
    if (current === undefined) {
      return err(
        systemError(
          "McpAuthYmlMissing",
          `Cannot inject the auth action: '${ymlPath}' was not produced by the render phase.`
        )
      );
    }
    const block = [
      `  - uses: ${authActionUses(authType)}`,
      `    with:`,
      `      name: ${deriveMcpServerName(serverUrl)}`,
      `    writeToEnvironmentFile:`,
      `      registrationId: ${registrationId(serverUrl)}`,
    ].join("\n");
    const injected = injectUnderProvision(current.toString("utf8"), block);
    ctx.write(ymlPath, Buffer.from(injected, "utf8"));
    return ok(undefined);
  },
};

/**
 * `mcp-auth/persist-credential-env` — append the `MCP_DA_AUTH_ID_<NS>`
 * credential reference to the render-phase `env/.env.dev` (read-modify-write,
 * idempotent).
 */
export const mcpAuthPersistCredentialEnv: RegisteredStep = {
  validateParams(resolved: StepParams): string | undefined {
    if (stringParam(resolved, "authType") === undefined) {
      return "missing string parameter 'authType'";
    }
    if (stringParam(resolved, "mcpServerUrl") === undefined) {
      return "missing string parameter 'mcpServerUrl'";
    }
    return undefined;
  },
  apply(resolved: StepParams, ctx: StepContext): Result<void, FxError> {
    const serverUrl = stringParam(resolved, "mcpServerUrl");
    if (serverUrl === undefined) {
      return err(systemError("McpAuthPersistParams", "resolved parameters are not all strings"));
    }
    const varName = registrationId(serverUrl);
    const current = ctx.read(CREATE_ENV_FILE);
    const base = current ? current.toString("utf8") : "";
    if (containsEnvVar(base, varName)) {
      return ok(undefined);
    }
    const separator = base.length > 0 && !base.endsWith("\n") ? "\n" : "";
    ctx.write(CREATE_ENV_FILE, Buffer.from(base + separator + varName + "=\n", "utf8"));
    return ok(undefined);
  },
};
