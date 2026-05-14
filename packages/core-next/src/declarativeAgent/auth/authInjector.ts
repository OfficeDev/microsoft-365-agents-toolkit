// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs";
import { Result, ok, err } from "neverthrow";
import type { AtkError } from "../../core/error";
import { userError, systemError } from "../../core/error";
import type { AuthActionInjectResult } from "../types";

/**
 * Inject an OAuth registration action into a teamsapp.yml file.
 *
 * Parses the YAML, locates the `provision` section, and inserts an
 * `oauth/register` action after the `teamsApp/create` entry.
 *
 * @param ymlPath            Absolute path to the teamsapp.yml / teamsapp.local.yml.
 * @param authName           Human-readable auth scheme name (e.g. "myOAuth").
 * @param specRelativePath   Relative path to the OpenAPI spec (from project root).
 * @param isMicrosoftEntra   Whether this is a Microsoft Entra (AAD) identity provider.
 * @param enablePKCE         Whether to enable PKCE for the OAuth flow.
 * @param registrationId     Optional pre-existing registrationId env var name.
 */
export async function injectOAuthAction(
  ymlPath: string,
  authName: string,
  specRelativePath: string,
  isMicrosoftEntra: boolean,
  enablePKCE?: boolean,
  registrationId?: string
): Promise<Result<AuthActionInjectResult, AtkError>> {
  try {
    // Dynamic import to avoid hard dep on `yaml` for consumers that don't need it
    const { parseDocument } = await import("yaml");
    const content = await fs.promises.readFile(ymlPath, "utf-8");
    const doc = parseDocument(content);
    const provision = doc.get("provision") as any;

    if (!provision) {
      return err(
        userError("MissingProvisionSection", "YAML file has no 'provision' section.", {
          source: "declarativeAgent/auth",
        })
      );
    }

    const actionName = "oauth/register";

    // Check if an action with same name + spec already exists
    if (hasAction(provision, actionName, authName, specRelativePath)) {
      return ok({ defaultRegistrationIdEnvName: undefined, registrationIdEnvName: registrationId });
    }

    const teamsAppIdEnvName = getTeamsAppIdEnvName(provision) ?? "TEAMS_APP_ID";
    const existingEnvNames = collectEnvNames(provision);

    const envName =
      registrationId ?? findNextAvailableEnv("OAUTH2_CONFIGURATION_ID", existingEnvNames);

    const action: Record<string, unknown> = {
      uses: actionName,
      with: {
        name: authName,
        appId: `\${{${teamsAppIdEnvName}}}`,
        apiSpecPath: specRelativePath,
        flow: "authorizationCode",
        ...(enablePKCE && { isPKCEEnabled: true }),
        ...(isMicrosoftEntra && { identityProvider: "MicrosoftEntra" }),
      },
      writeToEnvironmentFile: {
        configurationId: envName,
        ...(isMicrosoftEntra && {
          applicationIdUri: `${authName}_APPLICATION_ID_URI`,
        }),
      },
    };

    insertAfterTeamsAppCreate(provision, doc.createNode(action));
    await fs.promises.writeFile(ymlPath, doc.toString(), "utf-8");

    return ok({
      defaultRegistrationIdEnvName: "OAUTH2_CONFIGURATION_ID",
      registrationIdEnvName: envName,
    });
  } catch (e) {
    return err(
      systemError("InjectOAuthActionFailed", `Failed to inject OAuth action: ${e}`, {
        source: "declarativeAgent/auth",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

/**
 * Inject an API-key registration action into a teamsapp.yml file.
 */
export async function injectApiKeyAction(
  ymlPath: string,
  authName: string,
  specRelativePath: string,
  registrationId?: string
): Promise<Result<AuthActionInjectResult, AtkError>> {
  try {
    const { parseDocument } = await import("yaml");
    const content = await fs.promises.readFile(ymlPath, "utf-8");
    const doc = parseDocument(content);
    const provision = doc.get("provision") as any;

    if (!provision) {
      return err(
        userError("MissingProvisionSection", "YAML file has no 'provision' section.", {
          source: "declarativeAgent/auth",
        })
      );
    }

    const actionName = "apiKey/register";

    if (hasAction(provision, actionName, authName, specRelativePath)) {
      return ok({ defaultRegistrationIdEnvName: undefined, registrationIdEnvName: registrationId });
    }

    const teamsAppIdEnvName = getTeamsAppIdEnvName(provision) ?? "TEAMS_APP_ID";
    const existingEnvNames = collectEnvNames(provision);

    const envName =
      registrationId ?? findNextAvailableEnv("APIKEY_REGISTRATION_ID", existingEnvNames);

    const action: Record<string, unknown> = {
      uses: actionName,
      with: {
        name: authName,
        appId: `\${{${teamsAppIdEnvName}}}`,
        apiSpecPath: specRelativePath,
      },
      writeToEnvironmentFile: {
        registrationId: envName,
      },
    };

    insertAfterTeamsAppCreate(provision, doc.createNode(action));
    await fs.promises.writeFile(ymlPath, doc.toString(), "utf-8");

    return ok({
      defaultRegistrationIdEnvName: "APIKEY_REGISTRATION_ID",
      registrationIdEnvName: envName,
    });
  } catch (e) {
    return err(
      systemError("InjectApiKeyActionFailed", `Failed to inject API key action: ${e}`, {
        source: "declarativeAgent/auth",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

/**
 * Inject an OAuth action specifically for MCP server auth.
 */
export async function injectMCPOAuthAction(
  ymlPath: string,
  authName: string,
  registrationIdEnvName: string,
  mcpServerUrl: string,
  oauthUrls?: { authorizationUrl: string; tokenUrl: string; refreshUrl?: string },
  isMicrosoftEntra?: boolean
): Promise<Result<AuthActionInjectResult, AtkError>> {
  try {
    const { parseDocument } = await import("yaml");
    const content = await fs.promises.readFile(ymlPath, "utf-8");
    const doc = parseDocument(content);
    const provision = doc.get("provision") as any;

    if (!provision) {
      return err(
        userError("MissingProvisionSection", "YAML file has no 'provision' section.", {
          source: "declarativeAgent/auth",
        })
      );
    }

    const teamsAppIdEnvName = getTeamsAppIdEnvName(provision) ?? "TEAMS_APP_ID";

    const action: Record<string, unknown> = {
      uses: "oauth/register",
      with: {
        name: authName,
        appId: `\${{${teamsAppIdEnvName}}}`,
        flow: "authorizationCode",
        baseUrl: mcpServerUrl,
        ...(isMicrosoftEntra
          ? { identityProvider: "MicrosoftEntra" }
          : { identityProvider: "Custom" }),
        ...(oauthUrls && {
          authorizationUrl: oauthUrls.authorizationUrl,
          tokenUrl: oauthUrls.tokenUrl,
          ...(oauthUrls.refreshUrl && { refreshUrl: oauthUrls.refreshUrl }),
        }),
      },
      writeToEnvironmentFile: {
        configurationId: registrationIdEnvName,
      },
    };

    insertAfterTeamsAppCreate(provision, doc.createNode(action));
    await fs.promises.writeFile(ymlPath, doc.toString(), "utf-8");

    return ok({
      defaultRegistrationIdEnvName: "OAUTH2_CONFIGURATION_ID",
      registrationIdEnvName,
    });
  } catch (e) {
    return err(
      systemError("InjectMCPOAuthActionFailed", `Failed to inject MCP OAuth action: ${e}`, {
        source: "declarativeAgent/auth",
        inner: e instanceof Error ? e : new Error(String(e)),
      })
    );
  }
}

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

function hasAction(
  provisionNode: any,
  actionUses: string,
  name: string,
  specPath: string
): boolean {
  if (!provisionNode?.items) return false;
  return provisionNode.items.some(
    (item: any) =>
      item.get("uses") === actionUses &&
      item.get("with")?.get("name") === name &&
      item.get("with")?.get("apiSpecPath") === specPath
  );
}

function getTeamsAppIdEnvName(provisionNode: any): string | undefined {
  if (!provisionNode?.items) return undefined;
  for (const item of provisionNode.items) {
    if (item.get("uses") === "teamsApp/create") {
      return item.get("writeToEnvironmentFile")?.get("teamsAppId") as string | undefined;
    }
  }
  return undefined;
}

function collectEnvNames(provisionNode: any): string[] {
  const names: string[] = [];
  if (!provisionNode?.items) return names;
  for (const item of provisionNode.items) {
    const writeEnv = item.get("writeToEnvironmentFile");
    if (writeEnv?.items) {
      for (const pair of writeEnv.items) {
        if (typeof pair.value?.value === "string") {
          names.push(pair.value.value);
        }
      }
    }
  }
  return names;
}

/**
 * Insert a YAML node right after the teamsApp/create action.
 */
function insertAfterTeamsAppCreate(provisionNode: any, newNode: unknown): void {
  let insertIdx = -1;
  for (let i = 0; i < provisionNode.items.length; i++) {
    if (provisionNode.items[i].get("uses") === "teamsApp/create") {
      insertIdx = i + 1;
      break;
    }
  }
  if (insertIdx === -1) {
    // No teamsApp/create found — append at the end
    provisionNode.add(newNode);
  } else {
    provisionNode.items.splice(insertIdx, 0, newNode);
  }
}

/**
 * Find next available env var name to avoid conflicts
 * (e.g. OAUTH2_CONFIGURATION_ID → OAUTH2_CONFIGURATION_ID_1).
 */
export function findNextAvailableEnv(baseName: string, existing: string[]): string {
  if (!existing.includes(baseName)) return baseName;
  let i = 1;
  while (existing.includes(`${baseName}_${i}`)) i++;
  return `${baseName}_${i}`;
}
