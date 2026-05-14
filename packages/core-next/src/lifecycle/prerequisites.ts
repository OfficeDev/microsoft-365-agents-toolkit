// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, ok, err } from "neverthrow";
import type { AtkContext } from "../core/context";
import type { AtkError } from "../core/error";
import { userError } from "../core/error";
import type {
  M365TenantInfo,
  AzureAccountInfo,
  SubscriptionInfo,
  ResourceGroupInfo,
} from "./types";

const SOURCE = "lifecycle/prerequisites";

// M365 Graph API scope for token acquisition
const M365_SCOPE = "https://graph.microsoft.com/.default";

/**
 * Ensure the user is authenticated to M365 and retrieve tenant info.
 *
 * Acquires an M365 token via the context's token provider. If the user
 * is not signed in, the provider triggers the platform's login flow.
 *
 * This is a composable prerequisite — consumers can call it directly
 * or let the lifecycle operations call it automatically.
 */
export async function ensureM365Auth(ctx: AtkContext): Promise<Result<M365TenantInfo, AtkError>> {
  const tokenResult = await ctx.auth.m365TokenProvider.getJsonObject({
    scopes: [M365_SCOPE],
    showDialog: true,
  });

  if (tokenResult.isErr()) {
    return err(
      userError("M365AuthFailed", "Failed to sign in to Microsoft 365.", {
        source: SOURCE,
      })
    );
  }

  const claims = tokenResult.value;
  const tenantId = (claims.tid as string) ?? (claims.tenantId as string) ?? "";
  const displayName =
    (claims.name as string) ?? (claims.preferred_username as string) ?? (claims.upn as string);

  if (!tenantId) {
    return err(
      userError("M365TenantNotFound", "Could not determine M365 tenant ID from token.", {
        source: SOURCE,
      })
    );
  }

  return ok({ tenantId, displayName });
}

/**
 * Ensure the user is authenticated to Azure.
 *
 * Triggers the Azure login flow if needed and returns basic account info.
 */
export async function ensureAzureAuth(
  ctx: AtkContext
): Promise<Result<AzureAccountInfo, AtkError>> {
  const credential = await ctx.auth.azureAccountProvider.getIdentityCredentialAsync(true);

  if (!credential) {
    return err(
      userError("AzureAuthFailed", "Failed to sign in to Azure.", {
        source: SOURCE,
      })
    );
  }

  const accountInfo = ctx.auth.azureAccountProvider.getAccountInfo?.();
  return ok({
    accountId: accountInfo?.["email"] ?? accountInfo?.["name"],
    tenantId: accountInfo?.["tenantId"],
  });
}

/**
 * Ensure an Azure subscription is selected.
 *
 * If `AZURE_SUBSCRIPTION_ID` is already in envMap, returns it.
 * Otherwise prompts the user to select from their available subscriptions.
 */
export async function ensureSubscription(
  ctx: AtkContext,
  envMap: Map<string, string>
): Promise<Result<SubscriptionInfo, AtkError>> {
  // Already resolved?
  const existing = envMap.get("AZURE_SUBSCRIPTION_ID");
  if (existing) {
    return ok({
      subscriptionId: existing,
      subscriptionName: envMap.get("AZURE_SUBSCRIPTION_NAME") ?? "",
      tenantId: envMap.get("AZURE_TENANT_ID") ?? "",
    });
  }

  // List subscriptions
  const subs = await ctx.auth.azureAccountProvider.listSubscriptions();
  if (subs.length === 0) {
    return err(
      userError("NoAzureSubscription", "No Azure subscriptions found for the current account.", {
        source: SOURCE,
      })
    );
  }

  // Single sub → auto-select
  if (subs.length === 1) {
    const sub = subs[0];
    envMap.set("AZURE_SUBSCRIPTION_ID", sub.subscriptionId);
    envMap.set("AZURE_SUBSCRIPTION_NAME", sub.subscriptionName);
    envMap.set("AZURE_TENANT_ID", sub.tenantId);
    return ok(sub);
  }

  // Multiple → prompt user
  const options = subs.map((s) => ({
    id: s.subscriptionId,
    label: s.subscriptionName,
    description: s.subscriptionId,
  }));

  const selectResult = await ctx.ui.selectOption({
    name: "subscription",
    title: "Select an Azure subscription",
    options,
  });

  if (selectResult.isErr()) {
    return err(
      userError("UserCancelled", "Azure subscription selection was cancelled.", {
        source: SOURCE,
      })
    );
  }

  if (selectResult.value.type !== "success" || !selectResult.value.result) {
    return err(
      userError("UserCancelled", "Azure subscription selection was cancelled.", {
        source: SOURCE,
      })
    );
  }

  const selectedId =
    typeof selectResult.value.result === "string"
      ? selectResult.value.result
      : (selectResult.value.result as { id: string }).id;

  const selected = subs.find((s) => s.subscriptionId === selectedId);
  if (!selected) {
    return err(
      userError("SubscriptionNotFound", `Subscription "${selectedId}" not found.`, {
        source: SOURCE,
      })
    );
  }

  envMap.set("AZURE_SUBSCRIPTION_ID", selected.subscriptionId);
  envMap.set("AZURE_SUBSCRIPTION_NAME", selected.subscriptionName);
  envMap.set("AZURE_TENANT_ID", selected.tenantId);
  await ctx.auth.azureAccountProvider.setSubscription(selected.subscriptionId);

  return ok(selected);
}

/**
 * Ensure an Azure resource group exists (or prompt to create one).
 *
 * If `AZURE_RESOURCE_GROUP_NAME` is already in envMap, returns it.
 * Otherwise prompts the user to enter a name and location.
 */
export async function ensureResourceGroup(
  ctx: AtkContext,
  envMap: Map<string, string>,
  subscriptionId: string,
  projectName: string,
  envName: string
): Promise<Result<ResourceGroupInfo, AtkError>> {
  const existing = envMap.get("AZURE_RESOURCE_GROUP_NAME");
  if (existing) {
    return ok({
      name: existing,
      location: envMap.get("AZURE_RESOURCE_GROUP_LOCATION") ?? "",
      isNew: false,
    });
  }

  // Generate default resource group name
  const suffix = ensureResourceSuffix(envMap);
  const safeName = projectName.replace(/[^a-zA-Z0-9]/g, "");
  const defaultRgName = `rg-${safeName}${suffix}-${envName}`;

  const nameResult = await ctx.ui.inputText({
    name: "resourceGroupName",
    title: "Enter Azure resource group name",
    default: defaultRgName,
  });

  if (nameResult.isErr() || nameResult.value.type !== "success" || !nameResult.value.result) {
    return err(
      userError("UserCancelled", "Resource group name input was cancelled.", {
        source: SOURCE,
      })
    );
  }

  const rgName = nameResult.value.result;

  const locationResult = await ctx.ui.inputText({
    name: "resourceGroupLocation",
    title: "Enter Azure resource group location",
    default: "centralus",
  });

  if (
    locationResult.isErr() ||
    locationResult.value.type !== "success" ||
    !locationResult.value.result
  ) {
    return err(
      userError("UserCancelled", "Resource group location input was cancelled.", {
        source: SOURCE,
      })
    );
  }

  const location = locationResult.value.result;

  envMap.set("AZURE_RESOURCE_GROUP_NAME", rgName);
  envMap.set("AZURE_RESOURCE_GROUP_LOCATION", location);

  return ok({ name: rgName, location, isNew: true });
}

/**
 * Ensure RESOURCE_SUFFIX is set in envMap. Generates a 6-char random suffix if missing.
 */
export function ensureResourceSuffix(envMap: Map<string, string>): string {
  let suffix = envMap.get("RESOURCE_SUFFIX");
  if (!suffix) {
    suffix = randomAlphanumeric(6);
    envMap.set("RESOURCE_SUFFIX", suffix);
  }
  return suffix;
}

/**
 * Show a confirmation dialog before provisioning Azure resources.
 *
 * Displays the target environment, M365 tenant, and Azure subscription.
 */
export async function confirmProvision(
  ctx: AtkContext,
  envName: string,
  m365Info?: M365TenantInfo,
  azureInfo?: SubscriptionInfo
): Promise<Result<void, AtkError>> {
  const parts: string[] = [`Environment: ${envName}`];
  if (m365Info) {
    parts.push(`M365 tenant: ${m365Info.displayName ?? m365Info.tenantId}`);
  }
  if (azureInfo) {
    parts.push(`Azure subscription: ${azureInfo.subscriptionName} (${azureInfo.subscriptionId})`);
  }

  const message = `Provision resources?\n\n${parts.join("\n")}`;
  const confirmResult = await ctx.ui.confirm?.({
    name: "confirmProvision",
    title: message,
  });

  if (!confirmResult || confirmResult.isErr()) {
    return err(
      userError("UserCancelled", "Provision was cancelled by the user.", { source: SOURCE })
    );
  }

  if (confirmResult.value.type !== "success" || confirmResult.value.result !== true) {
    return err(
      userError("UserCancelled", "Provision was cancelled by the user.", { source: SOURCE })
    );
  }

  return ok(undefined);
}

/**
 * Show a confirmation dialog before deploying.
 * Skipped for local/test/playground environments.
 */
export async function confirmDeploy(
  ctx: AtkContext,
  envName: string
): Promise<Result<void, AtkError>> {
  const localEnvs = ["local", "testtool", "playground", "sandbox"];
  if (localEnvs.includes(envName.toLowerCase())) {
    return ok(undefined);
  }

  const confirmResult = await ctx.ui.confirm?.({
    name: "confirmDeploy",
    title: `Deploy to environment "${envName}"?`,
  });

  if (!confirmResult || confirmResult.isErr()) {
    return err(userError("UserCancelled", "Deploy was cancelled by the user.", { source: SOURCE }));
  }

  if (confirmResult.value.type !== "success" || confirmResult.value.result !== true) {
    return err(userError("UserCancelled", "Deploy was cancelled by the user.", { source: SOURCE }));
  }

  return ok(undefined);
}

function randomAlphanumeric(len: number): string {
  const chars = "abcdefghijklmnopqrstuvwxyz0123456789"; // eslint-disable-line no-secrets/no-secrets
  let result = "";
  for (let i = 0; i < len; i++) {
    result += chars[Math.floor(Math.random() * chars.length)];
  }
  return result;
}
