// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SubscriptionClient, TenantIdDescription } from "@azure/arm-resources-subscriptions";
import { TokenCredential } from "@azure/core-auth";
import * as vscode from "vscode";
import * as azureEnv from "@azure/ms-rest-azure-env";
import { AzureScopes, featureFlagManager, FeatureFlags } from "@microsoft/teamsfx-core";
import { LoginFailureError } from "./codeFlowLogin";
import { Environment } from "@azure/ms-rest-azure-env";

export const Microsoft = "microsoft";

// Licensed under the MIT license.
export class VSCodeAzureSubscriptionProvider {
  private async getSubscriptionClient(
    tenantId?: string,
    scopes?: string[]
  ): Promise<{
    client: SubscriptionClient;
    credential: TokenCredential;
    authentication: AzureAuthentication;
  }> {
    const armSubs = await import("@azure/arm-resources-subscriptions");
    const session = await getSessionFromVSCode(scopes, tenantId, {
      createIfNone: false,
      silent: true,
    });
    if (!session) {
      return Promise.reject(LoginFailureError());
    }

    const credential: TokenCredential = {
      // eslint-disable-next-line @typescript-eslint/require-await
      getToken: async () => {
        return {
          token: session.accessToken,
          expiresOnTimestamp: 0,
        };
      },
    };

    const configuredAzureEnv = getConfiguredAzureEnv();
    const endpoint = configuredAzureEnv.resourceManagerEndpointUrl;

    return {
      client: new armSubs.SubscriptionClient(credential, { endpoint }),
      credential: credential,
      authentication: {
        getSession: () => session,
      },
    };
  }

  /**
   * Gets a list of tenants available to the user.
   *
   * @returns A list of tenants.
   */
  public async getTenants(): Promise<TenantIdDescription[]> {
    const { client } = await this.getSubscriptionClient(undefined, AzureScopes);

    const results: TenantIdDescription[] = [];

    for await (const tenant of client.tenants.list()) {
      results.push(tenant);
    }

    return results;
  }

  /**
   * Gets a list of Azure subscriptions available to the user.
   */
  public async getSubscriptions(tenantId?: string): Promise<AzureSubscription[]> {
    const results: AzureSubscription[] = [];

    if (featureFlagManager.getBooleanValue(FeatureFlags.MultiTenant) && tenantId) {
      results.push(...(await this.getSubscriptionsForTenant(tenantId)));
    } else {
      for (const tenant of await this.getTenants()) {
        try {
          // Get the list of tenants
          // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
          const tenantId = tenant.tenantId!;

          // For each tenant, get the list of subscriptions
          results.push(...(await this.getSubscriptionsForTenant(tenantId)));
        } catch (e) {}
      }
    }

    const sortSubscriptions = (subscriptions: AzureSubscription[]): AzureSubscription[] =>
      subscriptions.sort((a, b) => a.name.localeCompare(b.name));
    return sortSubscriptions(results);
  }

  /**
   * Gets the subscriptions for a given tenant.
   *
   * @param tenantId The tenant ID to get subscriptions for.
   *
   * @returns The list of subscriptions for the tenant.
   */
  private async getSubscriptionsForTenant(tenantId: string): Promise<AzureSubscription[]> {
    const { client, credential, authentication } = await this.getSubscriptionClient(
      tenantId,
      AzureScopes
    );
    const environment = getConfiguredAzureEnv();

    const subscriptions: AzureSubscription[] = [];

    for await (const subscription of client.subscriptions.list()) {
      subscriptions.push({
        authentication: authentication,
        environment: environment,
        credential: credential,
        isCustomCloud: environment.isCustomCloud,
        /* eslint-disable @typescript-eslint/no-non-null-assertion */
        name: subscription.displayName!,
        subscriptionId: subscription.subscriptionId!,
        /* eslint-enable @typescript-eslint/no-non-null-assertion */
        tenantId: tenantId,
      });
    }

    return subscriptions;
  }
}

export async function getSessionFromVSCode(
  scopes?: string | string[],
  tenantId?: string,
  options?: vscode.AuthenticationGetSessionOptions
): Promise<vscode.AuthenticationSession | undefined> {
  return await vscode.authentication.getSession(Microsoft, getScopes(scopes, tenantId), options);
}

function ensureEndingSlash(value: string): string {
  return value.endsWith("/") ? value : `${value}/`;
}

function getResourceScopes(scopes?: string | string[]): string[] {
  if (scopes === undefined || scopes === "" || scopes.length === 0) {
    scopes = ensureEndingSlash(getConfiguredAzureEnv().managementEndpointUrl);
  }
  return Array.from(new Set<string>(scopes));
}

function addTenantIdScope(scopes: string[], tenantId: string): string[] {
  const scopeSet = new Set<string>(scopes);
  scopeSet.add(`VSCODE_TENANT:${tenantId}`);
  return Array.from(scopeSet);
}

function getScopes(scopes: string | string[] | undefined, tenantId?: string): string[] {
  let scopeArr = getResourceScopes(scopes);
  if (tenantId) {
    scopeArr = addTenantIdScope(scopeArr, tenantId);
  }
  return scopeArr;
}

/**
 * Represents a means of obtaining authentication data for an Azure subscription.
 */
export interface AzureAuthentication {
  /**
   * Gets a VS Code authentication session for an Azure subscription.
   *
   * @param scopes - The scopes for which the authentication is needed.
   *
   * @returns A VS Code authentication session or undefined, if none could be obtained.
   */
  getSession(scopes?: string[]): vscode.ProviderResult<vscode.AuthenticationSession>;
}

/**
 * Gets the configured Azure environment.
 *
 * @returns The configured Azure environment from the settings in the built-in authentication provider extension
 */
export function getConfiguredAzureEnv(): azureEnv.Environment & { isCustomCloud: boolean } {
  return {
    ...azureEnv.Environment.get(azureEnv.Environment.AzureCloud.name),
    isCustomCloud: false,
  };
}

/**
 * Represents an Azure subscription.
 */
export interface AzureSubscription {
  /**
   * Access to the authentication session associated with this subscription.
   */
  readonly authentication: AzureAuthentication;

  /**
   * The Azure environment to which this subscription belongs.
   */
  readonly environment: Environment;

  /**
   * Whether this subscription belongs to a custom cloud.
   */
  readonly isCustomCloud: boolean;

  /**
   * The display name of this subscription.
   */
  readonly name: string;

  /**
   * The ID of this subscription.
   */
  readonly subscriptionId: string;

  /**
   * The ID of the tenant to which this subscription belongs.
   */
  readonly tenantId: string;

  /**
   * The credential for authentication to this subscription. Compatible with Azure track 2 SDKs.
   */
  readonly credential: TokenCredential;
}
