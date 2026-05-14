// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SubscriptionClient } from "@azure/arm-subscriptions";
import { Configuration, LogLevel } from "@azure/msal-node";
import {
  type AuthenticationWWWAuthenticateRequest,
  type AzureAccountProvider,
  type FxError,
  type ITeamsFxTokenCredential,
  type Result,
  type SubscriptionInfo,
  ok,
  signedIn,
  signedOut,
  UserError,
} from "@microsoft/teamsfx-core-next";
import { CryptoCachePlugin, loadTenantId, saveTenantId } from "./cacheAccess";
import { CodeFlowLogin } from "./codeFlowLogin";
import { convertTokenToJson, checkIsOnline } from "./utils";
import { MSAL_CLIENT_ID, AZURE_CACHE_NAME, AzureScopes, MFA_CODE } from "./constants";

type LoginStatus = { status: string; token?: string; accountInfo?: Record<string, unknown> };

const scopes = ["https://management.core.windows.net/user_impersonation"];
const cachePlugin = new CryptoCachePlugin(AZURE_CACHE_NAME);

function getConfig(tenantId?: string): Configuration {
  const authority = tenantId
    ? `https://login.microsoftonline.com/${tenantId}`
    : "https://login.microsoftonline.com/organizations";

  const config: Configuration = {
    auth: {
      clientId: MSAL_CLIENT_ID,
      authority,
      clientCapabilities: ["CP1"],
    },
    system: {
      loggerOptions: {
        loggerCallback() {},
        piiLoggingEnabled: false,
        logLevel: LogLevel.Error,
      },
    },
    cache: { cachePlugin },
  };

  if (process.platform === "win32") {
    try {
      const { NativeBrokerPlugin } = require("@azure/msal-node-extensions");
      config.broker = { nativeBrokerPlugin: new NativeBrokerPlugin() };
    } catch {
      // NativeBrokerPlugin not available
    }
  }

  return config;
}

/**
 * TokenCredential wrapper around CodeFlowLogin for Azure SDK clients.
 */
class TeamsFxTokenCredential implements ITeamsFxTokenCredential {
  private codeFlowInstance: CodeFlowLogin;
  private tenantId = "";
  private silent = false;

  constructor(codeFlowInstance: CodeFlowLogin, silent = false) {
    this.codeFlowInstance = codeFlowInstance;
    this.silent = silent;
  }

  setTenantId(tenantId: string): void {
    this.tenantId = tenantId;
  }

  setSilent(silent: boolean): void {
    this.silent = silent;
  }

  async getToken(
    tokenScopes: string | string[] | AuthenticationWWWAuthenticateRequest
  ): Promise<{ token: string; expiresOnTimestamp: number } | null> {
    const tokenRes = await this.codeFlowInstance.getTokenByScopes(
      tokenScopes,
      !this.silent,
      this.tenantId || undefined
    );
    if (tokenRes.isOk()) {
      const json = convertTokenToJson(tokenRes.value);
      return {
        token: tokenRes.value,
        expiresOnTimestamp: ((json as Record<string, number>).exp ?? 0) * 1000,
      };
    }
    return null;
  }
}

/**
 * Azure account provider — interactive MSAL login with subscription listing.
 * Singleton.
 */
class AzureAccountManager implements AzureAccountProvider {
  private static instance: AzureAccountManager;
  private codeFlow: CodeFlowLogin;
  private teamsFxCredential: TeamsFxTokenCredential;
  private selectedTenantId: string | undefined;
  private selectedSubscriptionId: string | undefined;
  private selectedSubscriptionName: string | undefined;

  private constructor() {
    this.codeFlow = new CodeFlowLogin(scopes, getConfig(), 0, AZURE_CACHE_NAME);
    this.teamsFxCredential = new TeamsFxTokenCredential(this.codeFlow);
  }

  static getInstance(): AzureAccountManager {
    if (!AzureAccountManager.instance) {
      AzureAccountManager.instance = new AzureAccountManager();
    }
    return AzureAccountManager.instance;
  }

  getIdentityCredentialAsync(
    _showDialog?: boolean,
    authRequest?: AuthenticationWWWAuthenticateRequest
  ): Promise<ITeamsFxTokenCredential | undefined> {
    if (authRequest?.wwwAuthenticate) {
      throw new UserError({
        name: "MFARequired",
        message:
          "Multi-factor authentication required. Run `atk auth login azure` to re-authenticate interactively.",
        source: "login",
      });
    }
    return Promise.resolve(this.teamsFxCredential);
  }

  async switchTenant(tenantId: string): Promise<Result<ITeamsFxTokenCredential, FxError>> {
    await saveTenantId(AZURE_CACHE_NAME, tenantId);
    return ok(this.teamsFxCredential);
  }

  async getJsonObject(
    _showDialog?: boolean,
    tenantId?: string
  ): Promise<Record<string, unknown> | undefined> {
    const token = await this.codeFlow.getTokenByScopes(AzureScopes(), true, tenantId);
    if (token.isOk()) {
      return convertTokenToJson(token.value);
    }
    return undefined;
  }

  async signout(): Promise<boolean> {
    this.codeFlow.account = undefined;
    await this.codeFlow.logout();
    return true;
  }

  async getStatus(): Promise<LoginStatus> {
    if (!this.codeFlow.account) {
      await this.codeFlow.reloadCache();
    }
    if (this.codeFlow.account) {
      const tokenRes = await this.codeFlow.getTokenByScopes(scopes, false);
      if (tokenRes.isOk()) {
        const accountJson = convertTokenToJson(tokenRes.value);
        return { status: signedIn, token: tokenRes.value, accountInfo: accountJson };
      }
      if (await checkIsOnline()) {
        return { status: signedOut, token: undefined, accountInfo: undefined };
      }
      return {
        status: signedIn,
        token: undefined,
        accountInfo: { upn: this.codeFlow.account.username },
      };
    }
    return { status: signedOut, token: undefined, accountInfo: undefined };
  }

  async listSubscriptions(): Promise<SubscriptionInfo[]> {
    const arr: SubscriptionInfo[] = [];
    if (!this.teamsFxCredential) return arr;

    if (!this.selectedTenantId) {
      // Iterate all tenants available to the user
      const tenantClient = new SubscriptionClient(this.teamsFxCredential);
      const silentCred = new TeamsFxTokenCredential(this.codeFlow, true);
      const cachedTenantId = (await loadTenantId(AZURE_CACHE_NAME)) ?? undefined;

      for await (const page of tenantClient.tenants.list().byPage()) {
        for (const tenant of page) {
          if (cachedTenantId ? tenant.tenantId === cachedTenantId : tenant.tenantId) {
            try {
              silentCred.setTenantId(tenant.tenantId as string);
              const subClient = new SubscriptionClient(silentCred);
              for await (const subPage of subClient.subscriptions.list().byPage()) {
                for (const item of subPage) {
                  arr.push({
                    subscriptionId: item.subscriptionId!,
                    subscriptionName: item.displayName!,
                    tenantId: tenant.tenantId as string,
                  });
                }
              }
            } catch (error: unknown) {
              if (error instanceof Error && error.message.includes(MFA_CODE)) {
                console.log(
                  `Tenant ${tenant.tenantId} requires MFA. Use 'atk auth login azure --tenant ${tenant.tenantId}'.`
                );
              }
            }
          }
        }
      }
    } else {
      // Single-tenant mode
      this.teamsFxCredential.setTenantId(this.selectedTenantId);
      const subClient = new SubscriptionClient(this.teamsFxCredential);
      for await (const subPage of subClient.subscriptions.list().byPage()) {
        for (const item of subPage) {
          arr.push({
            subscriptionId: item.subscriptionId!,
            subscriptionName: item.displayName!,
            tenantId: this.selectedTenantId,
          });
        }
      }
    }
    return arr;
  }

  async setSubscription(subscriptionId: string): Promise<void> {
    if (subscriptionId === "") {
      this.selectedTenantId = undefined;
      this.selectedSubscriptionId = undefined;
      this.selectedSubscriptionName = undefined;
      return;
    }
    const list = await this.listSubscriptions();
    const found = list.find((s) => s.subscriptionId === subscriptionId);
    if (!found) {
      throw new UserError({
        name: "InvalidAzureSubscription",
        message: `Azure subscription '${subscriptionId}' not found.`,
        source: "login",
      });
    }
    this.selectedTenantId = found.tenantId;
    this.teamsFxCredential.setTenantId(found.tenantId);
    this.selectedSubscriptionId = found.subscriptionId;
    this.selectedSubscriptionName = found.subscriptionName;
  }

  getAccountInfo(): Record<string, string> | undefined {
    if (this.codeFlow.account) {
      return { username: this.codeFlow.account.username ?? "" };
    }
    return undefined;
  }

  async getSelectedSubscription(_triggerUI?: boolean): Promise<SubscriptionInfo | undefined> {
    if (this.codeFlow.account && this.selectedSubscriptionId) {
      return {
        subscriptionId: this.selectedSubscriptionId,
        tenantId: this.selectedTenantId!,
        subscriptionName: this.selectedSubscriptionName ?? "",
      };
    }
    return undefined;
  }

  setStatusChangeMap(
    _name: string,
    _statusChange: (
      status: string,
      token?: string,
      accountInfo?: Record<string, unknown>
    ) => Promise<void>
  ): Promise<boolean> {
    return Promise.resolve(true);
  }

  removeStatusChangeMap(_name: string): Promise<boolean> {
    return Promise.resolve(true);
  }
}

export default AzureAccountManager;
