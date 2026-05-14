// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  AccountInfo,
  Configuration,
  PublicClientApplication,
  type SilentFlowRequest,
} from "@azure/msal-node";
import {
  type AuthenticationWWWAuthenticateRequest,
  type FxError,
  type Result,
  UserError,
  ok,
  err,
} from "@microsoft/teamsfx-core-next";
import { Mutex } from "async-mutex";
import {
  clearCache,
  loadAccountId,
  loadTenantId,
  saveAccountId,
  saveTenantId,
} from "./cacheAccess";
import { AAD_ENDPOINT, azureLoginMessage, m365LoginMessage } from "./constants";
import { getAccountByHomeId, decodeClaimsChallenge, checkIsOnline } from "./utils";

const LOGIN_COMPONENT = "login";

/**
 * MSAL-based interactive login engine shared by both M365 and Azure flows.
 *
 * Supports browser-based interactive login via `acquireTokenInteractive()`
 * and silent token refresh via `acquireTokenSilent()`.
 */
export class CodeFlowLogin {
  pca: PublicClientApplication;
  account: AccountInfo | undefined;
  config: Configuration;
  port: number;
  mutex: Mutex;
  accountName: string;

  constructor(scopes: string[], config: Configuration, port: number, accountName: string) {
    this.config = config;
    this.port = port;
    this.mutex = new Mutex();
    this.pca = new PublicClientApplication(this.config);
    this.accountName = accountName;
  }

  async reloadCache(): Promise<void> {
    const accountCache = await loadAccountId(this.accountName);
    if (accountCache) {
      const allAccounts = await this.pca.getAllAccounts();
      const dataCache = getAccountByHomeId(accountCache, allAccounts);
      if (dataCache) {
        this.account = dataCache;
      }
      const tenantCache = await loadTenantId(this.accountName);
      if (tenantCache) {
        this.account = allAccounts.find((a) => a.tenantId === tenantCache);
      }
    } else {
      this.account = undefined;
    }
  }

  async login(
    requestScopes: string[] | AuthenticationWWWAuthenticateRequest,
    tenantId?: string
  ): Promise<string> {
    let scopes: string[];
    let claim: string | undefined;
    if (typeof requestScopes === "object" && "wwwAuthenticate" in requestScopes) {
      scopes = requestScopes.scopes ?? [];
      claim = decodeClaimsChallenge(
        (requestScopes as AuthenticationWWWAuthenticateRequest).wwwAuthenticate ?? ""
      );
    } else {
      scopes = requestScopes;
    }

    const authority = tenantId ? AAD_ENDPOINT + tenantId : undefined;

    const interactiveRequest = {
      scopes,
      authority,
      prompt: "select_account" as const,
      claims: claim,
      openBrowser: async (url: string) => {
        url += "#";
        const msg = this.accountName === "azure" ? azureLoginMessage : m365LoginMessage;
        console.log(msg + url);
        const open = (await import("open")).default;
        await open(url);
      },
    };

    const response = await this.pca.acquireTokenInteractive(interactiveRequest);
    if (response && response.account) {
      await this.mutex.runExclusive(async () => {
        this.account = response.account!;
        await saveAccountId(this.accountName, this.account.homeAccountId);
      });
      return response.accessToken;
    }
    throw new UserError({
      name: "LoginFail",
      message: "Cannot retrieve user login information. Login with another account.",
      source: LOGIN_COMPONENT,
    });
  }

  async logout(): Promise<boolean> {
    const accounts = await this.pca.getAllAccounts();
    for (const account of accounts) {
      await this.pca.signOut({ account });
    }
    await clearCache(this.accountName);
    await saveAccountId(this.accountName, undefined);
    await saveTenantId(this.accountName, undefined);
    this.account = undefined;
    return true;
  }

  async switchTenant(tenantId: string): Promise<void> {
    await saveTenantId(this.accountName, tenantId);
  }

  /**
   * Core token acquisition: tries silent → falls back to interactive.
   */
  async getTokenByScopes(
    scopes: string | string[] | AuthenticationWWWAuthenticateRequest,
    refresh = true,
    tenantId?: string
  ): Promise<Result<string, FxError>> {
    if (!this.account) {
      await this.reloadCache();
    }

    if (!tenantId) {
      tenantId = (await loadTenantId(this.accountName)) ?? undefined;
    }

    // No cached account → do full interactive login
    if (!this.account) {
      const accessToken = await this.login(
        typeof scopes === "string" ? [scopes] : scopes,
        tenantId
      );
      return ok(accessToken);
    }

    // Have cached account → try silent
    let myScopes: string[];
    if (typeof scopes === "string") {
      myScopes = [scopes];
    } else if (typeof scopes === "object" && "wwwAuthenticate" in scopes) {
      myScopes = (scopes as AuthenticationWWWAuthenticateRequest).scopes ?? [];
    } else {
      myScopes = scopes;
    }

    let tenantedAccount: AccountInfo | undefined;
    if (tenantId) {
      const allAccounts = await this.pca.getAllAccounts();
      tenantedAccount = allAccounts.find((a) => a.tenantId === tenantId);
      this.account = tenantedAccount ?? this.account;
    }

    const tokenRequest: SilentFlowRequest = {
      account: this.account,
      scopes: myScopes,
      authority: tenantId ? AAD_ENDPOINT + tenantId : this.config.auth.authority,
      forceRefresh: !tenantedAccount,
    };

    try {
      const res = await this.pca.acquireTokenSilent(tokenRequest);
      if (res) {
        return ok(res.accessToken);
      }
      return err(
        new UserError({
          name: "LoginCodeFail",
          message: "No token response from silent acquisition.",
          source: LOGIN_COMPONENT,
        })
      );
    } catch (error: unknown) {
      if (!(await checkIsOnline())) {
        return err(
          new UserError({
            name: "CheckOnlineFail",
            message: "You appear to be offline. Please check your network connection.",
            source: LOGIN_COMPONENT,
          })
        );
      }
      if (refresh) {
        await this.logout();
        const accessToken = await this.login(myScopes, tenantId);
        return ok(accessToken);
      }
      return err(
        new UserError({
          name: "LoginCodeFail",
          message: "Cannot get login code for token exchange. Login with another account.",
          source: LOGIN_COMPONENT,
          error: error instanceof Error ? error : new Error(String(error)),
        })
      );
    }
  }
}
