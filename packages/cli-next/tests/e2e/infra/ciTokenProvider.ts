// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * CI-mode TokenProvider that uses username/password credentials from env vars.
 * No interactive login, no browser popups.
 *
 * Used by lifecycle tests in CI pipelines where env vars are injected by
 * GitHub Actions secrets.
 */

import * as msal from "@azure/msal-node";
import * as identity from "@azure/identity";
import {
  type TokenProvider,
  type AzureAccountProvider,
  type M365TokenProvider,
  type TokenRequest,
  type LoginStatus,
  type FxError,
  type ITeamsFxTokenCredential,
  type SubscriptionInfo,
  ok,
  err,
  signedIn,
  signedOut,
  UserError,
} from "@microsoft/teamsfx-core-next";
import type { Result } from "neverthrow";
import { getConfig, CLIENT_ID } from "./config";

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function convertTokenToJson(token: string): Record<string, unknown> {
  const parts = token.split(".");
  if (parts.length < 2) return {};
  const payload = Buffer.from(parts[1], "base64").toString("utf-8");
  return JSON.parse(payload);
}

// ---------------------------------------------------------------------------
// M365 CI provider — acquireTokenByUsernamePassword
// ---------------------------------------------------------------------------

class M365LoginCI implements M365TokenProvider {
  private pca: msal.PublicClientApplication;
  private username: string;
  private password: string;
  private tenantId: string;

  constructor(username: string, password: string, tenantId: string) {
    this.username = username;
    this.password = password;
    this.tenantId = tenantId;
    this.pca = new msal.PublicClientApplication({
      auth: {
        clientId: CLIENT_ID,
        authority: `https://login.microsoftonline.com/${tenantId || "organizations"}`,
      },
    });
  }

  async getAccessToken(tokenRequest: TokenRequest): Promise<Result<string, FxError>> {
    try {
      const response = await this.pca.acquireTokenByUsernamePassword({
        scopes: tokenRequest.scopes,
        username: this.username,
        password: this.password,
      });
      if (response?.accessToken) {
        return ok(response.accessToken);
      }
      return err(
        new UserError({
          name: "M365LoginCI_NoToken",
          message: "acquireTokenByUsernamePassword returned no token",
          source: "M365LoginCI",
        })
      );
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      return err(
        new UserError({
          name: "M365LoginCI_Failed",
          message: `M365 CI login failed: ${msg}`,
          source: "M365LoginCI",
        })
      );
    }
  }

  async getJsonObject(
    tokenRequest: TokenRequest,
    _tenantId?: string
  ): Promise<Result<Record<string, unknown>, FxError>> {
    const tokenRes = await this.getAccessToken(tokenRequest);
    if (tokenRes.isOk()) {
      return ok(convertTokenToJson(tokenRes.value));
    }
    return err(tokenRes.error);
  }

  async getStatus(tokenRequest: TokenRequest): Promise<Result<LoginStatus, FxError>> {
    const tokenRes = await this.getAccessToken(tokenRequest);
    if (tokenRes.isOk()) {
      return ok({
        status: signedIn,
        token: tokenRes.value,
        accountInfo: convertTokenToJson(tokenRes.value),
      });
    }
    return ok({ status: signedOut, token: undefined, accountInfo: undefined });
  }

  async signout(): Promise<boolean> {
    return true;
  }

  async switchTenant(_tenantId: string): Promise<Result<string, FxError>> {
    return ok("");
  }

  async setStatusChangeMap(): Promise<Result<boolean, FxError>> {
    return ok(true);
  }

  async removeStatusChangeMap(): Promise<Result<boolean, FxError>> {
    return ok(true);
  }
}

// ---------------------------------------------------------------------------
// Azure CI provider — UsernamePasswordCredential
// ---------------------------------------------------------------------------

class AzureLoginCIUserPassword implements AzureAccountProvider {
  private credential: ITeamsFxTokenCredential;
  private subscriptionId: string;
  private tenantId: string;

  constructor(tenantId: string, username: string, password: string, subscriptionId: string) {
    this.tenantId = tenantId;
    this.subscriptionId = subscriptionId;
    this.credential = new identity.UsernamePasswordCredential(
      tenantId,
      CLIENT_ID,
      username,
      password
    );
  }

  async getIdentityCredentialAsync(): Promise<ITeamsFxTokenCredential | undefined> {
    return this.credential;
  }

  async signout(): Promise<boolean> {
    return true;
  }

  async switchTenant(_tenantId: string): Promise<Result<ITeamsFxTokenCredential, FxError>> {
    return ok(this.credential);
  }

  async setStatusChangeMap(): Promise<boolean> {
    return true;
  }

  async removeStatusChangeMap(): Promise<boolean> {
    return true;
  }

  async getJsonObject(): Promise<Record<string, unknown> | undefined> {
    const token = await this.credential.getToken("https://management.core.windows.net/.default");
    if (token?.token) {
      return convertTokenToJson(token.token);
    }
    return undefined;
  }

  async listSubscriptions(): Promise<SubscriptionInfo[]> {
    return [
      {
        subscriptionId: this.subscriptionId,
        subscriptionName: "E2E Test Subscription",
        tenantId: this.tenantId,
      },
    ];
  }

  async setSubscription(): Promise<void> {
    // No-op — subscription is fixed from env vars
  }

  getAccountInfo(): Record<string, string> | undefined {
    return { subscriptionId: this.subscriptionId, tenantId: this.tenantId };
  }

  async getSelectedSubscription(): Promise<SubscriptionInfo | undefined> {
    return {
      subscriptionId: this.subscriptionId,
      subscriptionName: "E2E Test Subscription",
      tenantId: this.tenantId,
    };
  }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Create a TokenProvider for CI mode using username/password from env vars.
 * Falls back to the CLI's interactive token provider if not in CI mode.
 */
export function createCITokenProvider(): TokenProvider {
  const cfg = getConfig();

  return {
    m365TokenProvider: new M365LoginCI(
      cfg.m365AccountName,
      cfg.m365AccountPassword,
      cfg.m365TenantId
    ),
    azureAccountProvider: new AzureLoginCIUserPassword(
      cfg.azureTenantId,
      cfg.azureAccountName,
      cfg.azureAccountPassword,
      cfg.azureSubscriptionId
    ),
  };
}
