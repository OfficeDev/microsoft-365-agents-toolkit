// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Configuration, LogLevel } from "@azure/msal-node";
import {
  type M365TokenProvider,
  type TokenRequest,
  type FxError,
  type Result,
  ok,
  err,
  signedIn,
  signedOut,
} from "@microsoft/teamsfx-core-next";
import { CryptoCachePlugin, loadTenantId } from "./cacheAccess";
import { CodeFlowLogin } from "./codeFlowLogin";
import { convertTokenToJson } from "./utils";
import { MSAL_CLIENT_ID, M365_CACHE_NAME } from "./constants";

type LoginStatus = { status: string; token?: string; accountInfo?: Record<string, unknown> };

const cachePlugin = new CryptoCachePlugin(M365_CACHE_NAME);

function getConfig(): Configuration {
  const config: Configuration = {
    auth: {
      clientId: MSAL_CLIENT_ID,
      authority: "https://login.microsoftonline.com/common",
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
 * M365 token provider implementation for CLI — singleton.
 *
 * Uses MSAL interactive login (browser or broker on Windows)
 * with encrypted token cache.
 */
class M365Login implements M365TokenProvider {
  private static instance: M365Login;
  private codeFlow: CodeFlowLogin;

  private constructor() {
    this.codeFlow = new CodeFlowLogin([], getConfig(), 0, M365_CACHE_NAME);
  }

  static getInstance(): M365Login {
    if (!M365Login.instance) {
      M365Login.instance = new M365Login();
    }
    return M365Login.instance;
  }

  async getAccessToken(
    tokenRequest: TokenRequest,
    tenantId?: string
  ): Promise<Result<string, FxError>> {
    if (!this.codeFlow.account) {
      await this.codeFlow.reloadCache();
    }
    return this.codeFlow.getTokenByScopes(tokenRequest.scopes, true, tenantId);
  }

  async getJsonObject(
    tokenRequest: TokenRequest,
    tenantId?: string
  ): Promise<Result<Record<string, unknown>, FxError>> {
    const tokenRes = await this.getAccessToken(tokenRequest, tenantId);
    if (tokenRes.isOk()) {
      return ok(convertTokenToJson(tokenRes.value));
    }
    return err(tokenRes.error);
  }

  async getStatus(tokenRequest: TokenRequest): Promise<Result<LoginStatus, FxError>> {
    if (!this.codeFlow.account) {
      await this.codeFlow.reloadCache();
    }
    if (this.codeFlow.account) {
      const tokenRes = await this.codeFlow.getTokenByScopes(tokenRequest.scopes, false);
      if (tokenRes.isOk()) {
        const tokenJson = convertTokenToJson(tokenRes.value);
        return ok({ status: signedIn, token: tokenRes.value, accountInfo: tokenJson });
      }
      // Offline but has cached account
      return ok({
        status: signedIn,
        token: undefined,
        accountInfo: { upn: this.codeFlow.account.username },
      });
    }
    return ok({ status: signedOut, token: undefined, accountInfo: undefined });
  }

  async signout(): Promise<boolean> {
    this.codeFlow.account = undefined;
    await this.codeFlow.logout();
    return true;
  }

  async switchTenant(tenantId: string): Promise<Result<string, FxError>> {
    await this.codeFlow.switchTenant(tenantId);
    return ok("");
  }

  async getTenant(): Promise<string | undefined> {
    return (await loadTenantId(M365_CACHE_NAME)) ?? undefined;
  }

  setStatusChangeMap(
    _name: string,
    _tokenRequest: TokenRequest,
    _statusChange: (
      status: string,
      token?: string,
      accountInfo?: Record<string, unknown>
    ) => Promise<void>,
    _immediateCall?: boolean
  ): Promise<Result<boolean, FxError>> {
    return Promise.resolve(ok(true));
  }

  removeStatusChangeMap(_name: string): Promise<Result<boolean, FxError>> {
    return Promise.resolve(ok(true));
  }
}

export default M365Login;
