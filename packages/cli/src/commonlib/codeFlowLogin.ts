// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AccountInfo, Configuration, PublicClientApplication, TokenCache } from "@azure/msal-node";
import {
  AuthenticationWWWAuthenticateRequest,
  FxError,
  LogLevel,
  Result,
  UserError,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import { Mutex } from "async-mutex";
import * as crypto from "crypto";
import * as http from "http";
import open from "open";
import os from "os";
import * as path from "path";
import { TextType, colorize } from "../colorize";
import CliTelemetry from "../telemetry/cliTelemetry";
import {
  TelemetryErrorType,
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/cliTelemetryEvents";
import {
  UTF8,
  clearCache,
  loadAccountId,
  loadTenantId,
  saveAccountId,
  saveTenantId,
} from "./cacheAccess";
import { azureLoginMessage, env, m365LoginMessage } from "./common/constant";
import CliCodeLogInstance from "./log";
import { decodeClaimsChallenge } from "./common/utils";
import { getAccountByHomeId } from "./common/tokenCacheUtils";

export class ErrorMessage {
  static readonly loginFailureTitle = "LoginFail";
  static readonly loginFailureDescription =
    "Cannot retrieve user login information. Login with another account.";
  static readonly loginCodeFlowFailureTitle = "LoginCodeFail";
  static readonly loginCodeFlowFailureDescription =
    "Cannot get login code for token exchange. Login with another account.";
  static readonly loginTimeoutTitle = "LoginTimeout";
  static readonly loginTimeoutDescription = "Timeout waiting for login. Try again.";
  static readonly loginPortConflictTitle = "LoginPortConflict";
  static readonly loginPortConflictDescription = "Timeout waiting for port. Try again.";
  static readonly loginComponent = "login";
  static readonly checkOnlineFailTitle = "CheckOnlineFail";
  static readonly checkOnlineFailDetail =
    "You appear to be offline. Please check your network connection.";
  static readonly loginUsernamePasswordFailTitle = "UsernamePasswordLoginFail";
  static readonly loginUsernamePasswordFailDetail =
    "Fail to login via username and password. Please check your username or password";
}

export class CodeFlowLogin {
  pca: PublicClientApplication;
  account: AccountInfo | undefined;
  /**
   * @deprecated will be removed after unify m365 login
   */
  scopes: string[];
  config: Configuration;
  port: number;
  mutex: Mutex;
  accountName: string;

  constructor(scopes: string[], config: Configuration, port: number, accountName: string) {
    this.scopes = scopes;
    this.config = config;
    this.port = port;
    this.mutex = new Mutex();
    this.pca = new PublicClientApplication(this.config);
    this.accountName = accountName;
  }

  async reloadCache() {
    const accountCache = await loadAccountId(this.accountName);
    if (accountCache) {
      const dataCache = getAccountByHomeId(accountCache, await this.pca.getAllAccounts());
      if (dataCache) {
        this.account = dataCache;
      }

      const tenantCache = await loadTenantId(this.accountName);
      if (tenantCache) {
        const allAccounts = await this.pca.getAllAccounts();
        this.account = allAccounts.find((account) => account.tenantId == tenantCache);
      }
    } else {
      this.account = undefined;
    }
  }

  async login(
    requestScopes: Array<string> | AuthenticationWWWAuthenticateRequest,
    tenantId?: string
  ): Promise<string> {
    CliTelemetry.sendTelemetryEvent(TelemetryEvent.AccountLoginStart, {
      [TelemetryProperty.AccountType]: this.accountName,
    });
    let scopes: string[];
    let claim = undefined;
    if (typeof requestScopes === "object" && "wwwAuthenticate" in requestScopes) {
      scopes = requestScopes.scopes ?? [];
      claim = decodeClaimsChallenge(requestScopes.wwwAuthenticate);
    } else {
      scopes = requestScopes;
    }

    const authority = tenantId ? env.activeDirectoryEndpointUrl + tenantId : undefined;
    const interactiveRequest = {
      scopes: scopes,
      authority: authority,
      prompt: "select_account",
      claims: claim,
      openBrowser: async (url: string) => {
        url += "#";
        if (this.accountName == "azure") {
          CliCodeLogInstance.outputInfo(
            azureLoginMessage + colorize(url, TextType.Hyperlink) + os.EOL
          );
        } else {
          CliCodeLogInstance.outputInfo(
            m365LoginMessage + colorize(url, TextType.Hyperlink) + os.EOL
          );
        }
        await open(url);
      },
    };

    let accessToken = undefined;
    try {
      const response = await this.pca.acquireTokenInteractive(interactiveRequest);

      if (response && response.account) {
        await this.mutex?.runExclusive(async () => {
          this.account = response.account!;
          await saveAccountId(this.accountName, this.account.homeAccountId);
        });
        accessToken = response.accessToken;
      } else {
        throw new Error("No response or account from interactive login");
      }
    } catch (e: any) {
      CliTelemetry.sendTelemetryEvent(TelemetryEvent.AccountLogin, {
        [TelemetryProperty.AccountType]: this.accountName,
        [TelemetryProperty.Success]: TelemetrySuccess.No,
        [TelemetryProperty.UserId]: "",
        [TelemetryProperty.Internal]: "",
        [TelemetryProperty.ErrorType]:
          e instanceof UserError ? TelemetryErrorType.UserError : TelemetryErrorType.SystemError,
        [TelemetryProperty.ErrorCode]: `${e.source}.${e.name}`,
        [TelemetryProperty.ErrorMessage]: `${e.message}`,
      });
      throw e;
    } finally {
      if (accessToken) {
        const tokenJson = ConvertTokenToJson(accessToken);
        CliTelemetry.sendTelemetryEvent(TelemetryEvent.AccountLogin, {
          [TelemetryProperty.AccountType]: this.accountName,
          [TelemetryProperty.Success]: TelemetrySuccess.Yes,
          [TelemetryProperty.UserId]: (tokenJson as any).oid ? (tokenJson as any).oid : "",
          [TelemetryProperty.Internal]: (tokenJson as any).upn?.endsWith("@microsoft.com")
            ? "true"
            : "false",
        });
      }
    }

    return accessToken;
  }

  async logout(): Promise<boolean> {
    const accounts = await this.pca.getAllAccounts();
    for (const account of accounts) {
      await this.pca.signOut({ account: account });
    }
    await clearCache(this.accountName);
    await saveAccountId(this.accountName, undefined);
    await saveTenantId(this.accountName, undefined);
    this.account = undefined;
    return true;
  }

  async switchTenant(tenantId: string): Promise<void> {
    return await saveTenantId(this.accountName, tenantId);
  }

  async getTokenByScopes(
    scopes: string | string[] | AuthenticationWWWAuthenticateRequest,
    refresh = true,
    tenantId?: string
  ): Promise<Result<string, FxError>> {
    if (!this.account) {
      await this.reloadCache();
    }

    if (!tenantId) {
      tenantId = await loadTenantId(this.accountName);
    }

    if (!this.account) {
      const accessToken = await this.login(
        typeof scopes === "string" ? [scopes] : scopes,
        tenantId
      );
      return ok(accessToken);
    } else {
      let myScopes: string[] = [];
      if (typeof scopes === "string") {
        myScopes = [scopes];
      } else if (typeof scopes === "object" && "wwwAuthenticate" in scopes) {
        myScopes = (scopes as AuthenticationWWWAuthenticateRequest).scopes ?? [];
      } else {
        myScopes = scopes;
      }

      let tenantedAccount: AccountInfo | undefined = undefined;
      if (tenantId) {
        const allAccounts = await this.pca.getAllAccounts();
        tenantedAccount = allAccounts.find((account) => account.tenantId == tenantId);
        this.account = tenantedAccount ?? this.account;
      }
      try {
        const res = await this.pca.acquireTokenSilent({
          account: this.account,
          scopes: myScopes,
          forceRefresh: tenantedAccount ? false : true,
          authority: tenantId
            ? env.activeDirectoryEndpointUrl + tenantId
            : this.config.auth.authority,
        });
        if (res) {
          return ok(res.accessToken);
        } else {
          return err(LoginCodeFlowError(new Error("No token response")));
        }
      } catch (error: any) {
        CliCodeLogInstance.necessaryLog(
          LogLevel.Debug,
          "[Login] Failed to retrieve token silently. If you encounter this problem multiple times, you can delete `" +
            path.join(os.homedir(), ".fx", "account") +
            "` and try again. " +
            error.message
        );
        if (!(await checkIsOnline())) {
          return err(CheckOnlineError());
        }
        await this.logout();
        if (refresh) {
          const accessToken = await this.login(myScopes, tenantId);
          return ok(accessToken);
        }
        return err(LoginCodeFlowError(error));
      }
    }
  }

  static toBase64UrlEncoding(base64string: string) {
    return base64string.replace(/=/g, "").replace(/\+/g, "-").replace(/\//g, "_");
  }

  static sha256(s: string | Uint8Array): Promise<string> {
    return new Promise((solve) => solve(crypto.createHash("sha256").update(s).digest("base64")));
  }
}

export function LoginFailureError(innerError?: any): UserError {
  return new UserError({
    name: ErrorMessage.loginCodeFlowFailureTitle,
    message: ErrorMessage.loginCodeFlowFailureDescription,
    source: ErrorMessage.loginComponent,
    error: innerError,
  });
}

export function LoginCodeFlowError(innerError?: any): UserError {
  return new UserError({
    name: ErrorMessage.loginCodeFlowFailureTitle,
    message: ErrorMessage.loginCodeFlowFailureDescription,
    source: ErrorMessage.loginComponent,
    error: innerError,
  });
}

export function CheckOnlineError(): UserError {
  return new UserError({
    name: ErrorMessage.checkOnlineFailTitle,
    message: ErrorMessage.checkOnlineFailDetail,
    source: ErrorMessage.loginComponent,
  });
}

export function ConvertTokenToJson(token: string): any {
  const array = token.split(".");
  if (array.length === 5) {
    // this is a JWE
    return {};
  }
  const buff = Buffer.from(array[1], "base64");
  return JSON.parse(buff.toString(UTF8));
}

export async function checkIsOnline(): Promise<boolean> {
  const options = {
    hostname: "login.microsoftonline.com",
    path: "/",
    method: "head",
  };

  return new Promise((resolve, reject) => {
    const req = http.request(options, (res) => {
      res.on("data", () => {});
      res.on("end", () => {
        resolve(true);
      });
    });
    req.on("error", (e) => resolve(false));
    req.end();
  });
}
