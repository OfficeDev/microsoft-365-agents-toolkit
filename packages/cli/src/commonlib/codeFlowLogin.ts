// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AccountInfo, Configuration, PublicClientApplication, TokenCache } from "@azure/msal-node";
import { FxError, LogLevel, Result, SystemError, UserError, err, ok } from "@microsoft/teamsfx-api";
import { Mutex } from "async-mutex";
import * as crypto from "crypto";
import express from "express";
import * as fs from "fs-extra";
import * as http from "http";
import { AddressInfo } from "net";
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
import {
  MFACode,
  azureLoginMessage,
  env,
  m365LoginMessage,
  sendFileTimeout,
} from "./common/constant";
import CliCodeLogInstance from "./log";
import { featureFlagManager, FeatureFlags } from "@microsoft/teamsfx-core";

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

interface Deferred<T> {
  resolve: (result: T | Promise<T>) => void;
  reject: (reason: any) => void;
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
  msalTokenCache: TokenCache;
  accountName: string;
  socketMap: Map<number, any>;

  constructor(scopes: string[], config: Configuration, port: number, accountName: string) {
    this.scopes = scopes;
    this.config = config;
    this.port = port;
    this.mutex = new Mutex();
    this.pca = new PublicClientApplication(this.config);
    this.msalTokenCache = this.pca.getTokenCache();
    this.accountName = accountName;
    this.socketMap = new Map();
  }

  async reloadCache() {
    const accountCache = await loadAccountId(this.accountName);
    if (accountCache) {
      const dataCache = await this.msalTokenCache.getAccountByHomeId(accountCache);
      if (dataCache) {
        this.account = dataCache;
      }

      const tenantCache = await loadTenantId(this.accountName);
      if (tenantCache) {
        const allAccounts = await this.msalTokenCache.getAllAccounts();
        this.account = allAccounts.find((account) => account.tenantId == tenantCache);
      }
    } else {
      this.account = undefined;
    }
  }

  async login(scopes: Array<string>, tenantId?: string): Promise<string> {
    CliTelemetry.sendTelemetryEvent(TelemetryEvent.AccountLoginStart, {
      [TelemetryProperty.AccountType]: this.accountName,
    });
    const codeVerifier = CodeFlowLogin.toBase64UrlEncoding(
      crypto.randomBytes(32).toString("base64")
    );
    const codeChallenge = CodeFlowLogin.toBase64UrlEncoding(
      await CodeFlowLogin.sha256(codeVerifier)
    );
    let serverPort = this.port;

    // try get an unused port
    const app = express();
    const server = app.listen(serverPort);
    serverPort = (server.address() as AddressInfo).port;
    let lastSocketKey = 0;
    server.on("connection", (socket) => {
      const socketKey = ++lastSocketKey;
      this.socketMap.set(socketKey, socket);
      socket.on("close", () => {
        this.socketMap.delete(socketKey);
      });
    });

    server.on("close", () => {
      this.destroySockets();
    });

    const authority = tenantId ? env.activeDirectoryEndpointUrl + tenantId : undefined;
    const authCodeUrlParameters = {
      scopes: scopes,
      codeChallenge: codeChallenge,
      codeChallengeMethod: "S256",
      redirectUri: `http://localhost:${serverPort}`,
      prompt: "select_account",
      authority: authority,
    };

    let deferredRedirect: Deferred<string>;
    const redirectPromise: Promise<string> = new Promise<string>(
      (resolve, reject) => (deferredRedirect = { resolve, reject })
    );

    app.get("/", (req: express.Request, res: express.Response) => {
      const tokenRequest = {
        code: req.query.code as string,
        scopes: scopes,
        redirectUri: `http://localhost:${serverPort}`,
        codeVerifier: codeVerifier,
      };

      this.pca
        .acquireTokenByCode(tokenRequest)
        .then(async (response) => {
          if (response) {
            if (response.account) {
              await this.mutex?.runExclusive(async () => {
                this.account = response.account!;
                await saveAccountId(this.accountName, this.account.homeAccountId);
              });
              await sendFile(
                res,
                path.join(__dirname, "./codeFlowResult/index.html"),
                "text/html; charset=utf-8",
                this.accountName
              );
              this.destroySockets();
              deferredRedirect.resolve(response.accessToken);
            }
          } else {
            throw new Error("get no response");
          }
        })
        .catch((error) => {
          CliCodeLogInstance.necessaryLog(LogLevel.Error, "[Login] " + error.message);
          deferredRedirect.reject(new UserError({ error, source: ErrorMessage.loginComponent }));

          res.status(500).send(error);
        });
    });

    const codeTimer = setTimeout(() => {
      deferredRedirect.reject(
        new UserError(
          ErrorMessage.loginComponent,
          ErrorMessage.loginTimeoutTitle,
          ErrorMessage.loginTimeoutDescription
        )
      );
    }, 5 * 60 * 1000);

    function cancelCodeTimer() {
      clearTimeout(codeTimer);
    }

    let accessToken = undefined;
    try {
      await this.startServer(server, serverPort);
      void this.pca.getAuthCodeUrl(authCodeUrlParameters).then((url: string) => {
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
        void open(url);
      });

      redirectPromise.then(cancelCodeTimer, cancelCodeTimer);
      accessToken = await redirectPromise;
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
      server.close();
    }

    return accessToken;
  }

  async logout(): Promise<boolean> {
    (this.msalTokenCache as any).storage.setCache({});
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
    scopes: Array<string>,
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
      const accessToken = await this.login(scopes, tenantId);
      return ok(accessToken);
    } else {
      let tenantedAccount: AccountInfo | undefined = undefined;
      if (tenantId) {
        const allAccounts = await this.msalTokenCache.getAllAccounts();
        tenantedAccount = allAccounts.find((account) => account.tenantId == tenantId);
        this.account = tenantedAccount ?? this.account;
      }
      try {
        const res = await this.pca.acquireTokenSilent({
          account: this.account,
          scopes: scopes,
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
          const accessToken = await this.login(scopes, tenantId);
          return ok(accessToken);
        }
        return err(LoginCodeFlowError(error));
      }
    }
  }

  async startServer(server: http.Server, port: number): Promise<string> {
    // handle port timeout
    let defferedPort: Deferred<string>;
    const portPromise: Promise<string> = new Promise<string>(
      (resolve, reject) => (defferedPort = { resolve, reject })
    );
    const portTimer = setTimeout(() => {
      defferedPort.reject(
        new SystemError(
          ErrorMessage.loginComponent,
          ErrorMessage.loginPortConflictTitle,
          ErrorMessage.loginPortConflictDescription
        )
      );
    }, 5000);

    function cancelPortTimer() {
      clearTimeout(portTimer);
    }

    server.on("listening", () => {
      defferedPort.resolve(`Code login server listening on port ${port}`);
    });
    portPromise.then(cancelPortTimer, cancelPortTimer);
    return portPromise;
  }

  destroySockets(): void {
    for (const key of this.socketMap.keys()) {
      this.socketMap.get(key).destroy();
    }
  }

  static toBase64UrlEncoding(base64string: string) {
    return base64string.replace(/=/g, "").replace(/\+/g, "-").replace(/\//g, "_");
  }

  static sha256(s: string | Uint8Array): Promise<string> {
    return new Promise((solve) => solve(crypto.createHash("sha256").update(s).digest("base64")));
  }
}

function sendFile(
  res: http.ServerResponse,
  filepath: string,
  contentType: string,
  accountName: string
): Promise<void> {
  return new Promise((resolve, reject) => {
    void (async () => {
      let body = await fs.readFile(filepath);
      let data = body.toString();
      data = data.replace(/\${accountName}/g, accountName == "azure" ? "Azure" : "M365");
      body = Buffer.from(data, UTF8);
      res.writeHead(200, {
        "Content-Length": body.length,
        "Content-Type": contentType,
      });

      const timeout = setTimeout(() => {
        CliCodeLogInstance.necessaryLog(LogLevel.Error, sendFileTimeout);
        reject();
      }, 10000);

      res.end(body, () => {
        clearTimeout(timeout);
        resolve();
      });
    })();
  });
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
