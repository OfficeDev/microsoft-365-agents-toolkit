// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/* eslint-disable @typescript-eslint/ban-ts-comment */
/* eslint-disable @typescript-eslint/ban-types */
"use strict";

import {
  FxError,
  M365TokenProvider,
  ok,
  Result,
  TokenRequest,
  err,
  LoginStatus,
  BasicLogin,
  UserError,
  SystemError,
} from "@microsoft/teamsfx-api";
import { AccountInfo, LogLevel } from "@azure/msal-node";
import { ExtensionErrors, ExtensionSource } from "../error/error";
import { CodeFlowLogin, ConvertTokenToJson, UserCancelError } from "./codeFlowLogin";
import VsCodeLogInstance from "./log";
import * as vscode from "vscode";
import { CryptoCachePlugin, saveAccountId } from "./cacheAccess";
import {
  loggedIn,
  loggingIn,
  m365CacheName,
  signedIn,
  signedOut,
  signingIn,
  switching,
} from "./common/constant";
import * as util from "util";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  AccountType,
  TelemetryErrorType,
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";
import { getDefaultString, localize } from "../utils/localizeUtils";
import { AppStudioScopes } from "@microsoft/teamsfx-core";

const SERVER_PORT = 0;
const cachePlugin = new CryptoCachePlugin(m365CacheName);

const config = {
  auth: {
    clientId: "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0",
    authority: "https://login.microsoftonline.com/common",
  },
  system: {
    loggerOptions: {
      // @ts-ignore
      loggerCallback(loglevel, message, containsPii) {
        if (loglevel <= LogLevel.Error) {
          VsCodeLogInstance.error(message);
        }
      },
      piiLoggingEnabled: false,
      logLevel: LogLevel.Error,
    },
  },
  cache: {
    cachePlugin,
  },
};

export class M365Login extends BasicLogin implements M365TokenProvider {
  private static instance: M365Login;
  private static codeFlowInstance: CodeFlowLogin;

  private constructor() {
    super();
    M365Login.codeFlowInstance = new CodeFlowLogin([], config, SERVER_PORT, m365CacheName);
  }

  /**
   * Gets instance
   * @returns instance
   */
  public static getInstance(): M365Login {
    if (!M365Login.instance) {
      M365Login.instance = new M365Login();
    }

    return M365Login.instance;
  }

  /**
   * Get team access token
   */
  async getAccessToken(
    tokenRequest: TokenRequest,
    loginHint?: string
  ): Promise<Result<string, FxError>> {
    await M365Login.codeFlowInstance.reloadCache();
    let tokenRes: Result<string, FxError>;
    if (!M365Login.codeFlowInstance.account) {
      tokenRes = await this.signInWhenNoAccountInCache(tokenRequest, loginHint);
    } else {
      tokenRes = await M365Login.codeFlowInstance.getTokenByScopes(
        tokenRequest.scopes,
        true,
        loginHint
      );
    }

    if (tokenRes.isOk()) {
      return ok(tokenRes.value);
    } else {
      return tokenRes;
    }
  }

  private async signInWhenNoAccountInCache(
    tokenRequest: TokenRequest,
    loginHint: string | undefined,
    isInitiatedFromTdp?: boolean
  ): Promise<Result<string, FxError>> {
    if (tokenRequest.showDialog === undefined || tokenRequest.showDialog) {
      let userConfirmation = false;
      if (!isInitiatedFromTdp) {
        userConfirmation = await this.doesUserConfirmLogin();
      } else {
        userConfirmation = await this.doesUserConfirmLoginWhenIntiatedFromTdp();
      }

      if (!userConfirmation) {
        const cancelError = !isInitiatedFromTdp
          ? UserCancelError(getDefaultString("teamstoolkit.codeFlowLogin.loginComponent"))
          : TdpIntegrationLoginUserCancelError(
              getDefaultString("teamstoolkit.codeFlowLogin.loginComponent")
            );
        ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.Login, cancelError, {
          [TelemetryProperty.AccountType]: AccountType.M365,
          [TelemetryProperty.Success]: TelemetrySuccess.No,
          [TelemetryProperty.UserId]: "",
          [TelemetryProperty.Internal]: "",
          [TelemetryProperty.ErrorType]: TelemetryErrorType.UserError,
          [TelemetryProperty.ErrorCode]: `${getDefaultString(
            "teamstoolkit.codeFlowLogin.loginComponent"
          )}.${ExtensionErrors.UserCancel}`,
          [TelemetryProperty.ErrorMessage]: `${getDefaultString("teamstoolkit.common.userCancel")}`,
        });
        return err(cancelError);
      }
      M365Login.codeFlowInstance.status = loggingIn;
      void this.notifyStatus(tokenRequest);
    }

    if (loginHint) {
      const allAccounts = await M365Login.codeFlowInstance.msalTokenCache.getAllAccounts();
      const accountMatchedInCache = !allAccounts
        ? undefined
        : allAccounts.find((o) => o.username === loginHint);
      if (!!accountMatchedInCache) {
        // If there is an account in msal cache with the same login hint, we will use that account to sign in.
        M365Login.codeFlowInstance.account = accountMatchedInCache;
        await saveAccountId(
          M365Login.codeFlowInstance.accountName,
          accountMatchedInCache.homeAccountId
        );
      }
    }
    const tokenRes = await M365Login.codeFlowInstance.getTokenByScopes(
      tokenRequest.scopes,
      true,
      loginHint
    );
    await this.notifyStatus(tokenRequest);
    return tokenRes;
  }

  async getJsonObject(
    tokenRequest: TokenRequest
  ): Promise<Result<Record<string, unknown>, FxError>> {
    const tokenRes = await this.getAccessToken(tokenRequest);
    if (tokenRes.isOk()) {
      const tokenJson = ConvertTokenToJson(tokenRes.value);
      return ok(tokenJson as any);
    } else {
      return err(tokenRes.error);
    }
  }

  async signout(): Promise<boolean> {
    await M365Login.codeFlowInstance.reloadCache();
    const userConfirmation = await this.doesUserConfirmSignout();
    if (!userConfirmation) {
      ExtTelemetry.sendTelemetryErrorEvent(
        TelemetryEvent.SignOut,
        UserCancelError(getDefaultString("teamstoolkit.commands.signOut.title")),
        {
          [TelemetryProperty.AccountType]: AccountType.M365,
          [TelemetryProperty.Success]: TelemetrySuccess.No,
          [TelemetryProperty.UserId]: "",
          [TelemetryProperty.Internal]: "",
          [TelemetryProperty.ErrorType]: TelemetryErrorType.UserError,
          [TelemetryProperty.ErrorCode]: `${getDefaultString(
            "teamstoolkit.codeFlowLogin.loginComponent"
          )}.${ExtensionErrors.UserCancel}`,
          [TelemetryProperty.ErrorMessage]: `${getDefaultString("teamstoolkit.common.userCancel")}`,
        }
      );
      throw UserCancelError(getDefaultString("teamstoolkit.commands.signOut.title"));
    }
    await M365Login.codeFlowInstance.logout();
    await this.notifyStatus({ scopes: AppStudioScopes });
    return true;
  }

  async switchTenant(tenantId: string): Promise<Result<string, FxError>> {
    try {
      const res = await M365Login.codeFlowInstance.getTokenByScopes(
        AppStudioScopes,
        true,
        undefined,
        tenantId
      );
      if (res.isOk()) {
        await M365Login.codeFlowInstance.switchTenant(tenantId);
        await this.notifyStatus({ scopes: AppStudioScopes });
      }
      return res;
    } catch (e) {
      return err(
        new SystemError({
          error: e,
          source: ExtensionSource,
          name: ExtensionErrors.LoginCacheError,
        })
      );
    }
  }

  private async doesUserConfirmLogin(): Promise<boolean> {
    const message = localize("teamstoolkit.appStudioLogin.message");
    const signin = localize("teamstoolkit.common.signin");
    const createTestingTenant = localize("teamstoolkit.appStudioLogin.createM365TestingTenant");
    let userSelected: string | undefined;
    do {
      userSelected = await vscode.window.showInformationMessage(
        message,
        { modal: true },
        signin,
        createTestingTenant
      );
      if (userSelected === createTestingTenant) {
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.OpenSignInJoinM365);
        void vscode.env.openExternal(
          vscode.Uri.parse("https://developer.microsoft.com/en-us/microsoft-365/dev-program")
        );
      }
    } while (userSelected === createTestingTenant);
    return Promise.resolve(userSelected === signin);
  }

  private async doesUserConfirmLoginWhenIntiatedFromTdp(): Promise<boolean> {
    const message = localize("teamstoolkit.devPortalIntegration.appStudioLogin.message");
    const signin = localize("teamstoolkit.common.signin");

    const userSelected = await vscode.window.showInformationMessage(
      message,
      { modal: true },
      signin
    );

    return Promise.resolve(userSelected === signin);
  }

  private async doesUserConfirmSignout(): Promise<boolean> {
    const accountInfo = M365Login.codeFlowInstance.account;
    const email = accountInfo?.username;
    const confirm = localize("teamstoolkit.common.signout");
    const userSelected = await vscode.window.showInformationMessage(
      util.format(localize("teamstoolkit.common.signOutOf"), email),
      { modal: true },
      confirm
    );
    return Promise.resolve(userSelected === confirm);
  }

  async getStatus(tokenRequest: TokenRequest): Promise<Result<LoginStatus, FxError>> {
    if (M365Login.codeFlowInstance.status !== switching) {
      await M365Login.codeFlowInstance.reloadCache();
    }
    if (M365Login.codeFlowInstance.status === loggedIn) {
      const tokenRes = await M365Login.codeFlowInstance.getTokenByScopes(
        tokenRequest.scopes,
        false
      );
      if (tokenRes.isOk()) {
        const tokenJson = ConvertTokenToJson(tokenRes.value);
        // if token is empty, try to get token by app studio scope, normally this should only affect UX
        if (Object.keys(tokenJson).length === 0) {
          const appStudioRes = await M365Login.codeFlowInstance.getTokenByScopes(
            AppStudioScopes,
            false
          );
          if (appStudioRes.isOk()) {
            const appStudioJson = ConvertTokenToJson(appStudioRes.value);
            return ok({
              status: signedIn,
              token: appStudioRes.value,
              accountInfo: appStudioJson as any,
            });
          }
        }
        return ok({ status: signedIn, token: tokenRes.value, accountInfo: tokenJson as any });
      } else {
        if (
          tokenRes.error.name !==
          getDefaultString("teamstoolkit.codeFlowLogin.checkOnlineFailTitle")
        ) {
          return ok({ status: signedOut, token: undefined, accountInfo: undefined });
        } else {
          return ok({
            status: signedIn,
            token: undefined,
            accountInfo: { upn: M365Login.codeFlowInstance.account?.username },
          });
        }
      }
    } else if (M365Login.codeFlowInstance.status === loggingIn) {
      return ok({ status: signingIn, token: undefined, accountInfo: undefined });
    } else if (M365Login.codeFlowInstance.status === switching) {
      return ok({ status: switching, token: undefined, accountInfo: undefined });
    } else {
      return ok({ status: signedOut, token: undefined, accountInfo: undefined });
    }
  }

  getCachedAccountInfo(): AccountInfo | undefined {
    return M365Login.codeFlowInstance.account;
  }

  async signInWhenInitiatedFromTdp(
    tokenRequest: TokenRequest,
    loginHint: string
  ): Promise<Result<string, FxError>> {
    await M365Login.codeFlowInstance.reloadCache();
    let tokenRes;
    if (!M365Login.codeFlowInstance.account) {
      // If no account in cache file, we will ask to sign in directly.
      tokenRes = await this.signInWhenNoAccountInCache(tokenRequest, loginHint, true);
      VsCodeLogInstance.info(`Signed in with Microsoft 365 account: ${loginHint}`);
    } else if (
      !!loginHint &&
      M365Login.codeFlowInstance.account.username.toLowerCase() === loginHint.toLowerCase()
    ) {
      // If the account in cache matched with the loginHint, we will try to get access token for the currently cached account
      tokenRes = await this.getAccessToken(tokenRequest, loginHint);
      VsCodeLogInstance.info(`Already signed in with correct Microsoft 365 account: ${loginHint}`);
    } else {
      // need to switch account
      const userConfirmation = await this.doesUserConfirmSwitchAccount();
      if (!userConfirmation) {
        const cancelError = new UserError({
          name: ExtensionErrors.UserCancel,
          message: getDefaultString(
            "teamstoolkit.devPortalIntegration.switchAccountCancel.message"
          ),
          displayMessage: localize("teamstoolkit.devPortalIntegration.switchAccountCancel.message"),
          source: "switchAccount",
        });

        ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.Login, cancelError, {
          [TelemetryProperty.AccountType]: AccountType.M365,
          [TelemetryProperty.Success]: TelemetrySuccess.No,
          [TelemetryProperty.UserId]: "",
          [TelemetryProperty.Internal]: "",
          [TelemetryProperty.ErrorType]: TelemetryErrorType.UserError,
          [TelemetryProperty.ErrorCode]: `${getDefaultString(
            "teamstoolkit.codeFlowLogin.loginComponent"
          )}.${ExtensionErrors.UserCancel}`,
          [TelemetryProperty.ErrorMessage]: `${getDefaultString("teamstoolkit.common.userCancel")}`,
        });

        return err(cancelError);
      }
      M365Login.codeFlowInstance.status = switching;
      await this.notifyStatus(tokenRequest);
      tokenRes = await M365Login.codeFlowInstance.switchAccount(tokenRequest.scopes, loginHint);
      await this.notifyStatus(tokenRequest);
      VsCodeLogInstance.info(`Switched to another Microsoft 365 account: ${loginHint}`);
    }

    if (tokenRes.isErr()) {
      return err(tokenRes.error);
    }

    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Login, {
      [TelemetryProperty.AccountType]: AccountType.M365,
      [TelemetryProperty.Success]: TelemetrySuccess.Yes,
      [TelemetryProperty.UserId]: "",
      [TelemetryProperty.Internal]: "",
      [TelemetryProperty.ErrorType]: TelemetryErrorType.UserError,
      [TelemetryProperty.ErrorCode]: `${getDefaultString(
        "teamstoolkit.codeFlowLogin.loginComponent"
      )}.${ExtensionErrors.UserCancel}`,
      [TelemetryProperty.ErrorMessage]: `${getDefaultString("teamstoolkit.common.userCancel")}`,
    });

    return ok(tokenRes.value);
  }

  private async doesUserConfirmSwitchAccount(): Promise<boolean> {
    const message = localize("teamstoolkit.devPortalIntegration.appStudioSwitchAccount.message");
    const switchAccount = localize("teamstoolkit.devPortalIntegration.switchAccount");

    const userSelected = await vscode.window.showInformationMessage(
      message,
      { modal: true },
      switchAccount
    );

    return Promise.resolve(userSelected === switchAccount);
  }
}

export function TdpIntegrationLoginUserCancelError(source: string): UserError {
  return new UserError({
    name: ExtensionErrors.UserCancel,
    message: getDefaultString("teamstoolkit.devPortalIntegration.signInCancel.message"),
    displayMessage: localize("teamstoolkit.devPortalIntegration.signInCancel.message"),
    source: source,
  });
}

export default M365Login.getInstance();
