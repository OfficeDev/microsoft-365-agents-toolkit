// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { TokenCredential } from "@azure/core-auth";
import { ok, Result } from "neverthrow";
import { FxError } from "../error";

export type AzureCredential =
  | {
      type: "AuthorizationCode";
      /**
       * The user account's e-mail address (user name).
       */
      username: string;
      /**
       * The Microsoft Entra tenant (directory) ID.
       * This parameter is used for multi-tenant account scenario
       */
      tenantId?: string;
      /**
       * whether pop up sign in flow if the account is not signed in, default is true
       */
      popUpSignIn?: boolean;
    }
  | {
      type: "ClientSecretCredential";
      /**
       * The Microsoft Entra tenant (directory) ID.
       */
      tenantId: string;
      /**
       * The client (application) ID of an App Registration in the tenant.
       */
      clientId: string;
      /**
       * A client secret that was generated for the App Registration.
       */
      clientSecret: string;
    }
  | {
      type: "ClientCertificateCredential";
      /**
       * The Microsoft Entra tenant (directory) ID.
       */
      tenantId: string;
      /**
       * The client (application) ID of an App Registration in the tenant.
       */
      clientId: string;
      /**
       * The path to a PEM-encoded public/private key certificate on the filesystem.
       */
      certificatePath: string;
    };

/**
 * Difference between getAccountCredential and getIdentityCredential [Node Azure Authenticate](https://docs.microsoft.com/en-us/azure/developer/javascript/core/node-sdk-azure-authenticate)
 * You can search at [Azure JS SDK](https://docs.microsoft.com/en-us/javascript/api/overview/azure/?view=azure-node-latest) to see which credential you need.
 */
export interface AzureAccountProvider {
  /**
   * Async get identity [crendential](https://github.com/Azure/azure-sdk-for-js/blob/master/sdk/core/core-auth/src/tokenCredential.ts)
   * @param showDialog Control whether the UI layer displays pop-up windows.
   */
  getIdentityCredentialAsync(showDialog?: boolean): Promise<TokenCredential | undefined>;

  /**
   * To support credential per action feature, caller can specify credential info for on demand
   * This method will be optional until V3 first release, after that it will be changed to required
   * @param credential
   */
  getIdentityCredential?(credential: AzureCredential): Promise<TokenCredential | undefined>;

  /**
   * Azure sign out
   */
  signout(): Promise<boolean>;
  /**
   * Switch to specified tenant for current user account
   * @param tenantId id of tenant that user wants to switch to
   */
  switchTenant(tenantId: string): Promise<Result<TokenCredential, FxError>>;
  /**
   * Add update account info callback
   * @param name callback name
   * @param statusChange callback method
   * @param immediateCall whether callback when register, the default value is true
   */
  setStatusChangeMap(
    name: string,
    statusChange: (
      status: string,
      token?: string,
      accountInfo?: Record<string, unknown>
    ) => Promise<void>,
    immediateCall?: boolean
  ): Promise<boolean>;

  /**
   * Remove update account info callback
   * @param name callback name
   */
  removeStatusChangeMap(name: string): Promise<boolean>;

  /**
   * Get Azure token JSON object
   * - tid : tenantId
   * - unique_name : user name
   * - ...
   * @param showDialog Control whether the UI layer displays pop-up windows
   */
  getJsonObject(showDialog?: boolean): Promise<Record<string, unknown> | undefined>;

  /**
   * List subscription detail
   */
  listSubscriptions(): Promise<SubscriptionInfo[]>;

  /**
   * Set subscription id to memory
   * @param subscriptionId user used subscription id (subscriptionId==="" will clear sub information in Azure account provider
   * cache)
   */
  setSubscription(subscriptionId: string): Promise<void>;

  /**
   * Get account information
   */
  getAccountInfo(): Record<string, string> | undefined;

  /**
   * Get user select subscription, tenant information
   * @param triggerUI whether means trigger login or select subscription workflow when user has not logged in or selected subscription
   * @returns SubscriptionInfo.subscriptionId === "", means user does not select subscription
   */
  getSelectedSubscription(triggerUI?: boolean): Promise<SubscriptionInfo | undefined>;
}

export type SubscriptionInfo = {
  subscriptionName: string;
  subscriptionId: string;
  tenantId: string;
};

export declare type TokenRequest = {
  scopes: Array<string>;
  showDialog?: boolean;
};
export type LoginStatus = {
  status: string;
  token?: string;
  accountInfo?: Record<string, unknown>;
};

/**
 * Provide M365 accessToken and JSON object
 *
 */
export interface M365TokenProvider {
  /**
   * Get M365 access token
   * @param tokenRequest permission scopes or show user interactive UX
   */
  getAccessToken(tokenRequest: TokenRequest): Promise<Result<string, FxError>>;
  /**
   * Get M365 token Json object
   * - tid : tenantId
   * - unique_name : user name
   * - ...
   * @param tokenRequest permission scopes or show user interactive UX
   * @param tenantId id of tenant that user wants to login in to
   */
  getJsonObject(
    tokenRequest: TokenRequest,
    tenantId?: string
  ): Promise<Result<Record<string, unknown>, FxError>>;
  /**
   * Get user login status
   * @param tokenRequest permission scopes or show user interactive UX
   */
  getStatus(tokenRequest: TokenRequest): Promise<Result<LoginStatus, FxError>>;
  /**
   * Switch to specified tenant for current user account
   * @param tenantId id of tenant that user wants to switch to
   */
  switchTenant(tenantId: string): Promise<Result<string, FxError>>;
  /**
   * Add update account info callback
   * @param name callback name
   * @param tokenRequest permission scopes
   * @param statusChange callback method
   * @param immediateCall whether callback when register, the default value is true
   */
  setStatusChangeMap(
    name: string,
    tokenRequest: TokenRequest,
    statusChange: (
      status: string,
      token?: string,
      accountInfo?: Record<string, unknown>
    ) => Promise<void>,
    immediateCall?: boolean
  ): Promise<Result<boolean, FxError>>;
  /**
   * Remove update account info callback
   * @param name callback name
   */
  removeStatusChangeMap(name: string): Promise<Result<boolean, FxError>>;
}

export type TokenProvider = {
  azureAccountProvider: AzureAccountProvider;
  m365TokenProvider: M365TokenProvider;
};

/**
 * Provide basic login framework
 */
export abstract class BasicLogin {
  statusChangeMap = new Map();

  async setStatusChangeMap(
    name: string,
    tokenRequest: TokenRequest,
    statusChange: (
      status: string,
      token?: string,
      accountInfo?: Record<string, unknown>
    ) => Promise<void>,
    immediateCall = true
  ): Promise<Result<boolean, FxError>> {
    this.statusChangeMap.set(name, statusChange);
    if (immediateCall) {
      const loginStatusRes = await this.getStatus(tokenRequest);
      if (loginStatusRes.isOk()) {
        await statusChange(
          loginStatusRes.value.status,
          loginStatusRes.value.token,
          loginStatusRes.value.accountInfo
        );
      }
    }
    return ok(true);
  }

  removeStatusChangeMap(name: string): Promise<Result<boolean, FxError>> {
    this.statusChangeMap.delete(name);
    return Promise.resolve(ok(true));
  }

  abstract getStatus(tokenRequest: TokenRequest): Promise<Result<LoginStatus, FxError>>;

  async notifyStatus(tokenRequest: TokenRequest): Promise<void> {
    const loginStatusRes = await this.getStatus(tokenRequest);
    if (loginStatusRes.isOk()) {
      for (const entry of this.statusChangeMap.entries()) {
        entry[1](
          loginStatusRes.value.status,
          loginStatusRes.value.token,
          loginStatusRes.value.accountInfo
        );
      }
    }
  }
}
