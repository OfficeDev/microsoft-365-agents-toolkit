// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AccessToken } from "@azure/core-auth";
import { Result } from "neverthrow";
import { FxError } from "../error";

export const signedIn = "SignedIn";
export const signedOut = "SignedOut";

export type AzureCredential =
  | {
      type: "AuthorizationCode";
      username: string;
      tenantId?: string;
      popUpSignIn?: boolean;
    }
  | {
      type: "ClientSecretCredential";
      tenantId: string;
      clientId: string;
      clientSecret: string;
    }
  | {
      type: "ClientCertificateCredential";
      tenantId: string;
      clientId: string;
      certificatePath: string;
    };

export interface ITeamsFxTokenCredential {
  getToken(
    scopes: string | string[] | AuthenticationWWWAuthenticateRequest,
    options?: any
  ): Promise<AccessToken | null>;
}

export interface AuthenticationWWWAuthenticateRequest {
  readonly wwwAuthenticate: string;
  readonly scopes?: string[];
}

export interface AzureAccountProvider {
  getIdentityCredentialAsync(
    showDialog?: boolean,
    authenticationSessionRequest?: AuthenticationWWWAuthenticateRequest
  ): Promise<ITeamsFxTokenCredential | undefined>;

  getIdentityCredential?(credential: AzureCredential): Promise<ITeamsFxTokenCredential | undefined>;

  signout(): Promise<boolean>;

  switchTenant(tenantId: string): Promise<Result<ITeamsFxTokenCredential, FxError>>;

  setStatusChangeMap(
    name: string,
    statusChange: (
      status: string,
      token?: string,
      accountInfo?: Record<string, unknown>
    ) => Promise<void>,
    immediateCall?: boolean
  ): Promise<boolean>;

  removeStatusChangeMap(name: string): Promise<boolean>;

  getJsonObject(showDialog?: boolean): Promise<Record<string, unknown> | undefined>;

  listSubscriptions(): Promise<SubscriptionInfo[]>;

  setSubscription(subscriptionId: string): Promise<void>;

  getAccountInfo(): Record<string, string> | undefined;

  getSelectedSubscription(triggerUI?: boolean): Promise<SubscriptionInfo | undefined>;
}

export type SubscriptionInfo = {
  subscriptionName: string;
  subscriptionId: string;
  tenantId: string;
};

export type TokenRequest = {
  scopes: Array<string>;
  showDialog?: boolean;
};

export type LoginStatus = {
  status: string;
  token?: string;
  accountInfo?: Record<string, unknown>;
};

export interface M365TokenProvider {
  getAccessToken(tokenRequest: TokenRequest): Promise<Result<string, FxError>>;

  getJsonObject(
    tokenRequest: TokenRequest,
    tenantId?: string
  ): Promise<Result<Record<string, unknown>, FxError>>;

  getStatus(tokenRequest: TokenRequest): Promise<Result<LoginStatus, FxError>>;

  signout(): Promise<boolean>;

  switchTenant(tenantId: string): Promise<Result<string, FxError>>;

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

  removeStatusChangeMap(name: string): Promise<Result<boolean, FxError>>;
}

export interface TokenProvider {
  azureAccountProvider: AzureAccountProvider;
  m365TokenProvider: M365TokenProvider;
}
