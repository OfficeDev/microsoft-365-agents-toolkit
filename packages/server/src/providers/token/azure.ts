// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { MessageConnection } from "vscode-jsonrpc";
import { TokenCredential } from "@azure/core-auth";

import {
  AuthenticationWWWAuthenticateRequest,
  AzureAccountProvider,
  FxError,
  Result,
  SubscriptionInfo,
  TokenRequest,
} from "@microsoft/teamsfx-api";

import { RequestTypes } from "../../apis";
import { getResponseWithErrorHandling } from "../../utils";
import { AccessToken, GetTokenOptions } from "@azure/identity";
import { MFARequiredError, NotImplementedError } from "@microsoft/teamsfx-core";

class TeamsFxTokenCredential implements TokenCredential {
  private connection: MessageConnection;

  constructor(connection: MessageConnection) {
    this.connection = connection;
  }

  async getToken(
    scopes: string | string[],
    options?: GetTokenOptions | undefined
  ): Promise<AccessToken | null> {
    let myScopes: string[] = [];
    if (typeof scopes === "string") {
      myScopes = [scopes];
    } else {
      myScopes = scopes;
    }
    const tokenRequest: TokenRequest = { scopes: myScopes };
    const promise = this.connection.sendRequest(RequestTypes.azure.getAccessToken, tokenRequest);
    const result = await getResponseWithErrorHandling(promise);
    if (result.isErr()) {
      return null;
    } else {
      const accessToken = result.value;
      const tokenJson = ConvertTokenToJson(accessToken);
      return {
        token: accessToken,
        expiresOnTimestamp: tokenJson.exp * 1000,
      };
    }
  }
}

function ConvertTokenToJson(token: string): any {
  const array = token.split(".");
  const buff = Buffer.from(array[1], "base64");
  return JSON.parse(buff.toString("utf8"));
}

export default class ServerAzureAccountProvider implements AzureAccountProvider {
  connection: MessageConnection;
  teamsFxTokenCredential: TeamsFxTokenCredential;
  constructor(connection: MessageConnection) {
    this.connection = connection;
    this.teamsFxTokenCredential = new TeamsFxTokenCredential(this.connection);
  }

  async getIdentityCredentialAsync(
    showDialog?: boolean,
    authenticationSessionRequest?: AuthenticationWWWAuthenticateRequest
  ): Promise<TokenCredential | undefined> {
    if (authenticationSessionRequest && authenticationSessionRequest.wwwAuthenticate) {
      throw new MFARequiredError("FxServer");
    }
    return Promise.resolve(this.teamsFxTokenCredential);
  }

  signout(): Promise<boolean> {
    throw new NotImplementedError("FxServer", `azure/signout`);
  }

  setStatusChangeMap(
    name: string,
    statusChange: (
      status: string,
      token?: string,
      accountInfo?: Record<string, unknown>
    ) => Promise<void>,
    immediateCall?: boolean
  ): Promise<boolean> {
    throw new NotImplementedError("FxServer", `azure/setStatusChangeMap`);
  }

  removeStatusChangeMap(name: string): Promise<boolean> {
    throw new NotImplementedError("FxServer", `azure/removeStatusChangeMap`);
  }

  async getJsonObject(showDialog?: boolean): Promise<Record<string, unknown> | undefined> {
    const promise = this.connection.sendRequest(RequestTypes.azure.getJsonObject);
    const result = await getResponseWithErrorHandling(promise);
    if (result.isErr()) {
      throw result.error;
    }
    return JSON.parse(result.value);
  }

  switchTenant(tenantId: string): Promise<Result<TokenCredential, FxError>> {
    throw new NotImplementedError("FxServer", `azure/switchTenant`);
  }

  async listSubscriptions(): Promise<SubscriptionInfo[]> {
    const promise = this.connection.sendRequest(RequestTypes.azure.listSubscriptions);
    const result = await getResponseWithErrorHandling(promise);
    if (result.isErr()) {
      throw result.error;
    }
    return result.value;
  }

  async setSubscription(subscriptionId: string): Promise<void> {
    const promise = this.connection.sendRequest(RequestTypes.azure.setSubscription, subscriptionId);
    const result = await getResponseWithErrorHandling(promise);
    if (result.isErr()) {
      throw result.error;
    }
  }

  getAccountInfo(): Record<string, string> | undefined {
    throw new NotImplementedError("FxServer", `azure/getAccountInfo`);
  }

  async getSelectedSubscription(triggerUI?: boolean): Promise<SubscriptionInfo | undefined> {
    const promise = this.connection.sendRequest(RequestTypes.azure.getSelectedSubscription);
    const result = await getResponseWithErrorHandling(promise);
    if (result.isErr()) {
      throw result.error;
    }
    return result.value;
  }
}
