// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { Subscription, SubscriptionClient, TenantIdDescription } from "@azure/arm-subscriptions";
import { TokenCredential } from "@azure/core-auth";
import * as identity from "@azure/identity";
import dotenv from "dotenv";

import {
  AzureAccountProvider,
  FxError,
  Result,
  SubscriptionInfo,
  UserError,
} from "@microsoft/teamsfx-api";

import * as cfg from "./common/userPasswordConfig";
import { AzureScopes } from "@microsoft/teamsfx-core";
import { ConvertTokenToJson } from "./codeFlowLogin";

dotenv.config();

const user = cfg.AZURE_ACCOUNT_NAME || "";
const password = cfg.AZURE_ACCOUNT_PASSWORD || "";

type LoginStatus = {
  status: string;
  token?: string;
  accountInfo?: Record<string, unknown>;
};

export class AzureAccountProviderUserPassword implements AzureAccountProvider {
  static tokenCredential: TokenCredential;

  private static instance: AzureAccountProviderUserPassword;

  /**
   * Gets instance
   * @returns instance
   */
  public static getInstance(): AzureAccountProviderUserPassword {
    if (!AzureAccountProviderUserPassword.instance) {
      AzureAccountProviderUserPassword.instance = new AzureAccountProviderUserPassword();
    }

    return AzureAccountProviderUserPassword.instance;
  }

  async getIdentityCredentialAsync(): Promise<TokenCredential | undefined> {
    if (AzureAccountProviderUserPassword.tokenCredential == undefined) {
      const identityCredential = new identity.UsernamePasswordCredential(
        cfg.AZURE_TENANT_ID || "organizations",
        "04b07795-8ddb-461a-bbee-02f9e1bf7b46",
        user,
        password
      );
      const credentialChain = new identity.ChainedTokenCredential(identityCredential);
      AzureAccountProviderUserPassword.tokenCredential = credentialChain;
    }

    return new Promise((resolve) => {
      resolve(AzureAccountProviderUserPassword.tokenCredential);
    });
  }

  public async signout(): Promise<boolean> {
    return new Promise((resolve) => {
      resolve(true);
    });
  }

  public switchTenant(tenantId: string): Promise<Result<TokenCredential, FxError>> {
    throw new Error("Method not implemented.");
  }

  async getTenant(): Promise<string | undefined> {
    return Promise.resolve(cfg.AZURE_TENANT_ID);
  }

  public async getStatus(): Promise<LoginStatus> {
    return Promise.resolve({
      status: "SignedIn",
    });
  }

  setStatusChangeMap(
    name: string,
    statusChange: (
      status: string,
      token?: string,
      accountInfo?: Record<string, unknown>
    ) => Promise<void>
  ): Promise<boolean> {
    throw new Error("Method not implemented.");
  }
  removeStatusChangeMap(name: string): Promise<boolean> {
    throw new Error("Method not implemented.");
  }
  async getJsonObject(showDialog?: boolean): Promise<Record<string, unknown> | undefined> {
    const identity = await this.getIdentityCredentialAsync();
    const token = await identity?.getToken(AzureScopes);
    if (token?.token) {
      return ConvertTokenToJson(token?.token);
    } else {
      return undefined;
    }
  }

  async listSubscriptions(): Promise<SubscriptionInfo[]> {
    const credential = await this.getIdentityCredentialAsync();
    if (credential) {
      const client = new SubscriptionClient(credential);
      const tenants: TenantIdDescription[] = [];
      for await (const page of client.tenants.list().byPage({ maxPageSize: 100 })) {
        for (const tenant of page) {
          tenants.push(tenant);
        }
      }
      let answers: SubscriptionInfo[] = [];
      for (const tenant of tenants) {
        if (tenant.tenantId) {
          const cred = new identity.UsernamePasswordCredential(
            tenant.tenantId,
            "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0",
            user,
            password
          );
          const client = new SubscriptionClient(cred);
          const subscriptions: Subscription[] = [];
          for await (const page of client.subscriptions.list().byPage({ maxPageSize: 100 })) {
            for (const subscription of page) {
              subscriptions.push(subscription);
            }
          }
          const filteredsubs = subscriptions.filter(
            (sub) => !!sub.displayName && !!sub.subscriptionId
          );
          answers = answers.concat(
            filteredsubs.map((sub) => {
              return {
                subscriptionName: sub.displayName!,
                subscriptionId: sub.subscriptionId!,
                tenantId: tenant.tenantId!,
              };
            })
          );
        }
      }
      return answers;
    }
    return [];
  }

  async setSubscription(subscriptionId: string): Promise<void> {
    const list = await this.listSubscriptions();
    for (let i = 0; i < list.length; ++i) {
      const item = list[i];
      if (item.subscriptionId == subscriptionId) {
        return;
      }
    }
    throw new UserError(
      "CI",
      "NotFoundSubscriptionId",
      "Inputed subscription not found in your tenant"
    );
  }

  // For now, cli no need to get account information through this method
  getAccountInfo(): Record<string, string> | undefined {
    return {};
  }

  getSelectedSubscription(): Promise<SubscriptionInfo | undefined> {
    const selectedSub: SubscriptionInfo = {
      subscriptionId: "",
      tenantId: "",
      subscriptionName: "",
    };
    if (cfg.AZURE_TENANT_ID) {
      selectedSub.tenantId = cfg.AZURE_TENANT_ID;
    }
    if (cfg.AZURE_SUBSCRIPTION_ID) {
      selectedSub.subscriptionId = cfg.AZURE_SUBSCRIPTION_ID;
    }
    return Promise.resolve(selectedSub);
  }

  async selectSubscription(subscriptionId?: string): Promise<string | undefined> {
    if (subscriptionId) {
      await this.setSubscription(subscriptionId);
      return Promise.resolve(subscriptionId);
    } else {
      return Promise.resolve(undefined);
    }
  }

  public setRootPath(rootPath: string): void {}

  readSubscription(): Promise<SubscriptionInfo | undefined> {
    return Promise.resolve(undefined);
  }

  getSubscriptionInfoFromEnv(): Promise<SubscriptionInfo | undefined> {
    return Promise.resolve(undefined);
  }
}

export type AzureSubscription = {
  displayName: string;
  subscriptionId: string;
};

export default AzureAccountProviderUserPassword.getInstance();
