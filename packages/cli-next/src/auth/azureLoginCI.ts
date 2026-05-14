// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SubscriptionClient } from "@azure/arm-subscriptions";
import * as identity from "@azure/identity";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import {
  type AzureAccountProvider,
  type FxError,
  type ITeamsFxTokenCredential,
  type Result,
  type SubscriptionInfo,
  ok,
  signedIn,
  signedOut,
  UserError,
} from "@microsoft/teamsfx-core-next";
import { saveAzureSP, loadAzureSP, clearAzureSP } from "./cacheAccess";
import { convertTokenToJson } from "./utils";

type LoginStatus = { status: string; token?: string; accountInfo?: Record<string, unknown> };

const MGMT_SCOPE = "https://management.core.windows.net/.default";

/**
 * Azure account provider using service principal credentials (client secret or certificate).
 * Used for headless / CI scenarios.
 */
class AzureLoginCI implements AzureAccountProvider {
  private static instance: AzureLoginCI;
  private credential: ITeamsFxTokenCredential | undefined;
  private clientId = "";
  private secret = "";
  private tenantId = "";
  private subscriptionId: string | undefined;
  private subscriptionName: string | undefined;

  static getInstance(): AzureLoginCI {
    if (!AzureLoginCI.instance) {
      AzureLoginCI.instance = new AzureLoginCI();
    }
    return AzureLoginCI.instance;
  }

  /**
   * Initialize with service principal credentials and persist.
   */
  async init(clientId: string, secret: string, tenantId: string): Promise<void> {
    this.clientId = clientId;
    // Expand ~ in cert paths
    if (secret.startsWith("~")) {
      const expandedPath = path.join(os.homedir(), secret.slice(1));
      this.secret = fs.existsSync(expandedPath) ? expandedPath : secret;
    } else {
      this.secret = secret;
    }
    this.tenantId = tenantId;
    // Validate that credentials work
    await this.getIdentityCredentialAsync();
    await saveAzureSP(clientId, this.secret, tenantId);
  }

  private async load(): Promise<void> {
    const data = await loadAzureSP();
    if (data) {
      this.clientId = data.clientId;
      this.secret = data.secret;
      this.tenantId = data.tenantId;
    }
  }

  async getIdentityCredentialAsync(): Promise<ITeamsFxTokenCredential | undefined> {
    await this.load();
    if (!this.credential) {
      if (fs.existsSync(this.secret)) {
        this.credential = new identity.ClientCertificateCredential(
          this.tenantId,
          this.clientId,
          this.secret
        );
      } else {
        const secretCred = new identity.ClientSecretCredential(
          this.tenantId,
          this.clientId,
          this.secret
        );
        this.credential = new identity.ChainedTokenCredential(secretCred);
      }
    }
    return this.credential;
  }

  async signout(): Promise<boolean> {
    this.credential = undefined;
    await clearAzureSP();
    return true;
  }

  switchTenant(_tenantId: string): Promise<Result<ITeamsFxTokenCredential, FxError>> {
    return Promise.resolve(ok(this.credential!));
  }

  async getStatus(): Promise<LoginStatus> {
    await this.load();
    if (this.clientId && this.secret && this.tenantId) {
      return { status: signedIn };
    }
    return { status: signedOut };
  }

  async getJsonObject(): Promise<Record<string, unknown> | undefined> {
    const cred = await this.getIdentityCredentialAsync();
    const token = await cred?.getToken(MGMT_SCOPE);
    if (token?.token) {
      return convertTokenToJson(token.token);
    }
    return undefined;
  }

  async listSubscriptions(): Promise<SubscriptionInfo[]> {
    const cred = await this.getIdentityCredentialAsync();
    if (!cred) return [];

    const answers: SubscriptionInfo[] = [];
    if (this.tenantId) {
      let subCred: identity.TokenCredential;
      if (fs.existsSync(this.secret)) {
        subCred = new identity.ClientCertificateCredential(
          this.tenantId,
          this.clientId,
          this.secret
        );
      } else {
        subCred = new identity.ClientSecretCredential(this.tenantId, this.clientId, this.secret);
      }
      const client = new SubscriptionClient(subCred);
      for await (const page of client.subscriptions.list().byPage()) {
        for (const sub of page) {
          if (sub.subscriptionId && sub.displayName) {
            answers.push({
              subscriptionId: sub.subscriptionId,
              subscriptionName: sub.displayName,
              tenantId: this.tenantId,
            });
          }
        }
      }
    }
    return answers;
  }

  async setSubscription(subscriptionId: string): Promise<void> {
    const list = await this.listSubscriptions();
    const found = list.find((s) => s.subscriptionId === subscriptionId);
    if (!found) {
      throw new UserError({
        name: "InvalidAzureSubscription",
        message: `Azure subscription '${subscriptionId}' not found for service principal.`,
        source: "login",
      });
    }
    this.tenantId = found.tenantId;
    this.subscriptionId = found.subscriptionId;
    this.subscriptionName = found.subscriptionName;
  }

  getAccountInfo(): Record<string, string> | undefined {
    return this.clientId ? { clientId: this.clientId, tenantId: this.tenantId } : undefined;
  }

  async getSelectedSubscription(): Promise<SubscriptionInfo | undefined> {
    if (this.subscriptionId) {
      return {
        subscriptionId: this.subscriptionId,
        tenantId: this.tenantId,
        subscriptionName: this.subscriptionName ?? "",
      };
    }
    return undefined;
  }

  setStatusChangeMap(): Promise<boolean> {
    return Promise.resolve(true);
  }

  removeStatusChangeMap(): Promise<boolean> {
    return Promise.resolve(true);
  }
}

export default AzureLoginCI;
