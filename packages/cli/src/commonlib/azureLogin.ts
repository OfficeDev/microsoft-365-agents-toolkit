// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { SubscriptionClient } from "@azure/arm-subscriptions";
import { AccessToken, GetTokenOptions, TokenCredential } from "@azure/core-auth";
import { LogLevel } from "@azure/msal-node";
import {
  AzureAccountProvider,
  ConfigFolderName,
  FxError,
  LogLevel as LLevel,
  ok,
  OptionItem,
  Result,
  SingleSelectConfig,
  SubscriptionInfo,
  UserError,
} from "@microsoft/teamsfx-api";
import {
  AzureScopes,
  isValidProjectV3,
  InvalidAzureSubscriptionError,
  featureFlagManager,
  FeatureFlags,
} from "@microsoft/teamsfx-core";
import * as fs from "fs-extra";
import * as path from "path";
import CLIUIInstance from "../userInteraction";
import { AzureSpCrypto, CryptoCachePlugin, loadTenantId, saveTenantId } from "./cacheAccess";
import { CodeFlowLogin, ConvertTokenToJson, checkIsOnline } from "./codeFlowLogin";
import {
  MFACode,
  changeLoginTenantMessage,
  env,
  envDefaultJsonFile,
  failToFindSubscription,
  loginComponent,
  noSubscriptionFound,
  selectSubscription,
  signedIn,
  signedOut,
  subscription,
  subscriptionInfoFile,
} from "./common/constant";
import { LoginStatus, login } from "./common/login";
import CLILogProvider from "./log";
import { MemoryCache } from "./memoryCache";
import ui from "../userInteraction";

const accountName = "azure";
const scopes = ["https://management.core.windows.net/user_impersonation"];
const SERVER_PORT = 0;

const cachePlugin = new CryptoCachePlugin(accountName);

function getConfig(tenantId?: string) {
  let authority;
  if (tenantId && tenantId.length > 0) {
    authority = "https://login.microsoftonline.com/" + tenantId;
  } else {
    authority = "https://login.microsoftonline.com/organizations";
  }
  const config = {
    auth: {
      clientId: "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0",
      authority: authority,
    },
    system: {
      loggerOptions: {
        loggerCallback(loglevel: any, message: any, containsPii: any) {
          if (this.logLevel <= LogLevel.Error) {
            void CLILogProvider.log(4 - loglevel, message);
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
  return config;
}

// eslint-disable-next-line
// @ts-ignore
const memoryDictionary: { [tenantId: string]: MemoryCache } = {};

class TeamsFxTokenCredential implements TokenCredential {
  private codeFlowInstance: CodeFlowLogin;
  private tenantId: string;

  constructor(codeFlowInstance: CodeFlowLogin) {
    this.codeFlowInstance = codeFlowInstance;
    this.tenantId = "";
  }

  public setTenantId(tenantId: string) {
    this.tenantId = tenantId;
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
    const tokenRes: Result<string, FxError> = await this.codeFlowInstance.getTokenByScopes(
      myScopes,
      true,
      this.tenantId
    );

    if (tokenRes.isOk()) {
      const tokenJson = ConvertTokenToJson(tokenRes.value);
      return {
        token: tokenRes.value,
        expiresOnTimestamp: tokenJson.exp * 1000,
      };
    } else {
      return null;
    }
  }
}

export class AzureAccountManager extends login implements AzureAccountProvider {
  private static instance: AzureAccountManager;
  private static codeFlowInstance: CodeFlowLogin;
  // default tenantId
  private static domain: string | undefined;
  private static username: string | undefined;
  private static subscriptionId: string | undefined;
  private static subscriptionName: string | undefined;
  private static rootPath: string | undefined;
  //user set tenantId
  private static tenantId: string | undefined;
  private static teamsFxTokenCredential: TeamsFxTokenCredential;

  private static statusChange?: (
    status: string,
    token?: string,
    accountInfo?: Record<string, unknown>
  ) => Promise<void>;

  private constructor() {
    super();
    AzureAccountManager.codeFlowInstance = new CodeFlowLogin(
      scopes,
      getConfig(),
      SERVER_PORT,
      accountName
    );
    AzureAccountManager.teamsFxTokenCredential = new TeamsFxTokenCredential(
      AzureAccountManager.codeFlowInstance
    );
  }

  /**
   * Gets instance
   * @returns instance
   */
  public static getInstance(): AzureAccountManager {
    if (!AzureAccountManager.instance) {
      AzureAccountManager.instance = new AzureAccountManager();
    }

    return AzureAccountManager.instance;
  }

  /**
   * Async get identity [crendential](https://github.com/Azure/azure-sdk-for-js/blob/master/sdk/core/core-auth/src/tokenCredential.ts)
   */
  getIdentityCredentialAsync(showDialog = true): Promise<TokenCredential | undefined> {
    return Promise.resolve(AzureAccountManager.teamsFxTokenCredential);
  }

  async switchTenant(tenantId: string): Promise<Result<TokenCredential, FxError>> {
    await saveTenantId(accountName, tenantId);
    return Promise.resolve(ok(AzureAccountManager.teamsFxTokenCredential));
  }

  async getTenant(): Promise<string | undefined> {
    return await loadTenantId(accountName);
  }

  private async updateLoginStatus(): Promise<void> {
    const checkCodeFlow =
      AzureAccountManager.codeFlowInstance !== undefined &&
      AzureAccountManager.codeFlowInstance.account;
    if (AzureAccountManager.statusChange !== undefined && checkCodeFlow) {
      const credential = await this.getIdentityCredentialAsync();
      const accessToken = await credential?.getToken(AzureScopes);
      const accountJson = await this.getJsonObject();
      await AzureAccountManager.statusChange("SignedIn", accessToken?.token, accountJson);
    }
    await this.notifyStatus();
  }

  private async login(showDialog: boolean, tenantId?: string): Promise<void> {
    const accessToken = await AzureAccountManager.codeFlowInstance.getTokenByScopes(AzureScopes);
    const tokenJson = await this.getJsonObject(false, tenantId);
    if (accessToken.isOk() && accessToken.value) {
      this.setMemoryCache(accessToken.value, tokenJson);
    }
  }

  private setMemoryCache(accessToken: string | undefined, tokenJson: any) {
    if (accessToken) {
      if (!AzureAccountManager.domain) {
        AzureAccountManager.domain = (tokenJson as any).tid;
      }
      AzureAccountManager.username = (tokenJson as any).upn ?? (tokenJson as any).unique_name;
      tokenJson = ConvertTokenToJson(accessToken);
      const tokenExpiresIn =
        Math.round(new Date().getTime() / 1000) - ((tokenJson as any).iat as number);
      if (!memoryDictionary[(tokenJson as any).tid]) {
        // eslint-disable-next-line
        // @ts-ignore
        memoryDictionary[(tokenJson as any).tid] = new MemoryCache();
      }

      memoryDictionary[(tokenJson as any).tid].add(
        [
          {
            tokenType: "Bearer",
            expiresIn: tokenExpiresIn,
            expiresOn: {},
            resource: env.activeDirectoryResourceId,
            accessToken: accessToken,
            userId: (tokenJson as any).upn ?? (tokenJson as any).unique_name,
            _clientId: getConfig().auth.clientId,
            _authority: env.activeDirectoryEndpointUrl + (tokenJson as any).tid,
          },
        ],
        function () {
          const _ = 1;
        }
      );
    }
  }

  async getJsonObject(
    showDialog = true,
    tenantId?: string
  ): Promise<Record<string, unknown> | undefined> {
    const token = await AzureAccountManager.codeFlowInstance.getTokenByScopes(
      AzureScopes,
      true,
      tenantId
    );

    if (token.isOk() && token.value) {
      const array = token.value!.split(".");
      const buff = Buffer.from(array[1], "base64");
      return Promise.resolve(JSON.parse(buff.toString("utf-8")));
    } else {
      return Promise.resolve(undefined);
    }
  }

  /**
   * singnout from Azure
   */
  async signout(): Promise<boolean> {
    AzureAccountManager.codeFlowInstance.account = undefined;
    if (AzureAccountManager.statusChange !== undefined) {
      await AzureAccountManager.statusChange("SignedOut", undefined, undefined);
    }
    await AzureAccountManager.codeFlowInstance.logout();
    await this.notifyStatus();
    await AzureSpCrypto.clearAzureSP();
    return Promise.resolve(true);
  }

  async getStatus(): Promise<LoginStatus> {
    if (!AzureAccountManager.codeFlowInstance.account) {
      await AzureAccountManager.codeFlowInstance.reloadCache();
    }
    if (AzureAccountManager.codeFlowInstance.account) {
      const loginToken = await AzureAccountManager.codeFlowInstance.getTokenByScopes(scopes, false);
      if (!loginToken) {
        if (await checkIsOnline()) {
          return Promise.resolve({ status: signedOut, token: undefined, accountInfo: undefined });
        } else {
          return Promise.resolve({
            status: signedIn,
            token: undefined,
            accountInfo: { upn: AzureAccountManager.codeFlowInstance.account?.username },
          });
        }
      }
      const credential = await this.getIdentityCredentialAsync();
      const token = await credential?.getToken(AzureScopes);
      const accountJson = await this.getJsonObject();
      return Promise.resolve({
        status: signedIn,
        token: token?.token,
        accountInfo: accountJson,
      });
    } else {
      return Promise.resolve({ status: signedOut, token: undefined, accountInfo: undefined });
    }
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

  async listSubscriptions(): Promise<SubscriptionInfo[]> {
    const arr: SubscriptionInfo[] = [];
    if (AzureAccountManager.teamsFxTokenCredential) {
      let showMFA = true;
      if (!AzureAccountManager.tenantId) {
        const tenantClient = new SubscriptionClient(AzureAccountManager.teamsFxTokenCredential);
        const tenantTokenCredential: TeamsFxTokenCredential = new TeamsFxTokenCredential(
          AzureAccountManager.codeFlowInstance
        );
        const cachedTenantId = await loadTenantId(accountName);
        for await (const page of tenantClient.tenants.list().byPage({ maxPageSize: 100 })) {
          for (const tenant of page) {
            if (
              cachedTenantId
                ? tenant.tenantId && tenant.tenantId == cachedTenantId
                : tenant.tenantId
            ) {
              try {
                tenantTokenCredential.setTenantId(tenant.tenantId as string);
                const subscriptionClient = new SubscriptionClient(tenantTokenCredential);
                for await (const subPage of subscriptionClient.subscriptions
                  .list()
                  .byPage({ maxPageSize: 100 })) {
                  for (const item of subPage) {
                    arr.push({
                      subscriptionId: item.subscriptionId!,
                      subscriptionName: item.displayName!,
                      tenantId: tenant.tenantId as string,
                    });
                  }
                }
              } catch (error: any) {
                if (error.message.indexOf(MFACode) >= 0) {
                  if (showMFA) {
                    CLILogProvider.necessaryLog(LLevel.Info, changeLoginTenantMessage);
                    showMFA = false;
                  }
                  CLILogProvider.necessaryLog(LLevel.Info, tenant.tenantId as string);
                }
              }
            }
          }
        }
      } else {
        AzureAccountManager.teamsFxTokenCredential.setTenantId(AzureAccountManager.tenantId);
        const subscriptionClient = new SubscriptionClient(
          AzureAccountManager.teamsFxTokenCredential
        );
        for await (const page of subscriptionClient.subscriptions
          .list()
          .byPage({ maxPageSize: 100 })) {
          for (const item of page) {
            arr.push({
              subscriptionId: item.subscriptionId!,
              subscriptionName: item.displayName!,
              tenantId: AzureAccountManager.tenantId,
            });
          }
        }
      }
    }
    return arr;
  }

  async setSubscription(subscriptionId: string): Promise<void> {
    if (subscriptionId === "") {
      AzureAccountManager.tenantId = undefined;
      AzureAccountManager.subscriptionId = undefined;
      AzureAccountManager.subscriptionName = undefined;
      return;
    }
    const list = await this.listSubscriptions();
    for (let i = 0; i < list.length; ++i) {
      const item = list[i];
      if (item.subscriptionId === subscriptionId) {
        AzureAccountManager.tenantId = item.tenantId;
        AzureAccountManager.teamsFxTokenCredential.setTenantId(item.tenantId);
        AzureAccountManager.subscriptionId = item.subscriptionId;
        AzureAccountManager.subscriptionName = item.subscriptionName;
        return;
      }
    }
    throw new InvalidAzureSubscriptionError(subscriptionId);
  }

  getAccountInfo(): Record<string, string> | undefined {
    if (AzureAccountManager.codeFlowInstance.account) {
      return this.getJsonObject() as unknown as Record<string, string>;
    } else {
      return undefined;
    }
  }

  async getSelectedSubscription(triggerUI = false): Promise<SubscriptionInfo | undefined> {
    if (triggerUI) {
      if (!AzureAccountManager.codeFlowInstance.account) {
        await this.login(false);
      }
      if (AzureAccountManager.codeFlowInstance.account && !AzureAccountManager.subscriptionId) {
        const subscriptionList = await this.listSubscriptions();
        if (!subscriptionList || subscriptionList.length === 0) {
          throw new UserError(loginComponent, noSubscriptionFound, failToFindSubscription);
        }
        if (subscriptionList && subscriptionList.length === 1) {
          await this.setSubscription(subscriptionList[0].subscriptionId);
        } else if (subscriptionList.length > 1) {
          const options: OptionItem[] = subscriptionList.map((sub) => {
            return {
              id: sub.subscriptionId,
              label: sub.subscriptionName,
              data: sub.tenantId,
            } as OptionItem;
          });
          const config: SingleSelectConfig = {
            name: subscription,
            title: selectSubscription,
            options: options,
          };
          const result = await CLIUIInstance.selectOption(config);
          if (result.isErr()) {
            throw result.error;
          } else {
            const subId = result.value.result as string;
            await this.setSubscription(subId);
          }
        }
      }
    } else {
      if (AzureAccountManager.codeFlowInstance.account && !AzureAccountManager.subscriptionId) {
        const subscriptionList = await this.listSubscriptions();
        if (subscriptionList && subscriptionList.length === 1) {
          await this.setSubscription(subscriptionList[0].subscriptionId);
        }
      }
    }
    if (AzureAccountManager.codeFlowInstance.account && AzureAccountManager.subscriptionId) {
      const selectedSub: SubscriptionInfo = {
        subscriptionId: AzureAccountManager.subscriptionId,
        tenantId: AzureAccountManager.tenantId!,
        subscriptionName: AzureAccountManager.subscriptionName ?? "",
      };
      return selectedSub;
    } else {
      return undefined;
    }
  }

  public setRootPath(rootPath: string): void {
    AzureAccountManager.rootPath = rootPath;
  }

  async saveSubscription(subscriptionInfo: SubscriptionInfo): Promise<void> {
    const subscriptionFilePath = await this.getSubscriptionInfoPath();
    if (!subscriptionFilePath) {
      return;
    } else {
      await fs.writeFile(subscriptionFilePath, JSON.stringify(subscriptionInfo, null, 4));
    }
  }

  readSubscription(): Promise<SubscriptionInfo | undefined> {
    return Promise.resolve(undefined);
  }

  getSubscriptionInfoPath(): Promise<string | undefined> {
    if (AzureAccountManager.rootPath) {
      if (isValidProjectV3(AzureAccountManager.rootPath)) {
        const subscriptionFile = path.join(
          AzureAccountManager.rootPath,
          `.${ConfigFolderName}`,
          subscriptionInfoFile
        );
        return Promise.resolve(subscriptionFile);
      } else {
        return Promise.resolve(undefined);
      }
    } else {
      return Promise.resolve(undefined);
    }
  }

  async getSubscriptionInfoFromEnv(): Promise<SubscriptionInfo | undefined> {
    if (AzureAccountManager.rootPath) {
      if (!isValidProjectV3(AzureAccountManager.rootPath)) {
        return undefined;
      }
      const envDefalultFile = path.join(
        AzureAccountManager.rootPath,
        `.${ConfigFolderName}`,
        envDefaultJsonFile
      );
      if (!fs.existsSync(envDefalultFile)) {
        return undefined;
      }
      const envDefaultJson = (await fs.readFile(envDefalultFile)).toString();
      const envDefault = JSON.parse(envDefaultJson);
      if (envDefault.solution && envDefault.solution.subscriptionId) {
        return {
          subscriptionId: envDefault.solution.subscriptionId,
          tenantId: envDefault.solution.tenantId,
          subscriptionName: "",
        };
      } else {
        return undefined;
      }
    } else {
      return undefined;
    }
  }
}

interface PartialList<T> extends Array<T> {
  nextLink?: string;
}

// Copied from https://github.com/microsoft/vscode-azure-account/blob/2b3c1a8e81e237580465cc9a1f4da5caa34644a6/sample/src/extension.ts
// to list all subscriptions
async function listAll<T>(
  client: { listNext(nextPageLink: string): Promise<PartialList<T>> },
  first: Promise<PartialList<T>>
): Promise<T[]> {
  const all: T[] = [];
  for (
    let list = await first;
    list.length || list.nextLink;
    list = list.nextLink ? await client.listNext(list.nextLink) : []
  ) {
    all.push(...list);
  }
  return all;
}

import AzureLoginCI from "./azureLoginCI";
import AzureAccountProviderUserPassword from "./azureLoginUserPassword";

// todo delete ciEnabled
const azureLogin = !ui.interactive
  ? AzureSpCrypto.checkAzureSPFile()
    ? AzureLoginCI
    : AzureAccountProviderUserPassword
  : AzureSpCrypto.checkAzureSPFile()
  ? AzureLoginCI
  : AzureAccountManager.getInstance();

export default azureLogin;

// todo merge with default export, this function fix bug when user already logins with service principal, and he logins interactively, default azureLogin will return azureLoginCIProvider
export function getAzureProvider() {
  if (AzureSpCrypto.checkAzureSPFile()) {
    return AzureLoginCI;
  }
  return !ui.interactive && process.env.AZURE_ACCOUNT_NAME && process.env.AZURE_ACCOUNT_PASSWORD
    ? AzureAccountProviderUserPassword
    : AzureAccountManager.getInstance();
}
