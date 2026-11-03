// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { TokenCacheContext } from "@azure/msal-node";
import { ConfigFolderName } from "@microsoft/teamsfx-api";
import * as crypto from "crypto";
import fs from "fs-extra";
import * as os from "os";
import VsCodeLogInstance from "./log";

const cacheDir = os.homedir() + `/.${ConfigFolderName}/account`;
const cachePath = os.homedir() + `/.${ConfigFolderName}/account/token.cache.`;
const accountPath = os.homedir() + `/.${ConfigFolderName}/account/homeId.cache.`;
const tenantPath = os.homedir() + `/.${ConfigFolderName}/account/tenantId.cache.`;
const cachePathEnd = ".json";
const azureSPPath = os.homedir() + `/.${ConfigFolderName}/account/azure.sp`;

// the friendly service name to store secret in keytar
const serviceName = "Microsoft 365 Agents Toolkit";
const secureStorageUnavailableMessage = "Secure credential storage is unavailable.";

export const UTF8 = "utf8";

export class AccountCrypto {
  private readonly algorithm: crypto.CipherGCMTypes = "aes-256-gcm";
  private readonly accountName: string;
  private readonly keytar?: typeof import("keytar");

  private currentKey?: string;

  constructor(accountName: string) {
    this.accountName = accountName;
    try {
      this.keytar = require("keytar");
    } catch {
      // keytar not installed, ingore
    }
  }

  public async encrypt(content: string): Promise<string> {
    const key = await this.getKey();
    const iv = crypto.randomBytes(16);
    const cipher = crypto.createCipheriv(this.algorithm, key, iv);
    const encrypted = Buffer.concat([cipher.update(content, UTF8), cipher.final()]);
    const tag = cipher.getAuthTag();
    return JSON.stringify({
      i: iv.toString("hex"),
      c: encrypted.toString("hex"),
      t: tag.toString("hex"),
    });
  }

  public async decrypt(content: string): Promise<string> {
    const key = await this.getKey();
    const object = JSON.parse(content);
    const decipher = crypto.createDecipheriv(this.algorithm, key, Buffer.from(object.i, "hex"));
    decipher.setAuthTag(Buffer.from(object.t, "hex"));
    const decrpyted = Buffer.concat([
      decipher.update(Buffer.from(object.c, "hex")),
      decipher.final(),
    ]);
    return decrpyted.toString(UTF8);
  }

  private async getKey(): Promise<string> {
    if (this.currentKey) {
      if (this.currentKey.length === 32) {
        return this.currentKey;
      }
      throw new Error(secureStorageUnavailableMessage);
    }

    if (!this.keytar) {
      throw new Error(secureStorageUnavailableMessage);
    }

    try {
      let key = await this.keytar.getPassword(serviceName, this.accountName);
      if (!key || key.length !== 32) {
        key = crypto.randomBytes(32).toString("hex").slice(0, 32);
        await this.keytar.setPassword(serviceName, this.accountName, key);

        const savedKey = await this.keytar.getPassword(serviceName, this.accountName);
        if (savedKey !== key) {
          throw new Error(secureStorageUnavailableMessage);
        }
      }
      this.currentKey = key;
      return key;
    } catch {
      throw new Error(secureStorageUnavailableMessage);
    }
  }
}

export class CryptoCachePlugin {
  private readonly accountName: string;
  private readonly accountCrypto: AccountCrypto;

  constructor(accountName: string, accountCrypto?: AccountCrypto) {
    this.accountName = accountName;
    this.accountCrypto = accountCrypto ?? new AccountCrypto(accountName);
  }

  public async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    await fs.ensureDir(cacheDir);
    const fileCachePath = this.getCachePath();
    if (await fs.pathExists(fileCachePath)) {
      try {
        const text = await fs.readFile(fileCachePath, UTF8);
        if (text && text.length > 0) {
          try {
            let data: string;
            try {
              data = await this.accountCrypto.decrypt(text);
            } catch (error) {
              const oldObj = JSON.parse(text);
              if (!isMsalCache(oldObj)) {
                throw error;
              }
              data = text;
              const encrypted = await this.accountCrypto.encrypt(data);
              await fs.writeFile(fileCachePath, encrypted, { encoding: UTF8, mode: 0o600 });
            }
            JSON.parse(data);
            cacheContext.tokenCache.deserialize(data);
          } catch (error) {
            await fs.remove(fileCachePath);
            throw error;
          }
        }
      } catch (err: any) {
        VsCodeLogInstance.warning(
          "Cannot read token from cache, clean up account cache. " + err.message
        );
      }
    }
  }

  public async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    if (cacheContext.cacheHasChanged) {
      await fs.ensureDir(cacheDir);
      const fileCachePath = this.getCachePath();
      try {
        const data = cacheContext.tokenCache.serialize();
        const text = await this.accountCrypto.encrypt(data);
        await fs.writeFile(fileCachePath, text, { encoding: UTF8, mode: 0o600 });
      } catch (err: any) {
        await fs.remove(fileCachePath);
        VsCodeLogInstance.warning(
          "Cannot save token to cache, clean up account cache. " + err.message
        );
      }
    }
  }

  private getCachePath() {
    return cachePath + this.accountName + cachePathEnd;
  }
}

export async function saveAccountId(accountName: string, accountId?: string) {
  await fs.ensureDir(cacheDir);
  try {
    if (accountId) {
      await fs.writeFile(accountPath + accountName, accountId, UTF8);
    } else {
      // this is to remove current account
      await fs.writeFile(accountPath + accountName, "", UTF8);
    }
  } catch (err: any) {
    VsCodeLogInstance.warning(
      "Cannot save home account id to cache, clean up account cache. " + err.message
    );
  }
}

export async function clearCache(accountName: string) {
  await fs.ensureDir(cacheDir);
  try {
    await fs.remove(cachePath + accountName + cachePathEnd);
  } catch (err: any) {
    VsCodeLogInstance.warning(
      "Cannot remove token cache, clean up account cache. " + err.message
    );
  }
}

export async function loadAccountId(accountName: string) {
  if (await fs.pathExists(accountPath + accountName)) {
    try {
      return await fs.readFile(accountPath + accountName, UTF8);
    } catch (err: any) {
      VsCodeLogInstance.warning(
        "Cannot read home account id from cache, clean up account cache. " + err.message
      );
    }
  }

  return undefined;
}

export async function saveTenantId(accountName: string, tenantId?: string) {
  await fs.ensureDir(cacheDir);
  try {
    if (tenantId) {
      await fs.writeFile(tenantPath + accountName, tenantId, UTF8);
    } else {
      await fs.writeFile(tenantPath + accountName, "", UTF8);
    }
  } catch (err: any) {
    VsCodeLogInstance.warning(
      "Cannot save home tenant id to cache, clean up account cache. " + err.message
    );
  }
}

export async function loadTenantId(accountName: string) {
  if (await fs.pathExists(tenantPath + accountName)) {
    try {
      return await fs.readFile(tenantPath + accountName, UTF8);
    } catch (err: any) {
      VsCodeLogInstance.warning(
        "Cannot read home tenant id from cache, clean up account cache. " + err.message
      );
    }
  }

  return undefined;
}

export class AzureSpCrypto {
  private static readonly accountCrypto: AccountCrypto = new AccountCrypto("azure");

  public static async saveAzureSP(
    clientId: string,
    secret: string,
    tenantId: string
  ): Promise<void> {
    await fs.ensureDir(cacheDir);
    try {
      const data: AzureSPConfig = {
        clientId: clientId,
        secret: secret,
        tenantId: tenantId,
      };
      const encrypted = await AzureSpCrypto.accountCrypto.encrypt(JSON.stringify(data));
      await fs.writeFile(azureSPPath, encrypted, { encoding: UTF8, mode: 0o600 });
    } catch (err: any) {
      await fs.remove(azureSPPath);
      VsCodeLogInstance.warning(
        "Cannot save Azure service principal credentials securely. " + err.message
      );
    }
  }

  public static async clearAzureSP(): Promise<void> {
    await fs.ensureDir(cacheDir);
    await fs.remove(azureSPPath);
  }

  public static async loadAzureSP(): Promise<AzureSPConfig | undefined> {
    await fs.ensureDir(cacheDir);
    if (await fs.pathExists(azureSPPath)) {
      try {
        const text = await fs.readFile(azureSPPath, UTF8);
        const parsed = JSON.parse(text);
        if (isAzureSPConfig(parsed)) {
          const encrypted = await AzureSpCrypto.accountCrypto.encrypt(text);
          await fs.writeFile(azureSPPath, encrypted, { encoding: UTF8, mode: 0o600 });
          return parsed;
        }

        const data = JSON.parse(await AzureSpCrypto.accountCrypto.decrypt(text));
        if (isAzureSPConfig(data)) {
          return data;
        }
      } catch (err: any) {
        await fs.remove(azureSPPath);
        VsCodeLogInstance.warning(
          "Cannot load Azure service principal credentials securely. " + err.message
        );
      }
    }

    return undefined;
  }

  public static checkAzureSPFile(): boolean {
    try {
      return fs.existsSync(azureSPPath) && fs.readFileSync(azureSPPath, UTF8).length > 0;
    } catch {
      return false;
    }
  }
}

function isMsalCache(value: unknown): value is { Account: object } {
  return (
    typeof value === "object" &&
    value !== null &&
    "Account" in value &&
    typeof value.Account === "object" &&
    value.Account !== null
  );
}

function isAzureSPConfig(value: unknown): value is AzureSPConfig {
  return (
    typeof value === "object" &&
    value !== null &&
    "clientId" in value &&
    typeof value.clientId === "string" &&
    "secret" in value &&
    typeof value.secret === "string" &&
    "tenantId" in value &&
    typeof value.tenantId === "string"
  );
}

export type AzureSPConfig = {
  clientId: string;
  secret: string;
  tenantId: string;
};
