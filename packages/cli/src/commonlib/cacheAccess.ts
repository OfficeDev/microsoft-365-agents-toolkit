// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

"use strict";

import { TokenCacheContext } from "@azure/msal-node";
import { ConfigFolderName } from "@microsoft/teamsfx-api";
import * as crypto from "crypto";
import * as fs from "fs-extra";
import * as keytarType from "keytar";
import VsCodeLogInstance from "./log";
import * as os from "os";

const cacheDir = os.homedir + `/.${ConfigFolderName}/account`;
const cachePath = os.homedir + `/.${ConfigFolderName}/account/token.cache.`;
const accountPath = os.homedir + `/.${ConfigFolderName}/account/homeId.cache.`;
const tenantPath = os.homedir() + `/.${ConfigFolderName}/account/tenantId.cache.`;
const cachePathEnd = ".json";
const azureSPPath = os.homedir + `/.${ConfigFolderName}/account/azure.sp`;

// the friendly service name to store secret in keytar
const serviceName = "Microsoft Teams Toolkit";

export const UTF8 = "utf8";

export class AccountCrypto {
  private readonly algorithm: crypto.CipherGCMTypes = "aes-256-gcm";
  private readonly accountName: string;
  private readonly keytar?: typeof keytarType;

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
    if (key) {
      const iv = crypto.randomBytes(16);
      const cipher = crypto.createCipheriv(this.algorithm, key, iv);
      const encrypted = Buffer.concat([cipher.update(content, UTF8), cipher.final()]);
      const tag = cipher.getAuthTag();
      return JSON.stringify({
        i: iv.toString("hex"),
        c: encrypted.toString("hex"),
        t: tag.toString("hex"),
      });
    } else {
      // no key, return plain text
      return content;
    }
  }

  public async decrypt(content: string): Promise<string> {
    const key = await this.getKey();
    if (key) {
      const object = JSON.parse(content);
      const decipher = crypto.createDecipheriv(this.algorithm, key, Buffer.from(object.i, "hex"));
      decipher.setAuthTag(Buffer.from(object.t, "hex"));
      const decrpyted = Buffer.concat([
        decipher.update(Buffer.from(object.c, "hex")),
        decipher.final(),
      ]);
      return decrpyted.toString(UTF8);
    } else {
      // no key, return plain text
      return content;
    }
  }

  private async getKey(): Promise<string | undefined> {
    if (this.currentKey) {
      // only return valid key
      return this.currentKey.length === 32 ? this.currentKey : undefined;
    } else {
      try {
        if (this.keytar) {
          let key = await this.keytar.getPassword(serviceName, this.accountName);
          if (!key || key.length !== 32) {
            key = crypto.randomBytes(256).toString("hex").slice(0, 32);
            await this.keytar.setPassword(serviceName, this.accountName, key);

            // validate key again
            const savedKey = await this.keytar.getPassword(serviceName, this.accountName);
            if (savedKey === key) {
              this.currentKey = key;
            }
          } else {
            this.currentKey = key;
          }
        }
      } catch {
        // ignore keytar error and assign an invalid value
        this.currentKey = "Unknown";
      }
    }

    return this.currentKey?.length === 32 ? this.currentKey : undefined;
  }
}

export class CryptoCachePlugin {
  private readonly accountName: string;
  private readonly accountCrypto: AccountCrypto;

  constructor(accountName: string) {
    this.accountName = accountName;
    this.accountCrypto = new AccountCrypto(accountName);
  }

  public async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    await fs.ensureDir(cacheDir);
    const fileCachePath = this.getCachePath();
    if (await fs.pathExists(fileCachePath)) {
      try {
        const text = await fs.readFile(fileCachePath, UTF8);
        if (text && text.length > 0) {
          try {
            const data = await this.accountCrypto.decrypt(text);
            JSON.parse(data);
            cacheContext.tokenCache.deserialize(data);
          } catch (error) {
            // throw error if the plain text is not token cache
            let needThrow = true;
            try {
              const oldObj = JSON.parse(text);
              if (oldObj.Account) {
                needThrow = false;
              }
            } catch {
              // plain text format error
            }

            if (needThrow) {
              await fs.writeFile(fileCachePath, "", UTF8);
              throw error;
            } else {
              cacheContext.tokenCache.deserialize(text);
            }
          }
        }
      } catch (err: any) {
        VsCodeLogInstance.warning(
          "Cannot read token from cache, clean up account cache. " + err.message
        );
      }
    } else {
      try {
        const data = cacheContext.tokenCache.serialize();
        const text = await this.accountCrypto.encrypt(data);
        await fs.writeFile(fileCachePath, text, UTF8);
      } catch (err: any) {
        VsCodeLogInstance.warning(
          "Cannot write token to cache, clean up account cache. " + err.message
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
        await fs.writeFile(fileCachePath, text, UTF8);
      } catch (err: any) {
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
    await fs.writeFile(cachePath + accountName + cachePathEnd, "");
  } catch (err: any) {
    VsCodeLogInstance.warning(
      "Cannot write token to cache, clean up account cache. " + err.message
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
    const data: AzureSPConfig = {
      clientId: clientId,
      secret: secret,
      tenantId: tenantId,
    };
    await fs.writeFile(
      azureSPPath,
      await AzureSpCrypto.accountCrypto.encrypt(JSON.stringify(data)),
      UTF8
    );
  }

  public static async clearAzureSP(): Promise<void> {
    await fs.ensureDir(cacheDir);
    await fs.remove(azureSPPath);
  }

  public static async loadAzureSP(): Promise<AzureSPConfig | undefined> {
    await fs.ensureDir(cacheDir);
    if (await fs.pathExists(azureSPPath)) {
      const data = await AzureSpCrypto.accountCrypto.decrypt(await fs.readFile(azureSPPath, UTF8));
      return JSON.parse(data);
    }

    return undefined;
  }

  public static checkAzureSPFile(): boolean {
    if (fs.existsSync(azureSPPath)) {
      const data = fs.readFileSync(azureSPPath, UTF8);
      const dataJson = JSON.parse(data);
      if (Object.keys(dataJson).length > 0) {
        return true;
      } else {
        return false;
      }
    } else {
      return false;
    }
  }
}

export type AzureSPConfig = {
  clientId: string;
  secret: string;
  tenantId: string;
};
