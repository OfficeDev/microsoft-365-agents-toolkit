// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  PublicClientApplication,
  LogLevel,
  TokenCacheContext,
} from "@azure/msal-node";
import * as crypto from "crypto";
import * as fs from "fs-extra";
import * as keytarType from "keytar";
import * as os from "os";

const extensionConfigFolderName = "fx";
const cacheDir = os.homedir + `/.${extensionConfigFolderName}/account`;
const cachePath =
  os.homedir + `/.${extensionConfigFolderName}/account/token.cache.`;
const accountPath =
  os.homedir + `/.${extensionConfigFolderName}/account/homeId.cache.`;
const homeAccountId = "home_account_id";
const UTF8 = "utf-8";
const cachePathEnd = ".json";

// the friendly service name to store secret in keytar
const serviceName = "Microsoft Teams Toolkit";

const accountName = "m365";
const scopes = ["https://dev.teams.microsoft.com/AppDefinitions.ReadWrite"];

const beforeCacheAccess = getBeforeCacheAccess(accountName);
const afterCacheAccess = getAfterCacheAccess(scopes, accountName);

const cachePluginNonWindows = {
  beforeCacheAccess,
  afterCacheAccess,
};

function getBeforeCacheAccess(accountName: string) {
  const beforeCacheAccess = async (cacheContext: any): Promise<void> => {
    return new Promise(async (resolve, reject) => {
      await fs.ensureDir(cacheDir);
      const fileCachePath = getCachePath(accountName);
      if (fs.existsSync(fileCachePath)) {
        fs.readFile(fileCachePath, UTF8, (err, data) => {
          try {
            JSON.parse(data);
          } catch (error) {
            fs.writeFileSync(fileCachePath, "", UTF8);
          }
          if (err) {
            console.error("read token fail: " + err.message);
            reject();
          } else {
            cacheContext.tokenCache.deserialize(data);
            resolve();
          }
        });
      } else {
        fs.writeFile(
          fileCachePath,
          cacheContext.tokenCache.serialize(),
          (err) => {
            if (err) {
              console.error("write token fail: " + err.message);
              reject();
            }
          }
        );
        resolve();
      }
    });
  };
  return beforeCacheAccess;
}

function getAfterCacheAccess(scopes: string[], accountName: string) {
  const afterCacheAccess = async (cacheContext: any) => {
    if (cacheContext.cacheHasChanged) {
      const fileCachePath = getCachePath(accountName);
      // save token
      fs.writeFile(
        fileCachePath,
        cacheContext.tokenCache.serialize(),
        (err) => {
          if (err) {
            console.error("save token fail: " + err.message);
          }
        }
      );
    }
  };
  return afterCacheAccess;
}

function getCachePath(accountName: string) {
  return cachePath + accountName + cachePathEnd;
}

class AccountCrypto {
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
      const encrypted = Buffer.concat([
        cipher.update(content, UTF8),
        cipher.final(),
      ]);
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
      const decipher = crypto.createDecipheriv(
        this.algorithm,
        key,
        Buffer.from(object.i, "hex")
      );
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
          let key = await this.keytar.getPassword(
            serviceName,
            this.accountName
          );
          if (!key || key.length !== 32) {
            key = crypto.randomBytes(256).toString("hex").slice(0, 32);
            await this.keytar.setPassword(serviceName, this.accountName, key);

            // validate key again
            const savedKey = await this.keytar.getPassword(
              serviceName,
              this.accountName
            );
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

  public async beforeCacheAccess(
    cacheContext: TokenCacheContext
  ): Promise<void> {
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
        console.error("read token fail: " + err.message);
      }
    } else {
      try {
        const data = cacheContext.tokenCache.serialize();
        const text = await this.accountCrypto.encrypt(data);
        await fs.writeFile(fileCachePath, text, UTF8);
      } catch (err: any) {
        console.error("write token fail: " + err.message);
      }
    }
  }

  public async afterCacheAccess(
    cacheContext: TokenCacheContext
  ): Promise<void> {
    if (cacheContext.cacheHasChanged) {
      await fs.ensureDir(cacheDir);
      const fileCachePath = this.getCachePath();
      try {
        const data = cacheContext.tokenCache.serialize();
        const text = await this.accountCrypto.encrypt(data);
        await fs.writeFile(fileCachePath, text, UTF8);
      } catch (err: any) {
        console.error("save token fail: " + err.message);
      }
    }
  }

  private getCachePath() {
    return cachePath + this.accountName + cachePathEnd;
  }
}

// Entry
const username = process.argv[2];
const password = process.argv[3];

if (!username || !password) {
  console.error(
    `Please provide username and password, e.g.,${os.EOL}\t npx ts-node m365Login.ts -- "username" "password"`
  );
  process.exit(-1);
}

const config = {
  auth: {
    clientId: "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0",
    authority: "https://login.microsoftonline.com/organizations",
  },
  system: {
    loggerOptions: {
      piiLoggingEnabled: false,
      logLevel: LogLevel.Verbose,
    },
  },
  cache: {
    // The account cache will be saved to a file with no encryption since VS Code keytar is not available on Windows. Teams toolkit will use plain text when decryption fails.
    cachePlugin: cachePluginNonWindows,
  },
};

const pca = new PublicClientApplication(config);

console.log(`Login as User: ${username}.`);

pca
  .acquireTokenByUsernamePassword({
    scopes: scopes,
    username: username,
    // Need to encode password for special characters to workaround the MSAL bug:
    // https://github.com/AzureAD/microsoft-authentication-library-for-js/issues/4326#issuecomment-995109619
    password: encodeURIComponent(password),
  })
  .then((result) => {
    fs.ensureDirSync(cacheDir);
    fs.writeFileSync(
      accountPath + accountName,
      result!.account!.homeAccountId,
      UTF8
    );
    console.log("Login Successfully!");
  })
  .catch((error) => {
    console.error(error);
    process.exit(-1);
  });
