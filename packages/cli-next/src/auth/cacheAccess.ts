// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TokenCacheContext } from "@azure/msal-node";
import * as crypto from "crypto";
import * as fs from "fs";
import * as path from "path";
import * as os from "os";
import { KEYTAR_SERVICE } from "./constants";

const CACHE_DIR = path.join(os.homedir(), ".fx", "account");
const CACHE_PATH_PREFIX = path.join(CACHE_DIR, "token.cache.");
const ACCOUNT_PATH_PREFIX = path.join(CACHE_DIR, "homeId.cache.");
const TENANT_PATH_PREFIX = path.join(CACHE_DIR, "tenantId.cache.");
const AZURE_SP_PATH = path.join(CACHE_DIR, "azure.sp");
const UTF8 = "utf8" as const;

function ensureCacheDir(): void {
  fs.mkdirSync(CACHE_DIR, { recursive: true });
}

/**
 * AES-256-GCM encryption/decryption with keytar-stored key.
 */
export class AccountCrypto {
  private readonly algorithm: crypto.CipherGCMTypes = "aes-256-gcm";
  private readonly accountName: string;
  private keytar: typeof import("keytar") | undefined;
  private currentKey: string | undefined;

  constructor(accountName: string) {
    this.accountName = accountName;
    try {
      this.keytar = require("keytar");
    } catch {
      // keytar not installed — tokens stored unencrypted
    }
  }

  async encrypt(content: string): Promise<string> {
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
    }
    return content;
  }

  async decrypt(content: string): Promise<string> {
    const key = await this.getKey();
    if (key) {
      const obj = JSON.parse(content);
      const decipher = crypto.createDecipheriv(this.algorithm, key, Buffer.from(obj.i, "hex"));
      decipher.setAuthTag(Buffer.from(obj.t, "hex"));
      const decrypted = Buffer.concat([
        decipher.update(Buffer.from(obj.c, "hex")),
        decipher.final(),
      ]);
      return decrypted.toString(UTF8);
    }
    return content;
  }

  private async getKey(): Promise<string | undefined> {
    if (this.currentKey) {
      return this.currentKey.length === 32 ? this.currentKey : undefined;
    }
    try {
      if (this.keytar) {
        let key = await this.keytar.getPassword(KEYTAR_SERVICE, this.accountName);
        if (!key || key.length !== 32) {
          key = crypto.randomBytes(256).toString("hex").slice(0, 32);
          await this.keytar.setPassword(KEYTAR_SERVICE, this.accountName, key);
          // Validate key was stored correctly
          const savedKey = await this.keytar.getPassword(KEYTAR_SERVICE, this.accountName);
          if (savedKey === key) {
            this.currentKey = key;
          }
        } else {
          this.currentKey = key;
        }
      }
    } catch {
      this.currentKey = "Unknown";
    }
    return this.currentKey?.length === 32 ? this.currentKey : undefined;
  }
}

/**
 * MSAL cache plugin that encrypts tokens at rest.
 */
export class CryptoCachePlugin {
  private readonly accountName: string;
  private readonly crypto: AccountCrypto;

  constructor(accountName: string) {
    this.accountName = accountName;
    this.crypto = new AccountCrypto(accountName);
  }

  async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    ensureCacheDir();
    const filePath = this.getCachePath();
    let text: string | undefined;
    try {
      text = fs.readFileSync(filePath, UTF8);
    } catch {
      // File does not exist or is unreadable — write initial cache
      try {
        const data = cacheContext.tokenCache.serialize();
        const encrypted = await this.crypto.encrypt(data);
        fs.writeFileSync(filePath, encrypted, UTF8);
      } catch {
        // Swallow cache write errors
      }
      return;
    }

    if (!text || text.length === 0) return;

    try {
      const data = await this.crypto.decrypt(text);
      JSON.parse(data); // validate JSON
      cacheContext.tokenCache.deserialize(data);
    } catch {
      // Try reading as unencrypted legacy cache
      try {
        const parsed = JSON.parse(text);
        if (parsed.Account) {
          cacheContext.tokenCache.deserialize(text);
        } else {
          fs.writeFileSync(filePath, "", UTF8);
        }
      } catch {
        fs.writeFileSync(filePath, "", UTF8);
      }
    }
  }

  async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
    if (cacheContext.cacheHasChanged) {
      ensureCacheDir();
      try {
        const data = cacheContext.tokenCache.serialize();
        const text = await this.crypto.encrypt(data);
        fs.writeFileSync(this.getCachePath(), text, UTF8);
      } catch {
        // Swallow cache write errors
      }
    }
  }

  private getCachePath(): string {
    return CACHE_PATH_PREFIX + this.accountName + ".json";
  }
}

// ── Simple file-based helpers ──────────────────────────────────────────

export async function saveAccountId(accountName: string, accountId?: string): Promise<void> {
  ensureCacheDir();
  try {
    fs.writeFileSync(ACCOUNT_PATH_PREFIX + accountName, accountId ?? "", UTF8);
  } catch {
    // ignore
  }
}

export async function loadAccountId(accountName: string): Promise<string | undefined> {
  const p = ACCOUNT_PATH_PREFIX + accountName;
  if (fs.existsSync(p)) {
    try {
      const val = fs.readFileSync(p, UTF8);
      return val || undefined;
    } catch {
      // ignore
    }
  }
  return undefined;
}

export async function saveTenantId(accountName: string, tenantId?: string): Promise<void> {
  ensureCacheDir();
  try {
    fs.writeFileSync(TENANT_PATH_PREFIX + accountName, tenantId ?? "", UTF8);
  } catch {
    // ignore
  }
}

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export async function loadTenantId(accountName: string): Promise<string | undefined> {
  const p = TENANT_PATH_PREFIX + accountName;
  if (fs.existsSync(p)) {
    try {
      const val = fs.readFileSync(p, UTF8);
      // Only return values that look like real Azure AD tenant IDs (UUIDs).
      // Stale test values like "faked_tenant_id" must be ignored.
      if (val && UUID_RE.test(val)) {
        return val;
      }
      return undefined;
    } catch {
      // ignore
    }
  }
  return undefined;
}

export async function clearCache(accountName: string): Promise<void> {
  ensureCacheDir();
  try {
    fs.writeFileSync(CACHE_PATH_PREFIX + accountName + ".json", "", UTF8);
  } catch {
    // ignore
  }
}

// ── Service principal credential storage ────────────────────────────────

export type AzureSPConfig = {
  clientId: string;
  secret: string;
  tenantId: string;
};

const spCrypto = new AccountCrypto("azure");

export async function saveAzureSP(
  clientId: string,
  secret: string,
  tenantId: string
): Promise<void> {
  ensureCacheDir();
  const data: AzureSPConfig = { clientId, secret, tenantId };
  fs.writeFileSync(AZURE_SP_PATH, await spCrypto.encrypt(JSON.stringify(data)), UTF8);
}

export async function clearAzureSP(): Promise<void> {
  ensureCacheDir();
  if (fs.existsSync(AZURE_SP_PATH)) {
    fs.unlinkSync(AZURE_SP_PATH);
  }
}

export async function loadAzureSP(): Promise<AzureSPConfig | undefined> {
  ensureCacheDir();
  if (fs.existsSync(AZURE_SP_PATH)) {
    try {
      const data = await spCrypto.decrypt(fs.readFileSync(AZURE_SP_PATH, UTF8));
      return JSON.parse(data);
    } catch {
      return undefined;
    }
  }
  return undefined;
}

export function checkAzureSPFile(): boolean {
  if (fs.existsSync(AZURE_SP_PATH)) {
    try {
      const data = fs.readFileSync(AZURE_SP_PATH, UTF8);
      const parsed = JSON.parse(data);
      return Object.keys(parsed).length > 0;
    } catch {
      return false;
    }
  }
  return false;
}
