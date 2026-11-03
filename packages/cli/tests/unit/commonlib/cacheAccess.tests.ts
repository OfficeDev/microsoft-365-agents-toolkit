// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs, { WriteFileOptions } from "fs-extra";
import { AzureAccountManager } from "../../../src/commonlib/azureLogin";
import {
  AccountCrypto,
  AzureSpCrypto,
  clearCache,
  CryptoCachePlugin,
  loadAccountId,
  loadTenantId,
  saveTenantId,
} from "../../../src/commonlib/cacheAccess";
import VsCodeLogInstance from "../../../src/commonlib/log";
import { expect } from "../utils";
import { vi } from "vitest";
class MockKeytar {
  public async getPassword(service: string, account: string): Promise<string | null> {
    return Promise.resolve("1234567890abcdefghijklmnopqrstuv");
  }

  public async setPassword(service: string, account: string, password: string): Promise<void> {
    return Promise.resolve();
  }

  public async deletePassword(service: string, account: string): Promise<boolean> {
    throw new Error("Not Supported");
  }

  public async findPassword(service: string): Promise<string | null> {
    throw new Error("Not Supported");
  }

  public async findCredentials(
    service: string
  ): Promise<Array<{ account: string; password: string }>> {
    throw new Error("Not Supported");
  }
}

async function getRejectedError(promise: Promise<unknown>): Promise<Error> {
  try {
    await promise;
  } catch (error) {
    if (error instanceof Error) {
      return error;
    }
    throw error;
  }
  throw new Error("Expected promise to reject.");
}

describe("AccountCrypto Tests", function () {
  // nothing to do with the test logic, but to avoid weird type error. (https://stackoverflow.com/questions/68051262/)
  const azureAccountManager = AzureAccountManager.getInstance();
  const cachePlugin = new CryptoCachePlugin("test");

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("Encrypt/Decrypt Content", async () => {
    const accountCrypto = new AccountCrypto("test");
    (<any>accountCrypto).keytar = new MockKeytar();

    const content =
      '{"clientId":"clientId","secret":"secret","tenantId":"3c8f28dd-b990-4925-96a6-3ea9495654b8"}';
    const encrypted = await accountCrypto.encrypt(content);
    expect(encrypted.includes(content)).to.be.false;
    const decrtpyed = await accountCrypto.decrypt(encrypted);
    expect(decrtpyed).equals(content);
  });

  it("Encrypt/Decrypt Content - Unknown key", async () => {
    const accountCrypto = new AccountCrypto("test");
    (<any>accountCrypto).keytar = new MockKeytar();
    (<any>accountCrypto).keytar.getPassword = async () => {
      throw new Error("unknown key");
    };

    const content =
      '{"clientId":"clientId","secret":"secret","tenantId":"3c8f28dd-b990-4925-96a6-3ea9495654b8"}';
    const encryptError = await getRejectedError(accountCrypto.encrypt(content));
    const decryptError = await getRejectedError(accountCrypto.decrypt(content));

    expect(encryptError.message).to.equal("Secure credential storage is unavailable.");
    expect(decryptError.message).to.equal("Secure credential storage is unavailable.");
  });

  it("removes a plaintext MSAL cache when secure storage is unavailable", async () => {
    const accountCrypto = new AccountCrypto("test");
    (<any>accountCrypto).keytar = undefined;
    const plugin = new CryptoCachePlugin("test", accountCrypto);
    const plaintextCache =
      '{"Account":{},"AccessToken":{"token":{"secret":"fake-access-token"}}}';
    const removeStub = vi.spyOn(fs, "remove").mockResolvedValue();

    vi.spyOn(fs, "ensureDir").mockResolvedValue();
    vi.spyOn(fs, "pathExists").mockResolvedValue(true);
    vi.spyOn(fs, "readFile").mockResolvedValue(plaintextCache as any);

    await plugin.beforeCacheAccess({
      tokenCache: {
        deserialize: vi.fn(),
      },
    } as any);

    expect(removeStub).toHaveBeenCalledOnce();
  });

  it("migrates a plaintext MSAL cache when secure storage is available", async () => {
    const accountCrypto = new AccountCrypto("test");
    (<any>accountCrypto).keytar = new MockKeytar();
    const plugin = new CryptoCachePlugin("test", accountCrypto);
    const plaintextCache =
      '{"Account":{},"AccessToken":{"token":{"secret":"fake-access-token"}}}';
    const deserializeStub = vi.fn();
    const writeFileStub = vi.spyOn(fs, "writeFile").mockResolvedValue();

    vi.spyOn(fs, "ensureDir").mockResolvedValue();
    vi.spyOn(fs, "pathExists").mockResolvedValue(true);
    vi.spyOn(fs, "readFile").mockResolvedValue(plaintextCache as any);

    await plugin.beforeCacheAccess({
      tokenCache: {
        deserialize: deserializeStub,
      },
    } as any);

    expect(deserializeStub).toHaveBeenCalledWith(plaintextCache);
    expect(writeFileStub).toHaveBeenCalledOnce();
    expect(writeFileStub.mock.calls[0][1]).not.to.equal(plaintextCache);
  });
});

describe("AccountCrypto Service principal Tests", function () {
  const sandbox = vi;

  before(() => {
    vi.spyOn(fs, "ensureDir").mockImplementation(async (path: fs.PathLike) => {
      return true;
    });
    vi.spyOn(fs, "writeFile").mockImplementation(
      async (folder: any, content: string, options?: WriteFileOptions | string) => {
        return;
      }
    );
    vi.spyOn(fs, "remove").mockImplementation(async (path: fs.PathLike) => {
      return;
    });
    vi.spyOn(fs, "pathExists").mockImplementation(async (path: fs.PathLike) => {
      return true;
    });
    vi.spyOn(fs, "existsSync").mockImplementation((path: fs.PathLike) => {
      return true;
    });
    vi.spyOn(fs, "readFileSync").mockImplementation((path: any, options: any) => {
      return '{"clientId":"clientId","secret":"secret","tenantId":"3c8f28dd-b990-4925-96a6-3ea9495654b8"}';
    });
    vi.spyOn(fs, "readFile").mockImplementation(
      async (file: string | Buffer | number, options?: any) => {
        return Buffer.from(
          '{"i":"1f26ae86a392931c124a60bdd87bcfad","c":"855944b46b9250d67494aa072bafe4bafe8dee9712ae4dd21bc31d9fa75ba4048ac4845fd68905fc1dca28fd5df06bed6af2c7ecb57121b0ba560fd1a71ed241eafef0aa503c0a51722aa11b1e8482dfded052bb9e66630fb785b3","t":"b1b705aa70948661e365e9f3e95c6cae"}',
          "utf-8"
        );
      }
    );
  });

  after(() => {
    vi.restoreAllMocks();
  });

  it("AzureSpCrypto test", async () => {
    (<any>AzureSpCrypto).accountCrypto.keytar = new MockKeytar();
    await AzureSpCrypto.saveAzureSP("clientId", "secret", "tenantId");
    const checkAzureSp = AzureSpCrypto.checkAzureSPFile();
    expect(checkAzureSp).to.be.true;
    await AzureSpCrypto.loadAzureSP();
    await AzureSpCrypto.clearAzureSP();
    await loadAccountId("abc");
    await clearCache("abc");
  });

  it("does not write a service principal secret when secure storage is unavailable", async () => {
    (<any>AzureSpCrypto).accountCrypto.keytar = undefined;
    (<any>AzureSpCrypto).accountCrypto.currentKey = undefined;
    const writeFileStub = vi.mocked(fs.writeFile);
    writeFileStub.mockClear();

    await AzureSpCrypto.saveAzureSP("clientId", "secret", "tenantId");

    expect(writeFileStub).not.toHaveBeenCalled();
  });

  it("migrates plaintext service principal credentials", async () => {
    (<any>AzureSpCrypto).accountCrypto.keytar = new MockKeytar();
    (<any>AzureSpCrypto).accountCrypto.currentKey = undefined;
    const plaintext =
      '{"clientId":"clientId","secret":"secret","tenantId":"3c8f28dd-b990-4925-96a6-3ea9495654b8"}';
    const readFileStub = vi.mocked(fs.readFile);
    const writeFileStub = vi.mocked(fs.writeFile);
    readFileStub.mockResolvedValueOnce(Buffer.from(plaintext));
    writeFileStub.mockClear();

    const loaded = await AzureSpCrypto.loadAzureSP();

    expect(loaded).to.deep.equal({
      clientId: "clientId",
      secret: "secret",
      tenantId: "3c8f28dd-b990-4925-96a6-3ea9495654b8",
    });
    expect(writeFileStub).toHaveBeenCalledOnce();
    expect(writeFileStub.mock.calls[0][1]).not.to.equal(plaintext);
  });
});

describe("tenant id save/load", () => {
  let tenantId: string | undefined = undefined;
  const sandbox = vi;

  context("Files read/write successfully", () => {
    beforeEach(() => {
      vi.spyOn(fs, "ensureDir").mockResolvedValue();
      vi.spyOn(fs, "writeFile").mockImplementation((dir, id) => {
        tenantId = id;
      });
      vi.spyOn(fs, "readFile").mockImplementation(async () => {
        return Promise.resolve(tenantId! as any);
      });
    });

    afterEach(() => {
      vi.restoreAllMocks();
    });

    it("save and load tenant id", async () => {
      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      await saveTenantId("m365", "faked_tenant_id");
      const tid = await loadTenantId("m365");
      expect(tid).to.equal("faked_tenant_id");
    });

    it("save and load empty tenant id", async () => {
      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      await saveTenantId("m365", "");
      const tid = await loadTenantId("m365");
      expect(tid).to.equal("");
    });

    it("should return undefined tenant id when cache path not exist", async () => {
      vi.spyOn(fs, "pathExists").mockResolvedValue(false);
      const tid = await loadTenantId("m365");
      expect(tid).to.equal(undefined);
    });
  });

  context("Error throws", () => {
    let logStub: any;
    beforeEach(() => {
      vi.spyOn(fs, "ensureDir").mockResolvedValue();
      vi.spyOn(fs, "writeFile").mockImplementation(() => {
        throw new Error();
      });
      vi.spyOn(fs, "readFile").mockImplementation(() => {
        throw new Error();
      });
      vi.spyOn(fs, "pathExists").mockResolvedValue(true);
      logStub = vi.spyOn(VsCodeLogInstance, "warning");
    });

    afterEach(() => {
      vi.restoreAllMocks();
    });

    it("load tenant fail", async () => {
      const tid = await loadTenantId("m365");

      expect(tid).equal(undefined);
      expect(logStub.mock.calls.length === 1).to.true;
    });

    it("save tenant fail", async () => {
      await saveTenantId("m365", "faked_tenant_id");

      expect(logStub.mock.calls.length === 1).to.true;
    });
  });
});
