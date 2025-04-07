// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import sinon from "sinon";
import AzureLoginCI from "../../../src/commonlib/azureLoginCI";
import { expect } from "../utils";
import fs, { WriteFileOptions } from "fs-extra";
import { signedOut } from "@microsoft/teamsfx-api";
import { AzureSPConfig, AzureSpCrypto } from "../../../src/commonlib/cacheAccess";

describe("Azure Service Principal login Tests", function () {
  const sandbox = sinon.createSandbox();

  beforeEach(async () => {
    sandbox.stub(fs, "ensureDir").callsFake(async (path: fs.PathLike) => {
      return true;
    });
    sandbox
      .stub(fs, "writeFile")
      .callsFake(async (folder: any, content: string, options?: WriteFileOptions | string) => {
        return;
      });
    sandbox.stub(fs, "remove").callsFake(async (path: fs.PathLike) => {
      return;
    });
    sandbox.stub(fs, "pathExists").callsFake(async (path: fs.PathLike) => {
      return true;
    });
    sandbox.stub(fs, "existsSync").callsFake((path: fs.PathLike) => {
      return true;
    });
    sandbox.stub(fs, "readFileSync").callsFake((path: any, options: any) => {
      return '{"clientId":"clientId","secret":"secret","tenantId":"3c8f28dd-b990-4925-96a6-3ea9495654b8"}';
    });
    sandbox
      .stub(fs, "readFile")
      .callsFake(async (file: string | Buffer | number, options?: any) => {
        return Buffer.from(
          '{"i":"1f26ae86a392931c124a60bdd87bcfad","c":"855944b46b9250d67494aa072bafe4bafe8dee9712ae4dd21bc31d9fa75ba4048ac4845fd68905fc1dca28fd5df06bed6af2c7ecb57121b0ba560fd1a71ed241eafef0aa503c0a51722aa11b1e8482dfded052bb9e66630fb785b3","t":"b1b705aa70948661e365e9f3e95c6cae"}',
          "utf-8"
        );
      });
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("init", async () => {
    sandbox.stub(AzureSpCrypto, "loadAzureSP").resolves(undefined);
    sandbox.stub(AzureSpCrypto, "saveAzureSP").resolves();
    sandbox.stub(AzureSpCrypto, "clearAzureSP").resolves();
    await AzureLoginCI.init("clientId", "secret", "tenantId");

    await AzureLoginCI.init("clientId", "~/3.pem", "tenantId");

    await AzureLoginCI.init("clientId", "D:/test/3.pem", "tenantId");
  });

  it("getIdentityCredentialAsync", async () => {
    sandbox.stub(AzureSpCrypto, "loadAzureSP").resolves(undefined);
    sandbox.stub(AzureSpCrypto, "saveAzureSP").resolves();
    sandbox.stub(AzureSpCrypto, "clearAzureSP").resolves();
    await AzureLoginCI.init("clientId", "secret", "tenantId");
    await AzureLoginCI.getIdentityCredentialAsync();
  });

  it("signout", async () => {
    sandbox.stub(AzureSpCrypto, "loadAzureSP").resolves({} as AzureSPConfig);
    sandbox.stub(AzureSpCrypto, "saveAzureSP").resolves();
    sandbox.stub(AzureSpCrypto, "clearAzureSP").resolves();
    const result = await AzureLoginCI.signout();
    expect(result).equals(true);

    const r1 = await AzureLoginCI.getStatus();
    expect(r1.status).equals(signedOut);
  });
});
