// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author xzf0587 <zhaofengxu@microsoft.com>
 */
import * as chai from "chai";
import * as sinon from "sinon";
import * as tools from "../../../../../src/common/utils";
import { AzureStorageStaticWebsiteConfigDriver } from "../../../../../src/component/driver/deploy/azure/azureStorageStaticWebsiteConfigDriver";
import { TestLogProvider } from "../../../util/logProviderMock";
import { DriverContext } from "../../../../../src/component/driver/interface/commonArgs";
import { ListAccountSasResponse, StorageManagementClient } from "@azure/arm-storage";
import { BlobServiceClient, ServiceGetPropertiesResponse } from "@azure/storage-blob";
import * as armStorage from "@azure/arm-storage";
import {
  MockedAzureAccountProvider,
  MockUserInteraction,
  MyTokenCredential,
} from "../../../../core/utils";
import { IProgressHandler } from "@microsoft/teamsfx-api";

describe("Azure Storage enable static website Driver test", () => {
  const sandbox = sinon.createSandbox();

  function getMockStorageAccount1() {
    return {
      // beginCreateAndWait: async function (
      //   resourceGroupName: string,
      //   accountName: string,
      //   parameters: StorageAccountCreateParameters,
      //   options?: StorageAccountsCreateOptionalParams
      // ): Promise<StorageAccountsCreateResponse> {
      //   return storageAccount!;
      // },
      listAccountSAS: async function (): Promise<ListAccountSasResponse> {
        return {
          accountSasToken: "abc",
        };
      },
    };
  }

  beforeEach(() => {
    sandbox.stub(tools, "waitSeconds").resolves();
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("Azure Storage enable static website happy path", async () => {
    const driver = new AzureStorageStaticWebsiteConfigDriver();
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui: new MockUserInteraction(),
      logProvider: new TestLogProvider(),
      progressBar: {
        start: async (detail?: string): Promise<void> => {},
        next: async (detail?: string): Promise<void> => {},
        end: async (): Promise<void> => {},
      } as IProgressHandler,
    } as any;
    // fake azure credentials
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());

    // fake sas account token
    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);

    // fake properties
    sandbox.stub(BlobServiceClient.prototype, "getProperties").resolves({
      staticWebsite: {
        enabled: false,
      },
    } as ServiceGetPropertiesResponse);

    const caller = sandbox.stub(BlobServiceClient.prototype, "setProperties").resolves();

    const res = await driver.run(
      {
        storageResourceId:
          "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
        indexPage: "index.html",
        errorPage: "error.html",
      },
      context
    );

    sinon.assert.calledOnce(caller);
    chai.assert.equal(res.isOk(), true);

    const rex = await driver.execute(
      {
        storageResourceId:
          "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
        indexPage: "index.html",
        errorPage: "error.html",
      },
      context
    );
    chai.assert.equal(rex.result.isOk(), true);
  });

  it("Azure Storage use default", async () => {
    const driver = new AzureStorageStaticWebsiteConfigDriver();
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui: new MockUserInteraction(),
      logProvider: new TestLogProvider(),
    } as any;
    // fake azure credentials
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());

    // fake sas account token
    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);

    // fake properties
    sandbox.stub(BlobServiceClient.prototype, "getProperties").resolves({
      staticWebsite: {
        enabled: false,
      },
    } as ServiceGetPropertiesResponse);

    const caller = sandbox.stub(BlobServiceClient.prototype, "setProperties").resolves();

    const res = await driver.run(
      {
        storageResourceId:
          "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
        errorPage: null,
      },
      context
    );

    sinon.assert.calledOnce(caller);
    chai.assert.equal(res.isOk(), true);
  });

  it("should skip enable static website", async () => {
    const driver = new AzureStorageStaticWebsiteConfigDriver();
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
    } as any;
    // fake azure credentials
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());

    // fake sas account token
    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);

    // fake properties
    sandbox.stub(BlobServiceClient.prototype, "getProperties").resolves({
      staticWebsite: {
        enabled: true,
      },
    } as ServiceGetPropertiesResponse);

    const caller = sandbox.stub(BlobServiceClient.prototype, "setProperties").resolves();

    await driver.run(
      {
        storageResourceId:
          "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
        indexPage: "index.html",
        errorPage: "error.html",
      },
      context
    );

    sinon.assert.notCalled(caller);
  });

  it("Azure Storage set properties error", async () => {
    const driver = new AzureStorageStaticWebsiteConfigDriver();
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui: new MockUserInteraction(),
      logProvider: new TestLogProvider(),
    } as any;
    // fake azure credentials
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());

    // fake sas account token
    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);

    // fake properties
    sandbox.stub(BlobServiceClient.prototype, "getProperties").resolves({
      staticWebsite: {
        enabled: false,
      },
    } as ServiceGetPropertiesResponse);

    const caller = sandbox
      .stub(BlobServiceClient.prototype, "setProperties")
      .throws({ statusCode: 404, message: "Not found" });

    const res = await driver.run(
      {
        storageResourceId:
          "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
        indexPage: "index.html",
        errorPage: "error.html",
      },
      context
    );

    sinon.assert.calledOnce(caller);
    chai.assert.equal(res.isErr(), true);
  });

  it("Azure Storage set properties remote server error", async () => {
    const driver = new AzureStorageStaticWebsiteConfigDriver();
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui: new MockUserInteraction(),
      logProvider: new TestLogProvider(),
    } as any;
    // fake azure credentials
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());

    // fake sas account token
    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);

    // fake properties
    sandbox.stub(BlobServiceClient.prototype, "getProperties").resolves({
      staticWebsite: {
        enabled: false,
      },
    } as ServiceGetPropertiesResponse);

    const caller = sandbox
      .stub(BlobServiceClient.prototype, "setProperties")
      .throws({ statusCode: 500 });

    const res = await driver.run(
      {
        storageResourceId:
          "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
        indexPage: "index.html",
        errorPage: "error.html",
      },
      context
    );

    sinon.assert.calledOnce(caller);
    chai.assert.equal(res.isErr(), true);
    console.log(res);
  });

  it("Azure Storage enable static website get properties error", async () => {
    const driver = new AzureStorageStaticWebsiteConfigDriver();
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui: new MockUserInteraction(),
      logProvider: new TestLogProvider(),
    } as any;
    // fake azure credentials
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());

    // fake sas account token
    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);

    // fake properties
    sandbox
      .stub(BlobServiceClient.prototype, "getProperties")
      .throws({ statusCode: 404, message: "Not found" });

    const res = await driver.run(
      {
        storageResourceId:
          "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
        indexPage: "index.html",
        errorPage: "error.html",
      },
      context
    );
    chai.assert.equal(res.isErr(), true);
    chai.assert.equal(res._unsafeUnwrapErr().name, "AzureStorageGetContainerPropertiesError");
  });

  it("Azure Storage enable static website get properties remote error", async () => {
    const driver = new AzureStorageStaticWebsiteConfigDriver();
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui: new MockUserInteraction(),
      logProvider: new TestLogProvider(),
    } as any;
    // fake azure credentials
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());

    // fake sas account token
    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);

    // fake properties
    sandbox.stub(BlobServiceClient.prototype, "getProperties").throws({ statusCode: 500 });

    const res = await driver.run(
      {
        storageResourceId:
          "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
        indexPage: "index.html",
        errorPage: "error.html",
      },
      context
    );
    chai.assert.equal(res.isErr(), true);
    chai.assert.equal(res._unsafeUnwrapErr().name, "AzureStorageGetContainerPropertiesError");
  });
});
