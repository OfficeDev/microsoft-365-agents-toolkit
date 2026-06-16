// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author xzf0587 <zhaofengxu@microsoft.com>
 */
import * as sinon from "sinon";
import * as tools from "../../../../../src/common/utils";
import { AzureStorageDeployDriver } from "../../../../../src/component/driver/deploy/azure/azureStorageDeployDriver";
import { DeployArgs } from "../../../../../src/component/driver/interface/buildAndDeployArgs";
import { TestLogProvider } from "../../../util/logProviderMock";
import { assert } from "chai";
import {
  ListAccountSasResponse,
  StorageAccounts,
  StorageManagementClient,
} from "@azure/arm-storage";
import {
  BlobDeleteResponse,
  BlockBlobClient,
  BlockBlobParallelUploadOptions,
  ContainerClient,
} from "@azure/storage-blob";
import * as armStorage from "@azure/arm-storage";
import {
  MockedAzureAccountProvider,
  MockUserInteraction,
  MyTokenCredential,
} from "../../../../core/utils";
import * as os from "os";
import * as uuid from "uuid";
import * as path from "path";
import * as fs from "fs-extra";
import * as chai from "chai";
import { IProgressHandler } from "@microsoft/teamsfx-api";

function getMockStorageAccount1() {
  return {
    listAccountSAS: async function (): Promise<ListAccountSasResponse> {
      return {
        accountSasToken: "abc",
      };
    },
  };
}

describe("Azure Storage Deploy Driver test", () => {
  const sandbox = sinon.createSandbox();
  const sysTmp = os.tmpdir();
  const folder = uuid.v4();
  const testFolder = path.join(sysTmp, folder);

  before(async () => {
    await fs.mkdirs(testFolder);
  });

  after(async () => {
    fs.rmSync(testFolder, { recursive: true, force: true });
  });

  beforeEach(() => {
    sandbox.stub(tools, "waitSeconds").resolves();
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("deploy to storage happy path", async () => {
    const deploy = new AzureStorageDeployDriver();
    await fs.open(path.join(testFolder, "test.txt"), "a");
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
    } as DeployArgs;
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
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    const clientStub = sandbox.createStubInstance(StorageManagementClient);
    clientStub.storageAccounts = {} as StorageAccounts;
    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);
    sandbox.stub(ContainerClient.prototype, "exists").resolves(false);
    sandbox.stub(ContainerClient.prototype, "create").resolves();
    sandbox.stub(ContainerClient.prototype, "listBlobsFlat").returns([
      {
        properties: {
          contentLength: 1,
        },
      },
    ] as any);
    //sandbox.stub(ContainerClient.prototype, "listBlobsFlat").resolves();
    sandbox
      .stub(ContainerClient.prototype, "deleteBlob")
      .resolves({ errorCode: undefined } as BlobDeleteResponse);
    /*const calls = sandbox.stub().callsFake(() => clientStub);
    Object.setPrototypeOf(StorageManagementClient, calls);*/
    sandbox.stub(ContainerClient.prototype, "getBlockBlobClient").returns({
      uploadFile: async (filePath: string, options?: BlockBlobParallelUploadOptions) => {
        return {};
      },
    } as BlockBlobClient);
    const rex = await deploy.execute(args, context);
    assert.equal(rex.result.isOk(), true);
  });

  it("deploy to storage happy path when storage exists", async () => {
    const deploy = new AzureStorageDeployDriver();
    await fs.open(path.join(testFolder, "test.txt"), "a");
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
    } as DeployArgs;
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
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    const clientStub = sandbox.createStubInstance(StorageManagementClient);
    clientStub.storageAccounts = {} as StorageAccounts;
    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);
    sandbox.stub(ContainerClient.prototype, "exists").resolves(true);
    sandbox.stub(ContainerClient.prototype, "listBlobsFlat").returns([
      {
        properties: {
          contentLengthNo: 1,
        },
      },
    ] as any);
    //sandbox.stub(ContainerClient.prototype, "listBlobsFlat").resolves();
    sandbox
      .stub(ContainerClient.prototype, "deleteBlob")
      .resolves({ errorCode: undefined } as BlobDeleteResponse);
    /*const calls = sandbox.stub().callsFake(() => clientStub);
    Object.setPrototypeOf(StorageManagementClient, calls);*/
    sandbox.stub(ContainerClient.prototype, "getBlockBlobClient").returns({
      uploadFile: async (filePath: string, options?: BlockBlobParallelUploadOptions) => {
        return {};
      },
    } as BlockBlobClient);
    const rex = await deploy.execute(args, context);
    assert.equal(rex.result.isOk(), true);
  });

  it("get azure account credential error", async () => {
    const deploy = new AzureStorageDeployDriver();
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
    } as DeployArgs;
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
    } as any;
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .throws(new Error("error"));

    const res = await deploy.execute(args, context);
    assert.equal(res.result.isErr(), true);
  });

  it("clear storage error", async () => {
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
    } as any;
    const deploy = new AzureStorageDeployDriver();
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    const clientStub = sandbox.createStubInstance(StorageManagementClient);
    clientStub.storageAccounts = {} as StorageAccounts;

    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
    } as DeployArgs;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);
    sandbox.stub(ContainerClient.prototype, "exists").resolves(false);
    sandbox.stub(ContainerClient.prototype, "create").resolves();
    sandbox.stub(ContainerClient.prototype, "listBlobsFlat").returns([
      {
        properties: {
          contentLength: 1,
        },
      },
    ] as any);
    //sandbox.stub(ContainerClient.prototype, "listBlobsFlat").resolves();
    sandbox
      .stub(ContainerClient.prototype, "deleteBlob")
      .resolves({ errorCode: "403" } as BlobDeleteResponse);
    const res = await deploy.execute(args, context);
    assert.equal(res.result.isErr(), true);
    chai.assert.equal(res.result._unsafeUnwrapErr().name, "AzureStorageClearBlobsError");
  });

  it("clear storage with remote server error", async () => {
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
    } as any;
    const deploy = new AzureStorageDeployDriver();
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    const clientStub = sandbox.createStubInstance(StorageManagementClient);
    clientStub.storageAccounts = {} as StorageAccounts;

    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
    } as DeployArgs;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);
    sandbox.stub(ContainerClient.prototype, "exists").resolves(false);
    sandbox.stub(ContainerClient.prototype, "create").resolves();
    sandbox.stub(ContainerClient.prototype, "listBlobsFlat").returns([
      {
        properties: {
          contentLength: 1,
        },
      },
    ] as any);
    //sandbox.stub(ContainerClient.prototype, "listBlobsFlat").resolves();
    sandbox
      .stub(ContainerClient.prototype, "deleteBlob")
      .resolves({ errorCode: "error", _response: { status: 500 } } as BlobDeleteResponse);
    const res = await deploy.execute(args, context);
    assert.equal(res.result.isErr(), true);
    chai.assert.equal(res.result._unsafeUnwrapErr().name, "AzureStorageClearBlobsError");
  });

  it("upload with error", async () => {
    const deploy = new AzureStorageDeployDriver();
    await fs.open(path.join(testFolder, "test.txt"), "a");
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
    } as DeployArgs;
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui: new MockUserInteraction(),
      logProvider: new TestLogProvider(),
    } as any;
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    const clientStub = sandbox.createStubInstance(StorageManagementClient);
    clientStub.storageAccounts = {} as StorageAccounts;
    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);
    sandbox.stub(ContainerClient.prototype, "exists").resolves(false);
    sandbox.stub(ContainerClient.prototype, "create").resolves();
    sandbox.stub(ContainerClient.prototype, "listBlobsFlat").returns([
      {
        properties: {
          contentLength: 1,
        },
      },
    ] as any);
    //sandbox.stub(ContainerClient.prototype, "listBlobsFlat").resolves();
    sandbox
      .stub(ContainerClient.prototype, "deleteBlob")
      .resolves({ errorCode: undefined } as BlobDeleteResponse);
    /*const calls = sandbox.stub().callsFake(() => clientStub);
    Object.setPrototypeOf(StorageManagementClient, calls);*/
    sandbox.stub(ContainerClient.prototype, "getBlockBlobClient").returns({
      uploadFile: async (filePath: string, options?: BlockBlobParallelUploadOptions) => {
        return { errorCode: "error" };
      },
    } as BlockBlobClient);
    const res = await deploy.execute(args, context);
    assert.equal(res.result.isErr(), true);
    chai.assert.equal(res.result._unsafeUnwrapErr().name, "AzureStorageUploadFilesError");
  });

  it("upload with remote server error", async () => {
    const deploy = new AzureStorageDeployDriver();
    await fs.open(path.join(testFolder, "test.txt"), "a");
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
    } as DeployArgs;
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui: new MockUserInteraction(),
      logProvider: new TestLogProvider(),
    } as any;
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    const clientStub = sandbox.createStubInstance(StorageManagementClient);
    clientStub.storageAccounts = {} as StorageAccounts;
    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);
    sandbox.stub(ContainerClient.prototype, "exists").resolves(false);
    sandbox.stub(ContainerClient.prototype, "create").resolves();
    sandbox.stub(ContainerClient.prototype, "listBlobsFlat").returns([
      {
        properties: {
          contentLength: 1,
        },
      },
    ] as any);
    //sandbox.stub(ContainerClient.prototype, "listBlobsFlat").resolves();
    sandbox
      .stub(ContainerClient.prototype, "deleteBlob")
      .resolves({ errorCode: undefined } as BlobDeleteResponse);
    /*const calls = sandbox.stub().callsFake(() => clientStub);
    Object.setPrototypeOf(StorageManagementClient, calls);*/
    sandbox.stub(ContainerClient.prototype, "getBlockBlobClient").returns({
      uploadFile: async (filePath: string, options?: BlockBlobParallelUploadOptions) => {
        return { errorCode: "error", _response: { status: 500 } };
      },
    } as BlockBlobClient);
    const res = await deploy.execute(args, context);
    assert.equal(res.result.isErr(), true);
    chai.assert.equal(res.result._unsafeUnwrapErr().name, "AzureStorageUploadFilesError");
  });

  it("get container with remote server error", async () => {
    const deploy = new AzureStorageDeployDriver();
    await fs.open(path.join(testFolder, "test.txt"), "a");
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
    } as DeployArgs;
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui: new MockUserInteraction(),
      logProvider: new TestLogProvider(),
    } as any;
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    const clientStub = sandbox.createStubInstance(StorageManagementClient);
    clientStub.storageAccounts = {} as StorageAccounts;
    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);
    sandbox.stub(ContainerClient.prototype, "exists").throws({ statusCode: 500 });
    sandbox.stub(ContainerClient.prototype, "getBlockBlobClient").returns({
      uploadFile: async (filePath: string, options?: BlockBlobParallelUploadOptions) => {
        return {};
      },
    } as BlockBlobClient);
    const res = await deploy.execute(args, context);
    assert.equal(res.result.isErr(), true);
    chai.assert.equal(res.result._unsafeUnwrapErr().name, "AzureStorageGetContainerError");
  });

  it("get container with normal error", async () => {
    const deploy = new AzureStorageDeployDriver();
    await fs.open(path.join(testFolder, "test.txt"), "a");
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Storage/storageAccounts/some-server-farm",
    } as DeployArgs;
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui: new MockUserInteraction(),
      logProvider: new TestLogProvider(),
    } as any;
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    const clientStub = sandbox.createStubInstance(StorageManagementClient);
    clientStub.storageAccounts = {} as StorageAccounts;
    const mockStorageManagementClient = new StorageManagementClient(new MyTokenCredential(), "id");
    mockStorageManagementClient.storageAccounts = getMockStorageAccount1() as any;
    sandbox.stub(armStorage, "StorageManagementClient").returns(mockStorageManagementClient);
    sandbox.stub(ContainerClient.prototype, "exists").throws({ statusCode: 400 });
    sandbox.stub(ContainerClient.prototype, "getBlockBlobClient").returns({
      uploadFile: async (filePath: string, options?: BlockBlobParallelUploadOptions) => {
        return {};
      },
    } as BlockBlobClient);
    const res = await deploy.execute(args, context);
    assert.equal(res.result.isErr(), true);
    chai.assert.equal(res.result._unsafeUnwrapErr().name, "AzureStorageGetContainerError");
  });
});
