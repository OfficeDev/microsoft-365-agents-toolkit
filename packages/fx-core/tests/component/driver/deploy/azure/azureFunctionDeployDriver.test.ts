// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author xzf0587 <zhaofengxu@microsoft.com>
 */
import "mocha";
import * as sinon from "sinon";
import * as tools from "../../../../../src/common/utils";
import { DeployArgs } from "../../../../../src/component/driver/interface/buildAndDeployArgs";
import { TestLogProvider } from "../../../util/logProviderMock";
import * as appService from "@azure/arm-appservice";
import * as Models from "@azure/arm-appservice/src/models";
import * as fileOpt from "../../../../../src/component/utils/fileOperation";
import { AzureDeployImpl } from "../../../../../src/component/driver/deploy/azure/impl/azureDeployImpl";
import { assert, expect } from "chai";
import * as fs from "fs-extra";
import { AzureFunctionDeployDriver } from "../../../../../src/component/driver/deploy/azure/azureFunctionDeployDriver";
import {
  MockedAzureAccountProvider,
  MockTelemetryReporter,
  MockUserInteraction,
  MyTokenCredential,
} from "../../../../core/utils";
import * as os from "os";
import * as uuid from "uuid";
import * as path from "path";
import { AxiosError } from "axios";

describe("Azure Function Deploy Driver test", () => {
  const sandbox = sinon.createSandbox();
  const sysTmp = os.tmpdir();
  const folder = uuid.v4();
  const testFolder = path.join(sysTmp, folder);

  before(async () => {
    await fs.mkdirs(testFolder);
    await fs.writeFile(path.join(testFolder, "test.txt"), "test");
    sandbox
      .stub(fs, "createReadStream")
      .withArgs(path.join(sysTmp, ".deployment/deployment.zip"))
      .returns({
        pipe: sandbox.stub().returns({
          pipe: sandbox.stub().returns({
            pipe: sandbox.stub().returns("responseMock"),
          }),
        }),
        on: sandbox.spy(() => true),
        destroy: sandbox.spy(() => true),
      } as any);
    sandbox.stub(fs, "existsSync").returns(true);
    sandbox.stub(fs, "remove").resolves();
  });

  after(async () => {
    fs.rmSync(testFolder, { recursive: true, force: true });
  });

  beforeEach(() => {
    sandbox.stub(tools, "waitSeconds").resolves();
    const fetchStub = sandbox.stub(global, "fetch");
    fetchStub.callsFake((input: any) => {
      const url: string = typeof input === "string" ? input : input.url;
      const info = url.split(/[\/|?]/);
      return Promise.resolve(
        new Response(
          JSON.stringify({
            properties: {
              enabledHostNames: [
                `${info[info.length - 2]}.azurewebsites.net`,
                `${info[info.length - 2]}.scm.azurewebsites.net`,
              ],
            },
          }),
          { status: 200, headers: { "Content-Type": "application/json" } }
        )
      );
    });
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("deploy happy path", async () => {
    const deploy = new AzureFunctionDeployDriver();
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Web/sites/some-server-farm",
      zipFilePath: path.join(testFolder, "test.zip"),
    } as DeployArgs;
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui: new MockUserInteraction(),
      logProvider: new TestLogProvider(),
      telemetryReporter: new MockTelemetryReporter(),
    } as any;
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    // ignore file
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").callsFake((file) => {
      if (file === "ignore") {
        return Promise.resolve(Buffer.from("node_modules"));
      }
      return Promise.resolve(Buffer.from("other file"));
    });
    const client = new appService.WebSiteManagementClient(new MyTokenCredential(), "z");
    sandbox.stub(client.webApps, "restart").resolves();
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    sandbox.stub(fs, "readFileSync").resolves("test");
    // mock klaw
    sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").resolves({
      status: 200,
      headers: {
        location: "/api/123",
      },
    });
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: 200,
    });
    const rex = await deploy.execute(args, context);
    expect(rex.result.unwrapOr(new Map([["a", "b"]])).size).to.equal(0);
  });

  it("deploy restart error!", async () => {
    const deploy = new AzureFunctionDeployDriver();
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Web/sites/some-server-farm",
    } as DeployArgs;
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
      telemetryReporter: new MockTelemetryReporter(),
    } as any;
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    // ignore file
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").callsFake((file) => {
      if (file === "ignore") {
        return Promise.resolve(Buffer.from("node_modules"));
      }
      return Promise.resolve(Buffer.from("other file"));
    });
    const client = new appService.WebSiteManagementClient(new MyTokenCredential(), "z");
    sandbox.stub(client.webApps, "restart").rejects();
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    sandbox.stub(fs, "readFileSync").resolves("test");
    // mock klaw
    sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").resolves({
      status: 200,
      headers: {
        location: "/api/123",
      },
    });
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: 200,
    });
    const res = await deploy.execute(args, context);
    expect(res.result.isErr()).to.equal(false);
  });

  it("deploy restart throws", async () => {
    const deploy = new AzureFunctionDeployDriver();
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Web/sites/some-server-farm",
    } as DeployArgs;
    const logger = new TestLogProvider();
    const caller = sandbox.stub(logger, "warning").resolves();
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: logger,
      telemetryReporter: new MockTelemetryReporter(),
    } as any;
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    // ignore file
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").callsFake((file) => {
      if (file === "ignore") {
        return Promise.resolve(Buffer.from("node_modules"));
      }
      return Promise.resolve(Buffer.from("other file"));
    });
    const client = new appService.WebSiteManagementClient(new MyTokenCredential(), "z");
    sandbox.stub(client.webApps, "restart").throws(new Error("test"));
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    sandbox.stub(fs, "readFileSync").resolves("test");
    // mock klaw
    sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").resolves({
      status: 200,
      headers: {
        location: "/api/123",
      },
    });
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: 200,
    });
    const res = await deploy.execute(args, context);
    expect(res.result.isErr()).to.equal(false);
    // log warning will print
    sinon.assert.calledOnce(caller);
  });

  it("Zip deploy throws when upload", async () => {
    const deploy = new AzureFunctionDeployDriver();
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Web/sites/some-server-farm",
    } as DeployArgs;
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
    } as any;
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    // ignore file
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").callsFake((file) => {
      if (file === "ignore") {
        return Promise.resolve(Buffer.from("node_modules"));
      }
      throw new Error("not found");
    });
    const client = new appService.WebSiteManagementClient(new MyTokenCredential(), "z");
    sandbox.stub(client.webApps, "restart").throws(new Error("test"));
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    sandbox.stub(fs, "readFileSync").resolves("test");
    // mock klaw
    sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").resolves({
      status: 403,
      headers: {
        location: "/api/123",
      },
    });
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: 200,
    });
    const res = await deploy.execute(args, context);
    expect(res.result.isErr()).to.equal(true);
  });

  it("Check deploy status error", async () => {
    const deploy = new AzureFunctionDeployDriver();
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Web/sites/some-server-farm",
    } as DeployArgs;
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
    } as any;
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    // ignore file
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").callsFake((file) => {
      if (file === "ignore") {
        return Promise.resolve(Buffer.from("node_modules"));
      }
      throw new Error("not found");
    });
    const client = new appService.WebSiteManagementClient(new MyTokenCredential(), "z");
    sandbox.stub(client.webApps, "restart").throws(new Error("test"));
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    sandbox.stub(fs, "readFileSync").resolves("test");
    // mock klaw
    sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").resolves({
      status: 200,
      headers: {
        location: "/api/123",
      },
    });
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: 403,
    });
    const res = await deploy.execute(args, context);
    expect(res.result.isErr()).to.equal(true);
  });

  it("Check deploy status ok but cannot start", async () => {
    const deploy = new AzureFunctionDeployDriver();
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Web/sites/some-server-farm",
    } as DeployArgs;
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
      telemetryReporter: new MockTelemetryReporter(),
    } as any;
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    // ignore file
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").callsFake((file) => {
      if (file === "ignore") {
        return Promise.resolve(Buffer.from("node_modules"));
      }
      return Promise.resolve(Buffer.from("other file"));
    });
    const client = new appService.WebSiteManagementClient(new MyTokenCredential(), "z");
    sandbox.stub(client.webApps, "restart").throws(new Error("test"));
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    sandbox.stub(fs, "readFileSync").resolves("test");
    // mock klaw
    sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").resolves({
      status: 200,
      headers: {
        location: "/api/123",
      },
    });
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: 200,
      data: { status: 3 },
    });
    const res = await deploy.execute(args, context);
    expect(res.result.isOk()).to.equal(true);
  });

  it("Check deploy throws", async () => {
    const deploy = new AzureFunctionDeployDriver();
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Web/sites/some-server-farm",
    } as DeployArgs;
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
    } as any;
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    // ignore file
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").callsFake((file) => {
      if (file === "ignore") {
        return Promise.resolve(Buffer.from("node_modules"));
      }
      throw new Error("not found");
    });
    const client = new appService.WebSiteManagementClient(new MyTokenCredential(), "z");
    sandbox.stub(client.webApps, "restart").throws(new Error("test"));
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    sandbox.stub(fs, "readFileSync").resolves("test");
    // mock klaw
    sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").resolves({
      status: 200,
      headers: {
        location: "/api/123",
      },
    });
    sandbox
      .stub(AzureDeployImpl.AXIOS_INSTANCE, "get")
      .throws({ isAxiosError: true } as AxiosError);

    const res = await deploy.execute(args, context);
    expect(res.result.isErr()).to.equal(true);
  });

  it("deploy dry run", async () => {
    const deploy = new AzureFunctionDeployDriver();
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Web/sites/some-server-farm",
      dryRun: true,
    } as DeployArgs;
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      ui: new MockUserInteraction(),
      logProvider: new TestLogProvider(),
    } as any;
    sandbox
      .stub(context.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(new MyTokenCredential());
    // ignore file
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").callsFake((file) => {
      if (file === "ignore") {
        return Promise.resolve(Buffer.from("node_modules"));
      }
      return Promise.resolve(Buffer.from("other file"));
    });
    const client = new appService.WebSiteManagementClient(new MyTokenCredential(), "z");
    sandbox.stub(client.webApps, "restart").resolves();
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    sandbox.stub(fs, "readFileSync").resolves("test");
    // mock klaw
    sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").resolves({
      status: 200,
      headers: {
        location: "/api/123",
      },
    });
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: 200,
    });
    const res = await deploy.execute(args, context);

    assert.equal(res.result.isOk(), true);
    const tmpFile = path.join(sysTmp, "./.deployment/deployment.zip");
    assert.equal(
      res.summaries[0],
      `Deployment preparations are completed. You can find the package in \`${tmpFile}\``
    );
  });
});
