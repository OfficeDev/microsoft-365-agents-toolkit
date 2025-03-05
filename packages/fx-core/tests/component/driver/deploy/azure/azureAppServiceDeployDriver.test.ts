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
import { expect, assert } from "chai";
import fs from "fs-extra";
import { AzureAppServiceDeployDriver } from "../../../../../src/component/driver/deploy/azure/azureAppServiceDeployDriver";
import { DeployConstant } from "../../../../../src/component/constant/deployConstant";
import {
  MockedAzureAccountProvider,
  MockTelemetryReporter,
  MockUserInteraction,
  MyTokenCredential,
} from "../../../../core/utils";
import * as os from "os";
import * as path from "path";
import * as uuid from "uuid";
import { IProgressHandler } from "@microsoft/teamsfx-api";

describe("Azure App Service Deploy Driver test", () => {
  const sandbox = sinon.createSandbox();
  const sysTmp = os.tmpdir();
  const folder = uuid.v4();
  const testFolder = path.join(sysTmp, folder);

  before(async () => {
    await fs.mkdirs(testFolder);
    await fs.writeFile(path.join(sysTmp, "ignore"), "ignore", {
      encoding: "utf8",
      flag: "a",
    });
    await fs.writeFile(path.join(testFolder, "test.txt"), "test");
    sandbox.stub(fs, "mkdirs").resolves();
    sandbox.stub(fs, "removeSync");
    sandbox.stub(fs, "readdirSync").returns([new fs.Dirent(), new fs.Dirent()]);
    sandbox.stub(fs, "readFileSync").resolves("test");
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

  after(async () => {
    if (fs.existsSync(testFolder)) {
      if (fs.existsSync(path.join(testFolder, ".deployment"))) {
        if (fs.existsSync(path.join(testFolder, ".deployment", "deployment.zip"))) {
          fs.unlinkSync(path.join(testFolder, ".deployment", "deployment.zip"));
        }
        fs.rmSync(path.join(testFolder, ".deployment"), { recursive: true });
      }
      fs.rmSync(testFolder, { recursive: true });
    }
  });

  beforeEach(async () => {
    sandbox.stub(tools, "waitSeconds").resolves();
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("deploy happy path", async () => {
    const deploy = new AzureAppServiceDeployDriver();
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Web/sites/some-server-farm",
      outputZipFile: ".deployment/deployment.zip",
    } as DeployArgs;
    const progressHandler: IProgressHandler = {
      start: async (): Promise<void> => {},
      next: async (): Promise<void> => {},
      end: async (): Promise<void> => {},
    };
    const ui = new MockUserInteraction();
    const progressNextCaller = sandbox.stub(progressHandler, "next").resolves();

    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
      ui: ui,
      telemetryReporter: new MockTelemetryReporter(),
      progressBar: progressHandler,
    } as any;
    const credential = new MyTokenCredential();
    sandbox.stub(credential, "getToken").resolves(undefined);
    sandbox.stub(context.azureAccountProvider, "getIdentityCredentialAsync").resolves(credential);
    // ignore file
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").callsFake((file) => {
      if (file.toString().indexOf("ignore") >= 0) {
        return Promise.resolve(Buffer.from("node_modules"));
      }
      return Promise.resolve(Buffer.from("any other content"));
    });
    sandbox.stub(fs, "mkdirs").resolves();
    sandbox.stub(fs, "removeSync");
    sandbox.stub(fs, "readdirSync").returns([new fs.Dirent(), new fs.Dirent()]);
    const client = new appService.WebSiteManagementClient(credential, "z");
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    // mock klaw
    // sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    // sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").resolves({
      status: 200,
      headers: {
        location: "/api/123",
      },
    });
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: 200,
      data: {
        status: 4,
        message: "success",
        received_time: 123,
        start_time: 111,
        end_time: 123,
        last_success_end_time: 100,
        complete: true,
        active: 1,
        is_readonly: true,
        site_name: "new_name",
      },
    });
    sandbox.stub(client.webApps, "restart").resolves();
    const rex = await deploy.execute(args, context);
    expect(rex.result.unwrapOr(new Map([["a", "a"]])).size).to.equal(0);
    // progress bar have 6 steps
    expect(progressNextCaller.callCount).to.equal(1);
  });

  it("deploy happy path with response data is empty", async () => {
    const deploy = new AzureAppServiceDeployDriver();
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Web/sites/some-server-farm",
      outputZipFile: ".deployment/deployment.zip",
    } as DeployArgs;
    const progressHandler: IProgressHandler = {
      start: async (detail?: string): Promise<void> => {},
      next: async (detail?: string): Promise<void> => {},
      end: async (): Promise<void> => {},
    };
    const ui = new MockUserInteraction();
    const progressNextCaller = sandbox.stub(progressHandler, "next").resolves();

    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
      ui: ui,
      telemetryReporter: new MockTelemetryReporter(),
      progressBar: progressHandler,
    } as any;
    const credential = new MyTokenCredential();
    sandbox.stub(credential, "getToken").resolves(undefined);
    sandbox.stub(context.azureAccountProvider, "getIdentityCredentialAsync").resolves(credential);
    // ignore file
    const client = new appService.WebSiteManagementClient(credential, "z");
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    // mock klaw
    // sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").resolves({
      status: 200,
      headers: {
        location: "/api/123",
      },
    });
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: 200,
      data: {},
    });
    sandbox.stub(client.webApps, "restart").resolves();
    const rex = await deploy.execute(args, context);
    expect(rex.result.unwrapOr(new Map([["a", "a"]])).size).to.equal(0);
    // progress bar have 6 steps
    expect(progressNextCaller.callCount).to.equal(1);
  });

  it("resource id error", async () => {
    const deploy = new AzureAppServiceDeployDriver();
    const args = {
      workingDirectory: "/",
      artifactFolder: "/",
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Web/sites",
    } as DeployArgs;
    const context = {
      logProvider: new TestLogProvider(),
      ui: new MockUserInteraction(),
    } as any;
    // await deploy.run(args, context);
    const res = await deploy.execute(args, context);
    assert.equal(res.result.isErr(), true);
  });

  it("missing resource id", async () => {
    const deploy = new AzureAppServiceDeployDriver();
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      ignoreFile: "./ignore",
    } as DeployArgs;
    const context = {
      logProvider: new TestLogProvider(),
    } as any;
    // await deploy.run(args, context);
    const res = await deploy.execute(args, context);
    assert.equal(res.result.isErr(), true);
  });

  it("deploy with ignore file not exists", async () => {
    const deploy = new AzureAppServiceDeployDriver();
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      ignoreFile: "",
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

    const client = new appService.WebSiteManagementClient(new MyTokenCredential(), "z");
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").resolves({
      status: 200,
      headers: {
        location: "/api/123",
      },
    });
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: 200,
    });
    sandbox.stub(client.webApps, "restart").resolves();
    // read deploy zip file error
    sandbox
      .stub(fs, "readFile")
      .withArgs(
        `./${DeployConstant.DEPLOYMENT_TMP_FOLDER}/${DeployConstant.DEPLOYMENT_ZIP_CACHE_FILE}`
      )
      .throws(new Error("test"));
    // mock klaw
    sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    const res = await deploy.execute(args, context);
    expect(res.result.unwrapOr(new Map([["a", "b"]])).size).to.equal(0);
  });

  it("zip deploy to azure error", async () => {
    const deploy = new AzureAppServiceDeployDriver();
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
    const client = new appService.WebSiteManagementClient(new MyTokenCredential(), "z");
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(fs, "readFileSync").resolves("test");
    // read deploy zip file error
    sandbox.stub(fs, "readFile").callsFake((file) => {
      if (file === "ignore") {
        return Promise.resolve(Buffer.from("node_modules"));
      }
      throw new Error("not found");
    });
    // mock klaw
    sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").throws(new Error("test"));
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: 200,
    });
    const res = await deploy.execute(args, context);
    assert.equal(res.result.isErr(), true);
  });

  it("should thrown when deploy remote 500 error", async () => {
    const deploy = new AzureAppServiceDeployDriver();
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
    const client = new appService.WebSiteManagementClient(new MyTokenCredential(), "z");
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(fs, "readFileSync").resolves("test");
    // read deploy zip file error
    sandbox.stub(fs, "readFile").callsFake((file) => {
      if (file === "ignore") {
        return Promise.resolve(Buffer.from("node_modules"));
      }
      throw new Error("not found");
    });
    // mock klaw
    sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").throws({
      response: {
        status: 503,
      },
      isAxiosError: true,
    });
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: 200,
    });
    const res = await deploy.execute(args, context);
    assert.equal(res.result.isErr(), true);
  });

  it("should thrown when deploy remote 400 error", async () => {
    const deploy = new AzureAppServiceDeployDriver();
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
    const client = new appService.WebSiteManagementClient(new MyTokenCredential(), "z");
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(fs, "readFileSync").resolves("test");
    // read deploy zip file error
    sandbox.stub(fs, "readFile").callsFake((file) => {
      if (file === "ignore") {
        return Promise.resolve(Buffer.from("node_modules"));
      }
      throw new Error("not found");
    });
    // mock klaw
    sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").throws({
      response: {
        status: 404,
      },
      isAxiosError: true,
    });
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: 200,
    });
    const res = await deploy.execute(args, context);
    assert.equal(res.result.isErr(), true);
  });

  it("working dir not exists", async () => {
    const deploy = new AzureAppServiceDeployDriver();
    const args = {
      workingDirectory: "/aaaa",
      artifactFolder: `./${folder}`,
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Web/sites/some-server-farm",
    } as DeployArgs;
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
    } as any;
    const res = await deploy.execute(args, context);
    assert.equal(res.result.isErr(), true);
  });

  it("test dry run", async () => {
    const deploy = new AzureAppServiceDeployDriver();
    const fh = await fs.open(path.join(sysTmp, folder, "test.txt"), "a");
    await fs.close(fh);
    await fs.writeFile(path.join(sysTmp, folder, "ignore"), "ignore", {
      encoding: "utf8",
      flag: "a",
    });
    const args = {
      workingDirectory: sysTmp,
      artifactFolder: `./${folder}`,
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Web/sites/some-server-farm",
      dryRun: true,
    } as DeployArgs;

    const progressHandler: IProgressHandler = {
      start: async (detail?: string): Promise<void> => {},
      next: async (detail?: string): Promise<void> => {},
      end: async (): Promise<void> => {},
    };
    const ui = new MockUserInteraction();
    const progressNextCaller = sandbox.stub(progressHandler, "next").resolves();

    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
      ui: ui,
      progressBar: progressHandler,
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
      return Promise.resolve(Buffer.from("any other content"));
    });
    const client = new appService.WebSiteManagementClient(new MyTokenCredential(), "z");
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox.stub(client.webApps, "beginListPublishingCredentialsAndWait").resolves({
      publishingUserName: "test-username",
      publishingPassword: "test-password",
    } as Models.WebAppsListPublishingCredentialsResponse);
    sandbox.stub(fs, "readFileSync").resolves("test");
    // mock klaw
    // sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
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
    sandbox.stub(client.webApps, "restart").resolves();
    const res = await deploy.execute(args, context);
    assert.equal(res.result.isOk(), true);
    const tmpFile = path.join(sysTmp, "./.deployment/deployment.zip");
    assert.equal(
      res.summaries[0],
      `Deployment preparations are completed. You can find the package in \`${tmpFile}\``
    );
    // dry run will have only one progress step
    assert.equal(progressNextCaller.callCount, 1);
  });

  it("list credential error", async () => {
    const deploy = new AzureAppServiceDeployDriver();
    const fh = await fs.open(path.join(sysTmp, folder, "test.txt"), "a");
    await fs.close(fh);
    await fs.writeFile(path.join(sysTmp, folder, "ignore"), "ignore", {
      encoding: "utf8",
      flag: "a",
    });
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
      ui: new MockUserInteraction(),
    } as any;
    const credential = new MyTokenCredential();
    credential.getToken = async () => {
      return null;
    };
    sandbox.stub(context.azureAccountProvider, "getIdentityCredentialAsync").resolves(credential);
    // ignore file
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").callsFake((file) => {
      if (file === "ignore") {
        return Promise.resolve(Buffer.from("node_modules"));
      }
      throw new Error("not found");
    });
    const client = new appService.WebSiteManagementClient(credential, "z");
    sandbox.stub(appService, "WebSiteManagementClient").returns(client);
    sandbox
      .stub(client.webApps, "beginListPublishingCredentialsAndWait")
      .throws(new Error("error"));
    sandbox.stub(fs, "readFileSync").resolves("test");
    // mock klaw
    // sandbox.stub(fileOpt, "forEachFileAndDir").resolves(undefined);
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
    sandbox.stub(client.webApps, "restart").resolves();
    const res = await deploy.execute(args, context);
    assert.equal(res.result.isOk(), false);
  });
});
