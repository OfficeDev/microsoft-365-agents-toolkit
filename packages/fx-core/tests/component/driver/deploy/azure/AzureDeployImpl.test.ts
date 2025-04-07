// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Siglud <siglud@gmail.com>
 */

import "mocha";
import {
  AzureUploadConfig,
  DeployArgs,
} from "../../../../../src/component/driver/interface/buildAndDeployArgs";
import { TestLogProvider } from "../../../util/logProviderMock";
import {
  MockedAzureAccountProvider,
  MockTelemetryReporter,
  MockUserInteraction,
  MyTokenCredential,
} from "../../../../core/utils";
import { AzureZipDeployImpl } from "../../../../../src/component/driver/deploy/azure/impl/AzureZipDeployImpl";
import * as tools from "../../../../../src/common/utils";
import * as sinon from "sinon";
import { AzureDeployImpl } from "../../../../../src/component/driver/deploy/azure/impl/azureDeployImpl";
import {
  CheckDeploymentStatusError,
  CheckDeploymentStatusTimeoutError,
  DeployZipPackageError,
  GetPublishingCredentialsError,
} from "../../../../../src/error/deploy";
import * as chai from "chai";
import chaiAsPromised = require("chai-as-promised");
chai.use(chaiAsPromised);
import * as appService from "@azure/arm-appservice";
import { RestError } from "@azure/storage-blob";
import {
  WebAppsListPublishingCredentialsResponse,
  WebSiteManagementClient,
} from "@azure/arm-appservice";
import { HttpStatusCode } from "../../../../../src/component/constant/commonConstant";
import { DeployStatus } from "../../../../../src/component/constant/deployConstant";
import * as fs from "fs-extra";
import * as os from "os";
import * as path from "path";
import { AzureResourceInfo } from "../../../../../src/component/driver/interface/commonArgs";

describe("AzureDeployImpl zip deploy acceleration", () => {
  const sandbox = sinon.createSandbox();
  const tempFile = path.join(os.tmpdir(), "test.zip");

  before(async () => {
    fs.writeFileSync(tempFile, "test");
  });

  after(async () => {
    fs.rmSync(tempFile, { recursive: true, force: true });
  });

  beforeEach(async () => {
    sandbox.stub(tools, "waitSeconds").resolves();
  });

  afterEach(async () => {
    sandbox.restore();
  });

  it("zip deploy need acceleration", async () => {
    const args = {
      workingDirectory: "./",
      artifactFolder: `./tmp`,
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-aa11aaaa1aa1/resourceGroups/hoho-rg/providers/Microsoft.Web/sites/some-server-farm",
    } as DeployArgs;
    const context = {
      azureAccountProvider: new MockedAzureAccountProvider(),
      logProvider: new TestLogProvider(),
      ui: new MockUserInteraction(),
    } as any;
    context.logProvider.info = async (msg: string | Array<any>) => {
      console.log(msg);
      return Promise.resolve(true);
    };
    const deploy = new AzureZipDeployImpl(args, context, "", "", [], []);
    sandbox.stub(deploy, "zipDeploy").resolves(5_000_000);
    await deploy.run();
  });

  it("Get zip deploy endpoint", async () => {
    const ar = {
      subscriptionId: "aaa",
      resourceGroupName: "bbb",
      instanceId: "ccc",
    } as AzureResourceInfo;
    const config = {
      headers: {
        "Content-Type": "AAA",
        "Cache-Control": "bbb",
        Authorization: "ccc",
      },
      maxContentLength: 1,
      maxBodyLength: 2,
      timeout: 3,
    } as AzureUploadConfig;
    const fetchStub = sandbox.stub(global, "fetch");
    fetchStub.resolves(
      new Response(
        JSON.stringify({
          properties: {
            enabledHostNames: [
              "ssssxx-h0gjdtbsa8bqhjhe.canadacentral-01.azurewebsites.net",
              "ssssxx-h0gjdtbsa8bqhjhe.scm.canadacentral-01.azurewebsites.net",
            ],
          },
        }),
        { status: 200 }
      )
    );
    const endpoint = await AzureZipDeployImpl.getZipDeployEndpoint(ar, config);
    chai.assert.equal(
      endpoint,
      "https://ssssxx-h0gjdtbsa8bqhjhe.scm.canadacentral-01.azurewebsites.net/api/zipdeploy?isAsync=true"
    );
    chai.expect(fetchStub.calledOnce).to.be.true;
    chai
      .expect(fetchStub.firstCall.args[0])
      .to.be.equal(
        "https://management.azure.com/subscriptions/aaa/resourceGroups/bbb/providers/Microsoft.Web/sites/ccc?api-version=2024-04-01"
      );
  });

  it("Get zip deploy endpoint with error response", async () => {
    const ar = {
      subscriptionId: "aaa",
      resourceGroupName: "bbb",
      instanceId: "ccc",
    } as AzureResourceInfo;
    const config = {
      headers: {
        "Content-Type": "AAA",
        "Cache-Control": "bbb",
        Authorization: "ccc",
      },
      maxContentLength: 1,
      maxBodyLength: 2,
      timeout: 3,
    } as AzureUploadConfig;
    const fetchStub = sandbox.stub(global, "fetch");
    fetchStub.resolves(
      new Response(
        JSON.stringify({
          properties: {
            enabledHostNames: [
              "ssssxx-h0gjdtbsa8bqhjhe.canadacentral-01.azurewebsites.net",
              "ssssxx-h0gjdtbsa8bqhjhe.canadacentral-01.azurewebsites.net",
            ],
          },
        }),
        { status: 200 }
      )
    );
    chai.expect(AzureZipDeployImpl.getZipDeployEndpoint(ar, config)).to.be.rejectedWith(Error);
    chai.expect(fetchStub.calledOnce).to.be.true;
    chai
      .expect(fetchStub.firstCall.args[0])
      .to.be.equal(
        "https://management.azure.com/subscriptions/aaa/resourceGroups/bbb/providers/Microsoft.Web/sites/ccc?api-version=2024-04-01"
      );
  });

  it("checkDeployStatus empty response", async () => {
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves("");
    const config = {
      headers: {
        "Content-Type": "text",
        "Cache-Control": "no-cache",
        Authorization: "no",
      },
      maxContentLength: 200,
      maxBodyLength: 200,
      timeout: 200,
    };
    const args = {
      workingDirectory: "/",
      artifactFolder: "/",
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-11111111111/resourceGroups/hoho-rg/providers/Microsoft.Web/sites",
    } as DeployArgs;
    const context = {
      logProvider: new TestLogProvider(),
      ui: new MockUserInteraction(),
    } as any;
    const impl = new AzureZipDeployImpl(
      args,
      context,
      "Azure App Service",
      "https://aka.ms/teamsfx-actions/azure-app-service-deploy",
      ["driver.deploy.azureAppServiceDeployDetailSummary"],
      ["driver.deploy.notice.deployDryRunComplete"]
    );
    await chai
      .expect(impl.checkDeployStatus("", config, new TestLogProvider()))
      .to.be.rejectedWith(CheckDeploymentStatusTimeoutError);
  });

  it("checkDeployStatus 500 response", async () => {
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: HttpStatusCode.INTERNAL_SERVER_ERROR,
      data: {
        status: DeployStatus.Failed,
        message: "fail to start app due to some reasons.",
      },
    });
    const config = {
      headers: {
        "Content-Type": "text",
        "Cache-Control": "no-cache",
        Authorization: "no",
      },
      maxContentLength: 200,
      maxBodyLength: 200,
      timeout: 200,
    };
    const args = {
      workingDirectory: "/",
      artifactFolder: "/",
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-11111111111/resourceGroups/hoho-rg/providers/Microsoft.Web/sites",
    } as DeployArgs;
    const context = {
      logProvider: new TestLogProvider(),
      ui: new MockUserInteraction(),
    } as any;
    const impl = new AzureZipDeployImpl(
      args,
      context,
      "Azure App Service",
      "https://aka.ms/teamsfx-actions/azure-app-service-deploy",
      ["driver.deploy.azureAppServiceDeployDetailSummary"],
      ["driver.deploy.notice.deployDryRunComplete"]
    );
    await chai
      .expect(impl.checkDeployStatus("", config, new TestLogProvider()))
      .to.be.rejectedWith(CheckDeploymentStatusError);
  });

  it("checkDeployStatus reject AxiosError", async () => {
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").rejects({
      isAxiosError: true,
      response: {
        status: 400,
        data: {
          error: {
            code: "Request_BadRequest",
            message:
              "Invalid value specified for property 'displayName' of resource 'Application'.",
          },
        },
      },
    });
    const config = {
      headers: {
        "Content-Type": "text",
        "Cache-Control": "no-cache",
        Authorization: "no",
      },
      maxContentLength: 200,
      maxBodyLength: 200,
      timeout: 200,
    };
    const args = {
      workingDirectory: "/",
      artifactFolder: "/",
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-11111111111/resourceGroups/hoho-rg/providers/Microsoft.Web/sites",
    } as DeployArgs;
    const context = {
      logProvider: new TestLogProvider(),
      ui: new MockUserInteraction(),
    } as any;
    const impl = new AzureZipDeployImpl(
      args,
      context,
      "Azure App Service",
      "https://aka.ms/teamsfx-actions/azure-app-service-deploy",
      ["driver.deploy.azureAppServiceDeployDetailSummary"],
      ["driver.deploy.notice.deployDryRunComplete"]
    );
    await chai
      .expect(impl.checkDeployStatus("", config, new TestLogProvider()))
      .to.be.rejectedWith(CheckDeploymentStatusError);
  });
  it("checkDeployStatus reject none AxiosError", async () => {
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").rejects(new Error("other error"));
    const config = {
      headers: {
        "Content-Type": "text",
        "Cache-Control": "no-cache",
        Authorization: "no",
      },
      maxContentLength: 200,
      maxBodyLength: 200,
      timeout: 200,
    };
    const args = {
      workingDirectory: "/",
      artifactFolder: "/",
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-11111111111/resourceGroups/hoho-rg/providers/Microsoft.Web/sites",
    } as DeployArgs;
    const context = {
      logProvider: new TestLogProvider(),
      ui: new MockUserInteraction(),
    } as any;
    const impl = new AzureZipDeployImpl(
      args,
      context,
      "Azure App Service",
      "https://aka.ms/teamsfx-actions/azure-app-service-deploy",
      ["driver.deploy.azureAppServiceDeployDetailSummary"],
      ["driver.deploy.notice.deployDryRunComplete"]
    );
    await chai
      .expect(impl.checkDeployStatus("", config, new TestLogProvider()))
      .to.be.rejectedWith(CheckDeploymentStatusError);
  });
  it("checkDeployStatus DeployRemoteStartError", async () => {
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: HttpStatusCode.OK,
      data: {
        status: DeployStatus.Failed,
        message: "fail to start app due to some reasons.",
      },
    });
    const config = {
      headers: {
        "Content-Type": "text",
        "Cache-Control": "no-cache",
        Authorization: "no",
      },
      maxContentLength: 200,
      maxBodyLength: 200,
      timeout: 200,
    };
    const args = {
      workingDirectory: "/",
      artifactFolder: "/",
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-11111111111/resourceGroups/hoho-rg/providers/Microsoft.Web/sites",
    } as DeployArgs;
    const context = {
      logProvider: new TestLogProvider(),
      ui: new MockUserInteraction(),
    } as any;
    const impl = new AzureZipDeployImpl(
      args,
      context,
      "Azure App Service",
      "https://aka.ms/teamsfx-actions/azure-app-service-deploy",
      ["driver.deploy.azureAppServiceDeployDetailSummary"],
      ["driver.deploy.notice.deployDryRunComplete"]
    );
    const res = await impl.checkDeployStatus("", config, new TestLogProvider());
    chai.assert.equal(res?.status, DeployStatus.Failed);
    chai.assert.equal(res?.message, "fail to start app due to some reasons.");
  });
  it("checkDeployStatus return status 400", async () => {
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "get").resolves({
      status: HttpStatusCode.BAD_REQUEST,
    });
    const config = {
      headers: {
        "Content-Type": "text",
        "Cache-Control": "no-cache",
        Authorization: "no",
      },
      maxContentLength: 200,
      maxBodyLength: 200,
      timeout: 200,
    };
    const args = {
      workingDirectory: "/",
      artifactFolder: "/",
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-11111111111/resourceGroups/hoho-rg/providers/Microsoft.Web/sites",
    } as DeployArgs;
    const context = {
      logProvider: new TestLogProvider(),
      ui: new MockUserInteraction(),
    } as any;
    const impl = new AzureZipDeployImpl(
      args,
      context,
      "Azure App Service",
      "https://aka.ms/teamsfx-actions/azure-app-service-deploy",
      ["driver.deploy.azureAppServiceDeployDetailSummary"],
      ["driver.deploy.notice.deployDryRunComplete"]
    );
    await chai
      .expect(impl.checkDeployStatus("", config, new TestLogProvider()))
      .to.be.rejectedWith(CheckDeploymentStatusError);
  });
  it("createAzureDeployConfig GetPublishingCredentialsError", async () => {
    const args = {
      workingDirectory: "/",
      artifactFolder: "/",
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-11111111111/resourceGroups/hoho-rg/providers/Microsoft.Web/sites",
    } as DeployArgs;
    const context = {
      logProvider: new TestLogProvider(),
      ui: new MockUserInteraction(),
      telemetryReporter: new MockTelemetryReporter(),
    } as any;
    const impl = new AzureZipDeployImpl(
      args,
      context,
      "Azure App Service",
      "https://aka.ms/teamsfx-actions/azure-app-service-deploy",
      ["driver.deploy.azureAppServiceDeployDetailSummary"],
      ["driver.deploy.notice.deployDryRunComplete"]
    );
    const webApps = {
      beginListPublishingCredentialsAndWait: async function (
        resourceGroupName: string,
        name: string
      ): Promise<WebAppsListPublishingCredentialsResponse> {
        throw new RestError("test message", { statusCode: 500 });
      },
    };
    const mockWebSiteManagementClient = new WebSiteManagementClient(new MyTokenCredential(), "sub");
    mockWebSiteManagementClient.webApps = webApps as any;
    sandbox.stub(appService, "WebSiteManagementClient").returns(mockWebSiteManagementClient);
    const token = new MyTokenCredential();
    sandbox.stub(token, "getToken").throws(new Error("test message"));
    await chai
      .expect(
        impl.createAzureDeployConfig(
          {
            subscriptionId: "e24d88be-bbbb-1234-ba25-11111111111",
            resourceGroupName: "mockGroupName",
            instanceId: "mockAppName",
          },
          token
        )
      )
      .to.be.rejectedWith(GetPublishingCredentialsError);
  });

  it("zipDeployPackage DeployZipPackageError throw 500", async () => {
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").rejects({
      isAxiosError: true,
      response: {
        status: 500,
        data: {
          error: {
            code: "InternalServerError",
            message: "Internal server error",
          },
        },
      },
    });
    const args = {
      workingDirectory: "/",
      artifactFolder: "/",
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-11111111111/resourceGroups/hoho-rg/providers/Microsoft.Web/sites",
    } as DeployArgs;
    const context = {
      logProvider: new TestLogProvider(),
      ui: new MockUserInteraction(),
    } as any;
    const impl = new AzureZipDeployImpl(
      args,
      context,
      "Azure App Service",
      "https://aka.ms/teamsfx-actions/azure-app-service-deploy",
      ["driver.deploy.azureAppServiceDeployDetailSummary"],
      ["driver.deploy.notice.deployDryRunComplete"]
    );
    const config = {
      headers: {
        "Content-Type": "text",
        "Cache-Control": "no-cache",
        Authorization: "no",
      },
      maxContentLength: 200,
      maxBodyLength: 200,
      timeout: 200,
    };
    await chai
      .expect(
        impl.zipDeployPackage(
          "mockEndPoint",
          fs.createReadStream(tempFile),
          config,
          new TestLogProvider()
        )
      )
      .to.be.rejectedWith(DeployZipPackageError);
  });
  it("zipDeployPackage DeployZipPackageError throw 404", async () => {
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").rejects({
      isAxiosError: true,
      response: {
        status: 400,
        data: {
          error: {
            code: "Request_BadRequest",
            message:
              "Invalid value specified for property 'displayName' of resource 'Application'.",
          },
        },
      },
    });
    const args = {
      workingDirectory: "/",
      artifactFolder: "/",
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-11111111111/resourceGroups/hoho-rg/providers/Microsoft.Web/sites",
    } as DeployArgs;
    const context = {
      logProvider: new TestLogProvider(),
      ui: new MockUserInteraction(),
    } as any;
    const impl = new AzureZipDeployImpl(
      args,
      context,
      "Azure App Service",
      "https://aka.ms/teamsfx-actions/azure-app-service-deploy",
      ["driver.deploy.azureAppServiceDeployDetailSummary"],
      ["driver.deploy.notice.deployDryRunComplete"]
    );
    const config = {
      headers: {
        "Content-Type": "text",
        "Cache-Control": "no-cache",
        Authorization: "no",
      },
      maxContentLength: 200,
      maxBodyLength: 200,
      timeout: 200,
    };
    await chai
      .expect(
        impl.zipDeployPackage(
          "mockEndPoint",
          fs.createReadStream(tempFile),
          config,
          new TestLogProvider()
        )
      )
      .to.be.rejectedWith(DeployZipPackageError);
  });
  it("zipDeployPackage DeployZipPackageError return 500", async () => {
    sandbox.stub(AzureDeployImpl.AXIOS_INSTANCE, "post").resolves({
      headers: {
        location: "abc",
      },
      status: 500,
    });
    const args = {
      workingDirectory: "/",
      artifactFolder: "/",
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-11111111111/resourceGroups/hoho-rg/providers/Microsoft.Web/sites",
    } as DeployArgs;
    const context = {
      logProvider: new TestLogProvider(),
      ui: new MockUserInteraction(),
    } as any;
    const impl = new AzureZipDeployImpl(
      args,
      context,
      "Azure App Service",
      "https://aka.ms/teamsfx-actions/azure-app-service-deploy",
      ["driver.deploy.azureAppServiceDeployDetailSummary"],
      ["driver.deploy.notice.deployDryRunComplete"]
    );
    const config = {
      headers: {
        "Content-Type": "text",
        "Cache-Control": "no-cache",
        Authorization: "no",
      },
      maxContentLength: 200,
      maxBodyLength: 200,
      timeout: 200,
    };
    await chai
      .expect(
        impl.zipDeployPackage(
          "mockEndPoint",
          fs.createReadStream(tempFile),
          config,
          new TestLogProvider()
        )
      )
      .to.be.rejectedWith(DeployZipPackageError);
  });

  it("throws Error when no basic auth allowed and Microsoft Entra request fail", async () => {
    const args = {
      workingDirectory: "/",
      artifactFolder: "/",
      ignoreFile: "./ignore",
      resourceId:
        "/subscriptions/e24d88be-bbbb-1234-ba25-11111111111/resourceGroups/hoho-rg/providers/Microsoft.Web/sites",
    } as DeployArgs;
    const context = {
      logProvider: new TestLogProvider(),
      ui: new MockUserInteraction(),
      telemetryReporter: new MockTelemetryReporter(),
    } as any;
    const impl = new AzureZipDeployImpl(
      args,
      context,
      "Azure App Service",
      "https://aka.ms/teamsfx-actions/azure-app-service-deploy",
      ["driver.deploy.azureAppServiceDeployDetailSummary"],
      ["driver.deploy.notice.deployDryRunComplete"]
    );
    process.env["TEAMSFX_AAD_DEPLOY_ONLY"] = "true";
    const webApps = {
      beginListPublishingCredentialsAndWait: async function (
        resourceGroupName: string,
        name: string
      ): Promise<WebAppsListPublishingCredentialsResponse> {
        throw new RestError("test message", { statusCode: 500 });
      },
    };
    const mockWebSiteManagementClient = new WebSiteManagementClient(new MyTokenCredential(), "sub");
    mockWebSiteManagementClient.webApps = webApps as any;
    sandbox.stub(appService, "WebSiteManagementClient").returns(mockWebSiteManagementClient);
    const token = new MyTokenCredential();
    sandbox.stub(token, "getToken").throws(new Error("test message"));
    await chai
      .expect(
        impl.createAzureDeployConfig(
          {
            subscriptionId: "e24d88be-bbbb-1234-ba25-11111111111",
            resourceGroupName: "mockGroupName",
            instanceId: "mockAppName",
          },
          token
        )
      )
      .to.be.rejectedWith(GetPublishingCredentialsError);
  });
});
