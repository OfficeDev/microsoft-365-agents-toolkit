import "mocha";

import { assert } from "chai";
import fs, { PathLike } from "fs-extra";
import * as os from "os";
import * as path from "path";
import * as sinon from "sinon";

import {
  err,
  Inputs,
  IProgressHandler,
  ok,
  Platform,
  SystemError,
  UserError,
} from "@microsoft/teamsfx-api";

import { MetadataV3, VersionInfo, VersionSource } from "../../../src/common/versionMetadata";
import { ExecutionResult, ProjectModel } from "../../../src/component/configManager/interface";
import { SummaryReporter } from "../../../src/component/coordinator/summary";
import { DriverContext } from "../../../src/component/driver/interface/commonArgs";
import { provisionUtils } from "../../../src/component/provisionUtils";
import { dotenvUtil, envUtil } from "../../../src/component/utils/envUtil";
import { metadataUtil } from "../../../src/component/utils/metadataUtil";
import { pathUtils } from "../../../src/component/utils/pathUtils";
import { resourceGroupHelper } from "../../../src/component/utils/ResourceGroupHelper";
import { settingsUtil } from "../../../src/component/utils/settingsUtil";
import { FxCore } from "../../../src/core/FxCore";
import { setTools } from "../../../src/common/globalVars";
import * as v3MigrationUtils from "../../../src/core/middleware/utils/v3MigrationUtils";
import {
  InvalidAzureCredentialError,
  ResourceGroupConflictError,
  SelectSubscriptionError,
} from "../../../src/error/azure";
import { UserCancelError } from "../../../src/error/common";
import { MockTools, randomAppName } from "../../core/utils";
import { mockedResolveDriverInstances } from "./coordinator.test";
import mockedEnv, { RestoreFn } from "mocked-env";

const versionInfo: VersionInfo = {
  version: MetadataV3.projectVersion,
  source: VersionSource.teamsapp,
};
const V3Version = MetadataV3.projectVersion;
describe("coordinator provision", () => {
  let envRestore: RestoreFn = () => {};
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  setTools(tools);
  afterEach(() => {
    sandbox.restore();
    envRestore?.();
  });

  beforeEach(() => {
    sandbox.stub(v3MigrationUtils, "getProjectVersion").resolves(versionInfo);
  });

  it("provision happy path from zero", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
      environmentFolderPath: "./envs",
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureSubscription").resolves(
      ok({
        subscriptionId: "mockSubId",
        tenantId: "mockTenantId",
        subscriptionName: "mockSubName",
      })
    );
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: true,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
      if (config.name === "env") {
        return ok({ type: "success", result: "dev" });
      } else {
        return ok({ type: "success", result: "" });
      }
    });
    sandbox.stub(resourceGroupHelper, "createNewResourceGroup").resolves(ok("test-rg"));
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(pathUtils, "getYmlFilePath").resolves(ok("m365agents.yml"));
    sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
    sandbox.stub(fs, "writeFile").resolves();
    const progressStartStub = sandbox.stub();
    const progressEndStub = sandbox.stub();
    sandbox.stub(tools.ui, "createProgressBar").returns({
      start: progressStartStub,
      end: progressEndStub,
    } as any as IProgressHandler);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
      isLocalDebug: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isOk());
    // getSelectedEnv
    const selectEnvRes = await fxCore.getSelectedEnv(inputs);
    if (selectEnvRes.isErr()) {
      console.log(selectEnvRes.error);
    }
    assert.isTrue(selectEnvRes.isOk());
    if (selectEnvRes.isOk()) {
      assert.equal(selectEnvRes.value, "dev");
    }
    assert.isTrue(progressStartStub.calledOnce);
    assert.isTrue(progressEndStub.calledOnceWithExactly(true));
  });
  it("provision success with subscriptionId in yml", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: {
              subscriptionId: "mockSubId",
            },
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: true,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(resourceGroupHelper, "createNewResourceGroup").resolves(ok("test-rg"));
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(pathUtils, "getYmlFilePath").resolves(ok("m365agents.yml"));
    sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isOk());
  });
  it("provision happy path from zero case 2", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(provisionUtils, "ensureSubscription").resolves(
      ok({
        subscriptionId: "mockSubId",
        tenantId: "mockTenantId",
        subscriptionName: "mockSubName",
      })
    );
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: true,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });
    sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
      if (config.name === "env") {
        return ok({ type: "success", result: "dev" });
      } else {
        return ok({ type: "success", result: "" });
      }
    });
    sandbox.stub(resourceGroupHelper, "createNewResourceGroup").resolves(ok("test-rg"));
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(pathUtils, "getYmlFilePath").resolves(ok("m365agents.yml"));
    sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isOk());
  });
  it("provision happy path: validate multi-env", async () => {
    envRestore = mockedEnv({
      TEAMSFX_ENV: "dev",
    });
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return [];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          const map = new Map();
          map.set("KEY1", "VALUE1");
          map.set("SECRET_KEY2", "VALUE2");
          return { result: ok(map), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(provisionUtils, "ensureSubscription").resolves(
      ok({
        subscriptionId: "mockSubId",
        tenantId: "mockTenantId",
        subscriptionName: "mockSubName",
      })
    );
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: true,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });

    sandbox.stub(resourceGroupHelper, "createNewResourceGroup").resolves(ok("test-rg"));

    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "ensureFile").resolves();
    sandbox
      .stub(settingsUtil, "readSettings")
      .resolves(ok({ version: "1.0", trackingId: "mockTrackingId" }));
    const fileDataMap = new Map();
    sandbox.stub(fs, "writeFile").callsFake(async (file: PathLike | number, data: any) => {
      fileDataMap.set(file, data);
    });
    const appName = randomAppName();
    const projectPath = path.resolve(os.tmpdir(), appName);
    const envFilePath = path.resolve(projectPath, "env", ".env.dev");
    const userDataFilePath = envFilePath + ".user";
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok(envFilePath));
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: projectPath,
      ignoreLockByUT: true,
      env: "dev",
      workflowFilePath: path.resolve(projectPath, "m365agents.yml"),
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isOk());
    const envData = fileDataMap.get(envFilePath);
    const pRes1 = dotenvUtil.deserialize(envData);
    const secretData = fileDataMap.get(userDataFilePath);
    const pRes2 = dotenvUtil.deserialize(secretData);
    assert.equal(pRes1.obj["KEY1"], "VALUE1");
    assert.isUndefined(pRes1.obj["SECRET_KEY2"]);
    assert.isTrue(pRes2.obj["SECRET_KEY2"].startsWith("crypto_"));
  });
  it("provision happy path with existing resource groups in VS Code", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(provisionUtils, "ensureSubscription").resolves(
      ok({
        subscriptionId: "mockSubId",
        tenantId: "mockTenantId",
        subscriptionName: "mockSubName",
      })
    );
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: false,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });
    sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
      if (config.name === "env") {
        return ok({ type: "success", result: "dev" });
      } else {
        return ok({ type: "success", result: "" });
      }
    });
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(pathUtils, "getYmlFilePath").resolves(ok("m365agents.yml"));
    sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isOk());
  });
  it("provision failed to get selected subscription", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(provisionUtils, "ensureSubscription").resolves(
      ok({
        subscriptionId: "mockSubId",
        tenantId: "mockTenantId",
        subscriptionName: "mockSubName",
      })
    );
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: true,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription")
      .resolves(undefined);
    sandbox.stub(resourceGroupHelper, "createNewResourceGroup").resolves(ok("test-rg"));
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(pathUtils, "getYmlFilePath").resolves(ok("m365agents.yml"));
    sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
      env: "dev",
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof SelectSubscriptionError);
    }
  });
  it("provision SPFx project shows success notification", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return [];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
      if (config.name === "env") {
        return ok({ type: "success", result: "dev" });
      } else {
        return ok({ type: "success", result: "" });
      }
    });
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(pathUtils, "getYmlFilePath").resolves(ok("m365agents.yml"));
    sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
    const stubShowMessage = sandbox.stub(tools.ui, "showMessage");

    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isOk());
    assert.isTrue(stubShowMessage.calledOnce);
  });
  it("provision failed when user directly update yml with empty subscriptionId", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: {
              subscriptionId: "",
            },
          },
        ],
        resolvePlaceholders: () => {
          return [];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isErr());
  });
  it("provision failed with parse error", async () => {
    sandbox.stub(metadataUtil, "parse").resolves(err(new UserError({})));
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      env: "dev",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isErr());
  });
  it("provision failed to get subInfo", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureSubscription").resolves(
      ok({
        subscriptionId: "mockSubId",
        tenantId: "mockTenantId",
        subscriptionName: "mockSubName",
      })
    );
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: true,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );

    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription")
      .resolves(undefined);
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();

    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      env: "dev",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isErr());
  });
  it("provision failed getLifecycleDescriptions Error", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(pathUtils, "getYmlFilePath").resolves(ok("m365agents.yml"));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureSubscription").resolves(
      ok({
        subscriptionId: "mockSubId",
        tenantId: "mockTenantId",
        subscriptionName: "mockSubName",
      })
    );
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: true,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
      if (config.name === "env") {
        return ok({ type: "success", result: "dev" });
      } else {
        return ok({ type: "success", result: "" });
      }
    });
    sandbox.stub(resourceGroupHelper, "createNewResourceGroup").resolves(ok("test-rg"));
    sandbox
      .stub(SummaryReporter.prototype, "getLifecycleDescriptions")
      .returns(err(new UserError({})));
    const progressStartStub = sandbox.stub();
    const progressEndStub = sandbox.stub();
    sandbox.stub(tools.ui, "createProgressBar").returns({
      start: progressStartStub,
      end: progressEndStub,
    } as any as IProgressHandler);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isErr());
    assert.isTrue(progressStartStub.calledOnce);
    assert.isTrue(progressEndStub.calledOnceWithExactly(false));
  });
  it("provision failed with partial success", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return {
            result: err({
              kind: "PartialSuccess",
              env: new Map(),
              reason: {
                kind: "DriverError",
                failedDriver: { uses: "", with: {} },
                error: new UserError({}),
              },
            }),
            summaries: [],
          };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(pathUtils, "getYmlFilePath").resolves(ok("m365agents.yml"));
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureSubscription").resolves(
      ok({
        subscriptionId: "mockSubId",
        tenantId: "mockTenantId",
        subscriptionName: "mockSubName",
      })
    );
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: true,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
      if (config.name === "env") {
        return ok({ type: "success", result: "dev" });
      } else {
        return ok({ type: "success", result: "" });
      }
    });
    sandbox.stub(resourceGroupHelper, "createNewResourceGroup").resolves(ok("test-rg"));
    const progressStartStub = sandbox.stub();
    const progressEndStub = sandbox.stub();
    sandbox.stub(tools.ui, "createProgressBar").returns({
      start: progressStartStub,
      end: progressEndStub,
    } as any as IProgressHandler);
    // sandbox
    //   .stub(SummaryReporter.prototype, "getLifecycleDescriptions")
    //   .resolves(err(new UserError({})));
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isErr());
    assert.isTrue(progressStartStub.calledOnce);
    assert.isTrue(progressEndStub.calledOnceWithExactly(false));
  });
  it("provision failed with getM365TenantId Error", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureSubscription").resolves(
      ok({
        subscriptionId: "mockSubId",
        tenantId: "mockTenantId",
        subscriptionName: "mockSubName",
      })
    );
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: true,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox
      .stub(provisionUtils, "getM365TenantId")
      .resolves(err(new UserError({ source: "Tst", name: "TestError" })));
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
      if (config.name === "env") {
        return ok({ type: "success", result: "dev" });
      } else {
        return ok({ type: "success", result: "" });
      }
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isErr());
  });
  it("provision failed with getSelectedSubscription Error", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureSubscription").resolves(
      ok({
        subscriptionId: "mockSubId",
        tenantId: "mockTenantId",
        subscriptionName: "mockSubName",
      })
    );
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: true,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription")
      .resolves(undefined);
    sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
      if (config.name === "env") {
        return ok({ type: "success", result: "dev" });
      } else {
        return ok({ type: "success", result: "" });
      }
    });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isErr());
  });
  it("provision happy path with CLI inputs", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(resourceGroupHelper, "createNewResourceGroup").resolves(ok("test-rg"));
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(pathUtils, "getYmlFilePath").resolves(ok("m365agents.yml"));
    sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: ".",
      env: "dev",
      targetSubscriptionId: "mockSubId",
      targetResourceGroupName: "test-rg",
      targetResourceLocationName: "Ease US",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isOk());
  });
  it("provision happy path with CLI inputs for existing resource group", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox
      .stub(resourceGroupHelper, "createNewResourceGroup")
      .resolves(err(new ResourceGroupConflictError("xxx", "sss")));
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(pathUtils, "getYmlFilePath").resolves(ok("m365agents.yml"));
    sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      env: "dev",
      targetSubscriptionId: "mockSubId",
      targetResourceGroupName: "test-rg",
      targetResourceLocationName: "Ease US",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isOk());
  });
  it("provision failed with CLI inputs: create resource group failed", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox
      .stub(resourceGroupHelper, "createNewResourceGroup")
      .resolves(err(new UserError({ source: "test", name: "OtherError" })));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: ".",
      env: "dev",
      targetSubscriptionId: "mockSubId",
      targetResourceGroupName: "test-rg",
      targetResourceLocationName: "Ease US",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isErr());
  });
  it("provision failed when getting azure credentials", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureSubscription").resolves(
      ok({
        subscriptionId: "mockSubId",
        tenantId: "mockTenantId",
        subscriptionName: "mockSubName",
      })
    );
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: true,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
      if (config.name === "env") {
        return ok({ type: "success", result: "dev" });
      } else {
        return ok({ type: "success", result: "" });
      }
    });
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getIdentityCredentialAsync")
      .resolves(undefined);
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(pathUtils, "getYmlFilePath").resolves(ok("m365agents.yml"));
    sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.isTrue(res.error instanceof InvalidAzureCredentialError);
    }
  });
  it("provision failed when checking resource group existence", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureSubscription").resolves(
      ok({
        subscriptionId: "mockSubId",
        tenantId: "mockTenantId",
        subscriptionName: "mockSubName",
      })
    );
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: true,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
      if (config.name === "env") {
        return ok({ type: "success", result: "dev" });
      } else {
        return ok({ type: "success", result: "" });
      }
    });
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getIdentityCredentialAsync").resolves({
      getToken: (scopes: string) => {
        return Promise.resolve({ token: "token", expiresOnTimestamp: 1 });
      },
    });
    sandbox
      .stub(resourceGroupHelper, "checkResourceGroupExistence")
      .resolves(err(new SystemError("test", "test", "", "")));
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(pathUtils, "getYmlFilePath").resolves(ok("m365agents.yml"));
    sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.name, "test");
    }
  });
  it("provision happy path (debug)", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [],
        resolvePlaceholders: () => {
          return [];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox
      .stub(settingsUtil, "readSettings")
      .resolves(ok({ trackingId: "mockId", version: V3Version }));
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      workflowFilePath: "./app.local.yml",
      env: "local",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    if (res.isErr()) {
      console.log(res?.error);
    }
    assert.isTrue(res.isOk());
  });

  it("provision happy path (VS debug)", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [],
        resolvePlaceholders: () => {
          return [];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox
      .stub(settingsUtil, "readSettings")
      .resolves(ok({ trackingId: "mockId", version: V3Version }));
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    const inputs: Inputs = {
      platform: Platform.VS,
      projectPath: ".",
      workflowFilePath: "./app.local.yml",
      env: "local",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    if (res.isErr()) {
      console.log(res?.error);
    }
    assert.isTrue(res.isOk());
  });

  it("provision failed with check whether m365 tenant matched fail", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureSubscription").resolves(
      ok({
        subscriptionId: "mockSubId",
        tenantId: "mockTenantId",
        subscriptionName: "mockSubName",
      })
    );
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: true,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox
      .stub(provisionUtils, "ensureM365TenantMatchesV3")
      .returns(err(new UserError("coordinator", "checkM365TenantError", "msg", "msg")));
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
      if (config.name === "env") {
        return ok({ type: "success", result: "dev" });
      } else {
        return ok({ type: "success", result: "" });
      }
    });
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(pathUtils, "getYmlFilePath").resolves(ok("m365agents.yml"));
    sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.name, "checkM365TenantError");
    }
  });
  it("provision failed with no subscription permission", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return [];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "setSubscription")
      .rejects(new UserError({ source: "Test", name: "NoPermission" }));
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      env: "dev",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isErr());
  });
  it("provision with no progress bar", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
      environmentFolderPath: "./envs",
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureSubscription").resolves(
      ok({
        subscriptionId: "mockSubId",
        tenantId: "mockTenantId",
        subscriptionName: "mockSubName",
      })
    );
    sandbox.stub(provisionUtils, "ensureResourceGroup").resolves(
      ok({
        createNewResourceGroup: true,
        name: "test-rg",
        location: "East US",
      })
    );
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(provisionUtils, "askForProvisionConsentV3").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription").resolves({
      subscriptionId: "mockSubId",
      tenantId: "mockTenantId",
      subscriptionName: "mockSubName",
    });
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "setSubscription").resolves();
    sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
      if (config.name === "env") {
        return ok({ type: "success", result: "dev" });
      } else {
        return ok({ type: "success", result: "" });
      }
    });
    sandbox.stub(resourceGroupHelper, "createNewResourceGroup").resolves(ok("test-rg"));
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(pathUtils, "getYmlFilePath").resolves(ok("m365agents.yml"));
    sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
    sandbox.stub(fs, "writeFile").resolves();
    sandbox.stub(tools.ui, "createProgressBar").returns(undefined as any as IProgressHandler);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
      isLocalDebug: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isOk());
    // getSelectedEnv
    const selectEnvRes = await fxCore.getSelectedEnv(inputs);
    if (selectEnvRes.isErr()) {
      console.log(selectEnvRes.error);
    }
    assert.isTrue(selectEnvRes.isOk());
    if (selectEnvRes.isOk()) {
      assert.equal(selectEnvRes.value, "dev");
    }
  });
  it("provision select subscription cancel", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      registerApp: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: undefined,
          },
          {
            uses: "teamsApp/create",
            with: undefined,
          },
        ],
        resolvePlaceholders: () => {
          return ["AZURE_SUBSCRIPTION_ID", "AZURE_RESOURCE_GROUP_NAME"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(provisionUtils, "ensureM365TenantMatchesV3").returns(ok(undefined));
    sandbox.stub(provisionUtils, "getM365TenantId").resolves(
      ok({
        tenantIdInToken: "mockM365Tenant",
        tenantUserName: "mockM365UserName",
      })
    );
    sandbox.stub(tools.tokenProvider.azureAccountProvider, "getIdentityCredentialAsync").resolves();
    sandbox
      .stub(tools.tokenProvider.azureAccountProvider, "getSelectedSubscription")
      .rejects(new UserCancelError());
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      env: "dev",
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.provisionResources(inputs);
    assert.isTrue(res.isErr());
  });
});
