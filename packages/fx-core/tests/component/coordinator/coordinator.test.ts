import "mocha";

import { assert } from "chai";
import fs from "fs-extra";
import * as sinon from "sinon";
import { Container } from "typedi";

import {
  err,
  FxError,
  Inputs,
  InputsWithProjectPath,
  LogProvider,
  ok,
  Platform,
  Result,
  UserError,
} from "@microsoft/teamsfx-api";

import { createContext, setTools } from "../../../src/common/globalVars";
import { MetadataV3, VersionInfo, VersionSource } from "../../../src/common/versionMetadata";
import {
  DriverInstance,
  ExecutionError,
  ExecutionOutput,
  ExecutionResult,
  ProjectModel,
} from "../../../src/component/configManager/interface";
import { coordinator } from "../../../src/component/coordinator";
import { DriverContext } from "../../../src/component/driver/interface/commonArgs";
import { ExecutionResult as DriverExecutionResult } from "../../../src/component/driver/interface/stepDriver";
import * as appStudio from "../../../src/component/driver/teamsApp/appStudio";
import { CreateAppPackageDriver } from "../../../src/component/driver/teamsApp/createAppPackage";
import { manifestUtils } from "../../../src/component/driver/teamsApp/utils/ManifestUtils";
import { ValidateManifestDriver } from "../../../src/component/driver/teamsApp/validate";
import { ValidateAppPackageDriver } from "../../../src/component/driver/teamsApp/validateAppPackage";
import { envUtil } from "../../../src/component/utils/envUtil";
import { metadataUtil } from "../../../src/component/utils/metadataUtil";
import { pathUtils } from "../../../src/component/utils/pathUtils";
import { settingsUtil } from "../../../src/component/utils/settingsUtil";
import { FxCore } from "../../../src/core/FxCore";
import * as v3MigrationUtils from "../../../src/core/middleware/utils/v3MigrationUtils";
import { MissingEnvironmentVariablesError } from "../../../src/error/common";
import { QuestionNames } from "../../../src/question";
import { MockedAzureAccountProvider, MockedM365Provider, MockTools } from "../../core/utils";

export function mockedResolveDriverInstances(log: LogProvider): Result<DriverInstance[], FxError> {
  return ok([
    {
      uses: "arm/deploy",
      with: undefined,
      instance: {
        execute: async (args: unknown, context: DriverContext): Promise<DriverExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
      },
    },
  ]);
}

const versionInfo: VersionInfo = {
  version: MetadataV3.projectVersion,
  source: VersionSource.teamsapp,
};
const V3Version = MetadataV3.projectVersion;
describe("component coordinator test", () => {
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  setTools(tools);
  afterEach(() => {
    sandbox.restore();
  });

  beforeEach(() => {
    sandbox.stub(v3MigrationUtils, "getProjectVersion").resolves(versionInfo);
  });

  describe("convertExecuteResult", () => {
    it("convertExecuteResult ok", async () => {
      const value = new Map([["key", "value"]]);
      const res: Result<ExecutionOutput, ExecutionError> = ok(value);
      const convertRes = coordinator.convertExecuteResult(res, ".");
      assert.deepEqual(convertRes[0], { key: "value" });
      assert.isUndefined(convertRes[1]);
    });

    it("convertExecuteResult Failure", async () => {
      const error = new UserError({ source: "test", name: "TestError", message: "test message" });
      const res: Result<ExecutionOutput, ExecutionError> = err({ kind: "Failure", error: error });
      const convertRes = coordinator.convertExecuteResult(res, ".");
      assert.deepEqual(convertRes[0], {});
      assert.equal(convertRes[1], error);
    });

    it("convertExecuteResult PartialSuccess - DriverError", async () => {
      const value = new Map([["key", "value"]]);
      const error = new UserError({ source: "test", name: "TestError", message: "test message" });
      const res: Result<ExecutionOutput, ExecutionError> = err({
        kind: "PartialSuccess",
        env: value,
        reason: {
          kind: "DriverError",
          error: error,
          failedDriver: { name: "TestDriver", uses: "testUse", with: "testWith" },
        },
      });
      const convertRes = coordinator.convertExecuteResult(res, ".");
      assert.deepEqual(convertRes[0], { key: "value" });
      assert.equal(convertRes[1], error);
    });

    it("convertExecuteResult PartialSuccess - MissingEnvironmentVariablesError", async () => {
      const value = new Map([["key", "value"]]);
      const res: Result<ExecutionOutput, ExecutionError> = err({
        kind: "PartialSuccess",
        env: value,
        reason: {
          kind: "UnresolvedPlaceholders",
          unresolvedPlaceHolders: ["TEST_PL"],
          failedDriver: { name: "TestDriver", uses: "testUse", with: "testWith" },
        },
      });
      const convertRes = coordinator.convertExecuteResult(res, ".");
      assert.deepEqual(convertRes[0], { key: "value" });
      assert.isTrue(convertRes[1]! instanceof MissingEnvironmentVariablesError);
    });
  });

  it("preProvisionForVS", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      registerApp: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "arm/deploy",
            with: {
              subscriptionId: "mockSubId",
              resourceGroupName: "mockRG",
            },
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
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(fs, "pathExistsSync").returns(true);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      env: "dev",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.preProvisionForVS(inputs);
    assert.isTrue(res.isOk());
    if (res.isOk()) {
      const value = res.value;
      assert.isTrue(value.needAzureLogin);
      assert.isTrue(value.needM365Login);
      assert.equal(value.resolvedAzureSubscriptionId, "mockSubId");
      assert.equal(value.resolvedAzureResourceGroupName, "mockRG");
    }
  });
  it("preCheckYmlAndEnvForVS", async () => {
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
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(fs, "pathExistsSync").returns(true);
    sandbox.stub(process, "env").value({ TEAMS_APP_ID: "faked_id" });
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      env: "local",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.preCheckYmlAndEnvForVS(inputs);
    assert.isTrue(res.isOk());
  });

  it("preCheckYmlAndEnvForVS - happy pass with empty provision actions", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(fs, "pathExistsSync").returns(true);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      env: "local",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.preCheckYmlAndEnvForVS(inputs);
    assert.isTrue(res.isOk());
  });

  it("preCheckYmlAndEnvForVS - happy pass without teamsApp/create action", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "botFramework/create",
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
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(fs, "pathExistsSync").returns(true);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      env: "local",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.preCheckYmlAndEnvForVS(inputs);
    assert.isTrue(res.isOk());
  });

  it("fail to get project model in preCheckYmlAndEnvForVS", async () => {
    sandbox.stub(metadataUtil, "parse").resolves(err(new UserError({})));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(fs, "pathExistsSync").returns(true);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      env: "local",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.preCheckYmlAndEnvForVS(inputs);
    assert.isTrue(res.isErr());
  });

  it("unresolvePlaceholders in preCheckYmlAndEnvForVS", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "botFramework/create",
            with: {
              botId: "${{BOT_ID}}",
            },
          },
        ],
        resolvePlaceholders: () => {
          return ["BotId"];
        },
        execute: async (ctx: DriverContext): Promise<ExecutionResult> => {
          return { result: ok(new Map()), summaries: [] };
        },
        resolveDriverInstances: mockedResolveDriverInstances,
      },
    };
    sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(fs, "pathExistsSync").returns(true);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      env: "local",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.preCheckYmlAndEnvForVS(inputs);
    assert.isTrue(res.isErr());
  });

  it("empty teams app id in preCheckYmlAndEnvForVS", async () => {
    const mockProjectModel: ProjectModel = {
      version: "1.0.0",
      provision: {
        name: "configureApp",
        driverDefs: [
          {
            uses: "teamsApp/create",
            with: undefined,
            writeToEnvironmentFile: { teamsAppId: "CUSTOMIZED_TEAMS_APP_ID" },
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
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
    sandbox.stub(fs, "pathExistsSync").returns(true);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      env: "local",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.preCheckYmlAndEnvForVS(inputs);
    assert.isTrue(res.isErr());
  });

  it("executeUserTaskNew", async () => {
    sandbox.stub(envUtil, "listEnv").resolves(ok(["dev"]));
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
    sandbox.stub(manifestUtils, "getTeamsAppManifestPath").resolves("");
    const driver1: ValidateManifestDriver = Container.get("teamsApp/validateManifest");
    const driver2: CreateAppPackageDriver = Container.get("teamsApp/zipAppPackage");
    const driver3: ValidateAppPackageDriver = Container.get("teamsApp/validateAppPackage");
    const mockRes: DriverExecutionResult = {
      summaries: [],
      result: ok(new Map()),
    };
    sandbox.stub(driver1, "execute").resolves(mockRes);
    sandbox.stub(driver2, "execute").resolves(mockRes);
    sandbox.stub(driver3, "execute").resolves(mockRes);
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: ".",
      ignoreLockByUT: true,
    };
    const fxCore = new FxCore(tools);
    const res1 = await fxCore.executeUserTask(
      { namespace: "", method: "getManifestTemplatePath", params: { manifestTemplatePath: "." } },
      inputs
    );
    if (res1.isErr()) console.log(res1.error);
    assert.isTrue(res1.isOk());
    const res2 = await fxCore.executeUserTask(
      { namespace: "", method: "validateManifest", params: { manifestPath: "." } },
      inputs
    );
    if (res2.isErr()) console.log(res2.error);
    assert.isTrue(res2.isOk());
    const res3 = await fxCore.executeUserTask(
      {
        namespace: "",
        method: "buildPackage",
        params: { manifestTemplatePath: ".", outputZipPath: ".", outputJsonPath: "." },
      },
      inputs
    );
    if (res3.isErr()) console.log(res3.error);
    assert.isTrue(res3.isOk());
    const res4 = await fxCore.executeUserTask(
      { namespace: "", method: "validateManifest", params: { appPackagePath: "." } },
      inputs
    );
    assert.isTrue(res4.isOk());
  });

  describe("getDotEnvs error", () => {
    it("getDotEnvs success", async () => {
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev1", "dev2"]));
      sandbox.stub(envUtil, "readEnv").resolves(ok({ k1: "v1" }));
      const inputs: InputsWithProjectPath = {
        platform: Platform.VSCode,
        projectPath: ".",
      };
      const fxCore = new FxCore(tools);
      const res = await fxCore.getDotEnvs(inputs);
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.deepEqual(Object.keys(res.value), ["dev1", "dev2"]);
      }
    });
    it("getDotEnvs error 1", async () => {
      sandbox.stub(envUtil, "listEnv").resolves(err(new UserError({})));
      const inputs: InputsWithProjectPath = {
        platform: Platform.VSCode,
        projectPath: ".",
      };
      const fxCore = new FxCore(tools);
      const res = await fxCore.getDotEnvs(inputs);
      assert.isTrue(res.isErr());
    });
    it("getDotEnvs error 2", async () => {
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev1", "dev2"]));
      sandbox.stub(envUtil, "readEnv").resolves(err(new UserError({})));
      const inputs: InputsWithProjectPath = {
        platform: Platform.VSCode,
        projectPath: ".",
      };
      const fxCore = new FxCore(tools);
      const res = await fxCore.getDotEnvs(inputs);
      assert.isTrue(res.isErr());
    });
  });

  it("getSelectedEnv", async () => {
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    const inputs: InputsWithProjectPath = {
      platform: Platform.VSCode,
      projectPath: ".",
      env: "dev",
    };
    const fxCore = new FxCore(tools);
    const res = await fxCore.getSelectedEnv(inputs);
    assert.isTrue(res.isOk());
  });

  describe("encrypt/decrypt", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("error", async () => {
      sandbox.stub(settingsUtil, "readSettings").resolves(err(new UserError({})));
      const inputs: InputsWithProjectPath = {
        platform: Platform.VSCode,
        projectPath: ".",
      };
      const fxCore = new FxCore(tools);
      const inputText = "abc";
      const res = await fxCore.encrypt(inputText, inputs);
      assert.isTrue(res.isErr());
      const res2 = await fxCore.decrypt("abc", inputs);
      assert.isTrue(res2.isErr());
    });
    it("happy path", async () => {
      sandbox
        .stub(settingsUtil, "readSettings")
        .resolves(ok({ version: "1", trackingId: "mockid" }));
      const inputs: InputsWithProjectPath = {
        platform: Platform.VSCode,
        projectPath: ".",
      };
      const fxCore = new FxCore(tools);
      const inputText = "abc";
      const res = await fxCore.encrypt(inputText, inputs);
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        const res2 = await fxCore.decrypt(res.value, inputs);
        assert.isTrue(res2.isOk());
        if (res2.isOk()) {
          assert.equal(res2.value, inputText);
        }
      }
    });
  });

  describe("publishInDeveloperPortal", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("missing token provider", async () => {
      const context = createContext();
      context.tokenProvider = undefined;
      const inputs: InputsWithProjectPath = {
        platform: Platform.VSCode,
        projectPath: "project-path",
        [QuestionNames.AppPackagePath]: "path",
      };
      const res = await coordinator.publishInDeveloperPortal(context, inputs);
      assert.isTrue(res.isErr());
    });

    it("missing appPackagePath", async () => {
      const context = createContext();
      context.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      const inputs: InputsWithProjectPath = {
        platform: Platform.VSCode,
        projectPath: "project-path",
      };
      const res = await coordinator.publishInDeveloperPortal(context, inputs);
      assert.isTrue(res.isErr());
    });

    it("success", async () => {
      const context = createContext();
      context.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      sandbox
        .stub(context.tokenProvider.m365TokenProvider, "getJsonObject")
        .resolves(ok({ unique_name: "test" }));
      sandbox.stub(appStudio, "updateTeamsAppV3ForPublish").resolves(ok("appId"));
      const openUrl = sandbox.stub(context.userInteraction, "openUrl").resolves(ok(true));
      const inputs: InputsWithProjectPath = {
        platform: Platform.VSCode,
        projectPath: "project-path",
        [QuestionNames.AppPackagePath]: "path",
      };

      const res = await coordinator.publishInDeveloperPortal(context, inputs);
      assert.isTrue(res.isOk());
      assert.isTrue(openUrl.calledOnce);
    });

    it("update manifest error", async () => {
      const context = createContext();
      context.tokenProvider = {
        m365TokenProvider: new MockedM365Provider(),
        azureAccountProvider: new MockedAzureAccountProvider(),
      };
      sandbox
        .stub(appStudio, "updateTeamsAppV3ForPublish")
        .resolves(err(new UserError("source", "error", "", "")));
      const inputs: InputsWithProjectPath = {
        platform: Platform.VSCode,
        projectPath: "project-path",
        [QuestionNames.AppPackagePath]: "path",
      };

      const res = await coordinator.publishInDeveloperPortal(context, inputs);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "error");
      }
    });

    it("ensureTeamsFxInCsproj  no .csproj found", async () => {
      sandbox.stub(fs, "readdir").resolves([] as any);
      const res = await coordinator.ensureTeamsFxInCsproj(".");
      assert.isTrue(res.isOk());
    });

    it("ensureTeamsFxInCsproj success: do nothing for existing ItemGroup", async () => {
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Project Sdk="Microsoft.NET.Sdk">
        <ItemGroup>
          <ProjectCapability Include="TeamsFx"/>
        </ItemGroup>
      </Project>`;
      sandbox.stub(fs, "readdir").resolves(["test.csproj"] as any);
      sandbox.stub(fs, "readFile").resolves(xml as any);
      const res = await coordinator.ensureTeamsFxInCsproj(".");
      assert.isTrue(res.isOk());
    });

    it("ensureTeamsFxInCsproj success: insert one", async () => {
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Project Sdk="Microsoft.NET.Sdk">
      </Project>`;
      sandbox.stub(fs, "readdir").resolves(["test.csproj"] as any);
      sandbox.stub(fs, "readFile").resolves(xml as any);
      sandbox.stub(fs, "writeFile").resolves();
      const res = await coordinator.ensureTeamsFxInCsproj(".");
      assert.isTrue(res.isOk());
    });
  });
});
