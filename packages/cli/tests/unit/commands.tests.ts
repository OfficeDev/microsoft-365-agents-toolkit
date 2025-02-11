import { CLIContext, SystemError, err, ok } from "@microsoft/teamsfx-api";
import {
  CollaborationStateResult,
  FeatureFlags,
  FuncToolChecker,
  FxCore,
  ListCollaboratorResult,
  LocalCertificateManager,
  LtsNodeChecker,
  PackageService,
  PermissionsResult,
  QuestionNames,
  UserCancelError,
  envUtil,
  featureFlagManager,
} from "@microsoft/teamsfx-core";
import * as tools from "@microsoft/teamsfx-core/build/common/tools";
import { assert } from "chai";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as sinon from "sinon";
import * as activate from "../../src/activate";
import { localTelemetryReporter } from "../../src/cmds/preview/localTelemetryReporter";
import {
  accountLoginAzureCommand,
  accountLoginM365Command,
  accountLogoutCommand,
  accountShowCommand,
  accountUtils,
  addCommand,
  addSPFxWebpartCommand,
  createSampleCommand,
  deployCommand,
  envAddCommand,
  envListCommand,
  getCreateCommand,
  helpCommand,
  listSamplesCommand,
  listTemplatesCommand,
  m365LaunchInfoCommand,
  m365SideloadingCommand,
  m365UnacquireCommand,
  m365utils,
  packageCommand,
  permissionGrantCommand,
  permissionStatusCommand,
  previewCommand,
  provisionCommand,
  publishCommand,
  upgradeCommand,
  validateCommand,
} from "../../src/commands/models";
import { DoctorChecker, teamsappDoctorCommand } from "../../src/commands/models/teamsapp/doctor";
import { teamsappPackageCommand } from "../../src/commands/models/teamsapp/package";
import { teamsappPublishCommand } from "../../src/commands/models/teamsapp/publish";
import { teamsappUpdateCommand } from "../../src/commands/models/teamsapp/update";
import { teamsappValidateCommand } from "../../src/commands/models/teamsapp/validate";
import AzureTokenProvider from "../../src/commonlib/azureLogin";
import { signedIn, signedOut } from "../../src/commonlib/common/constant";
import { logger } from "../../src/commonlib/logger";
import M365TokenProvider from "../../src/commonlib/m365Login";
import { MissingRequiredOptionError } from "../../src/error";
import * as utils from "../../src/utils";
import * as settingHelper from "@microsoft/teamsfx-core/build/common/projectSettingsHelper";
import { entraAppUpdateCommand } from "../../src/commands/models/entraAppUpdate";
import AzureTokenCIProvider from "../../src/commonlib/azureLoginCI";
import { envResetCommand } from "../../src/commands/models/envReset";
import { addPluginCommand } from "../../src/commands/models/addPlugin";
import { addAuthConfigCommand } from "../../src/commands/models/addAuthConfig";

describe("CLI commands", () => {
  const sandbox = sinon.createSandbox();

  let mockedEnvRestore: RestoreFn;

  beforeEach(() => {
    sandbox.stub(logger, "info").resolves(true);
    sandbox.stub(logger, "error").resolves(true);
  });

  afterEach(() => {
    sandbox.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });

  describe("getCreateCommand", async () => {
    it("happy path for donet", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "createProject").resolves(ok({ projectPath: "..." }));
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "new" },
        optionValues: {
          capabilities: "bot",
          nonInteractive: true,
        },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await getCreateCommand().handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("happy path for cli", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "createProject").resolves(ok({ projectPath: "..." }));
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "new" },
        optionValues: {
          capabilities: "bot",
          nonInteractive: true,
        },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await getCreateCommand().handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("core return error", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "createProject").resolves(err(new UserCancelError()));
      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "new" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await getCreateCommand().handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });

  describe("createSampleCommand", async () => {
    it("happy path", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "createSampleProject").resolves(ok({ projectPath: "..." }));
      const ctx: CLIContext = {
        command: { ...createSampleCommand, fullName: "new sample" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await createSampleCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("core return error", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "createProject").resolves(err(new UserCancelError()));
      const ctx: CLIContext = {
        command: { ...createSampleCommand, fullName: "new sample" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await createSampleCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });
  describe("listSampleCommand", async () => {
    it("happy path", async () => {
      sandbox.stub(utils, "getTemplates").resolves([]);
      const ctx: CLIContext = {
        command: { ...listSamplesCommand, fullName: "teamsapp list samples" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await listSamplesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("accountLoginAzureCommand", async () => {
    it("should success when service-principal = false", async () => {
      sandbox.stub(AzureTokenProvider, "signout");
      sandbox.stub(accountUtils, "outputAzureInfo").resolves();
      const ctx: CLIContext = {
        command: { ...accountLoginAzureCommand, fullName: "teamsapp auth login azure" },
        optionValues: { "service-principal": false },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountLoginAzureCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("should fail when service-principal = true", async () => {
      sandbox.stub(AzureTokenProvider, "signout");
      sandbox.stub(accountUtils, "outputAzureInfo").resolves();
      const ctx: CLIContext = {
        command: { ...accountLoginAzureCommand, fullName: "teamsapp auth login azure" },
        optionValues: { "service-principal": true },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountLoginAzureCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("should fail service-principal = false", async () => {
      sandbox.stub(AzureTokenProvider, "signout");
      sandbox.stub(accountUtils, "outputAzureInfo").resolves();
      const ctx: CLIContext = {
        command: { ...accountLoginAzureCommand, fullName: "teamsapp auth login azure" },
        optionValues: { "service-principal": false, username: "abc" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountLoginAzureCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });
  describe("accountLoginM365Command", async () => {
    it("should success", async () => {
      sandbox.stub(M365TokenProvider, "signout");
      sandbox.stub(accountUtils, "outputM365Info").resolves();
      const ctx: CLIContext = {
        command: { ...accountLoginM365Command, fullName: "teamsapp auth login m365" },
        optionValues: { "service-principal": false },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountLoginM365Command.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });

  describe("addSPFxWebpartCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "addWebpart").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...addSPFxWebpartCommand, fullName: "add spfx-web-part" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await addSPFxWebpartCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });

  describe("addPluginCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "addPlugin").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...addPluginCommand, fullName: "add plugin" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await addPluginCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });

  describe("getAddCommand", async () => {
    it("customize GPT is enabled", async () => {
      const commands = addCommand();
      assert.isTrue(commands.commands?.length === 3);
    });
  });

  describe("deployCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "deployArtifacts").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...deployCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await deployCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("success for customized yaml path", async () => {
      sandbox.stub(FxCore.prototype, "deployArtifacts").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...deployCommand, fullName: "teamsfx" },
        optionValues: { "config-file-path": "fakePath" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await deployCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("envAddCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "createEnv").resolves(ok(undefined));
      sandbox.stub(settingHelper, "isValidProjectV3").returns(true);
      const ctx: CLIContext = {
        command: { ...envAddCommand, fullName: "teamsfx" },
        optionValues: { projectPath: "." },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await envAddCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("isValidProjectV3: false", async () => {
      sandbox.stub(FxCore.prototype, "createEnv").resolves(ok(undefined));
      sandbox.stub(settingHelper, "isValidProjectV3").returns(false);
      const ctx: CLIContext = {
        command: { ...envAddCommand, fullName: "teamsfx" },
        optionValues: { projectPath: "." },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await envAddCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });
  describe("envListCommand", async () => {
    it("success", async () => {
      sandbox.stub(settingHelper, "isValidProjectV3").returns(true);
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev"]));
      const ctx: CLIContext = {
        command: { ...envListCommand, fullName: "teamsfx" },
        optionValues: { projectPath: "." },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await envListCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("isValidProjectV3: false", async () => {
      sandbox.stub(settingHelper, "isValidProjectV3").returns(false);
      const ctx: CLIContext = {
        command: { ...envListCommand, fullName: "teamsfx" },
        optionValues: { projectPath: "." },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await envListCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("listEnv error", async () => {
      sandbox.stub(settingHelper, "isValidProjectV3").returns(true);
      sandbox.stub(envUtil, "listEnv").resolves(err(new UserCancelError()));
      const ctx: CLIContext = {
        command: { ...envListCommand, fullName: "teamsfx" },
        optionValues: { projectPath: "." },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await envListCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });
  describe("envResetCommand", async () => {
    it("success with env", async () => {
      sandbox.stub(envUtil, "resetEnv").resolves();
      const ctx: CLIContext = {
        command: { ...envAddCommand, fullName: "teamsapp env reset" },
        optionValues: { env: "dev", projectPath: "." },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await envResetCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("success with env file", async () => {
      sandbox.stub(envUtil, "resetEnvFile").resolves();
      const ctx: CLIContext = {
        command: { ...envAddCommand, fullName: "teamsapp env reset" },
        optionValues: { "env-file": ".env.dev" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await envResetCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("provisionCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "provisionResources").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...provisionCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await provisionCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("non interactive mode", async () => {
      sandbox.stub(FxCore.prototype, "provisionResources").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...provisionCommand, fullName: "teamsfx" },
        optionValues: { nonInteractive: true, region: "East US" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await provisionCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("packageCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "createAppPackage").resolves(ok({ state: "OK" }));
      const ctx: CLIContext = {
        command: { ...packageCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await packageCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("permissionGrantCommand", async () => {
    it("success interactive = false", async () => {
      sandbox
        .stub(FxCore.prototype, "grantPermission")
        .resolves(ok({ state: "OK" } as PermissionsResult));
      const ctx: CLIContext = {
        command: { ...permissionGrantCommand, fullName: "teamsfx" },
        optionValues: { "manifest-file-path": "abc" },
        globalOptionValues: { interactive: false },
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionGrantCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("success interactive = true", async () => {
      sandbox
        .stub(FxCore.prototype, "grantPermission")
        .resolves(ok({ state: "OK" } as PermissionsResult));
      const ctx: CLIContext = {
        command: { ...permissionGrantCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: { interactive: true },
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionGrantCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("missing option", async () => {
      sandbox
        .stub(FxCore.prototype, "grantPermission")
        .resolves(ok({ state: "OK" } as PermissionsResult));
      const ctx: CLIContext = {
        command: { ...permissionGrantCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: { interactive: false },
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionGrantCommand.handler!(ctx);
      assert.isTrue(res.isErr() && res.error instanceof MissingRequiredOptionError);
    });
  });
  describe("permissionStatusCommand", async () => {
    it("listCollaborator", async () => {
      sandbox
        .stub(FxCore.prototype, "listCollaborator")
        .resolves(ok({ state: "OK" } as ListCollaboratorResult));
      const ctx: CLIContext = {
        command: { ...permissionStatusCommand, fullName: "teamsfx" },
        optionValues: { all: true },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionStatusCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("checkPermission", async () => {
      sandbox
        .stub(FxCore.prototype, "checkPermission")
        .resolves(ok({ state: "OK" } as CollaborationStateResult));
      const ctx: CLIContext = {
        command: { ...permissionStatusCommand, fullName: "teamsfx" },
        optionValues: { all: false },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionStatusCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("publishCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "publishApplication").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...publishCommand, fullName: "teamsfx" },
        optionValues: { env: "local" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await publishCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("previewCommand", async () => {
    it("success", async () => {
      sandbox.stub(localTelemetryReporter, "runWithTelemetryGeneric").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...previewCommand, fullName: "teamsfx" },
        optionValues: { env: "local" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await previewCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("error", async () => {
      sandbox
        .stub(localTelemetryReporter, "runWithTelemetryGeneric")
        .resolves(err(new UserCancelError()));
      const ctx: CLIContext = {
        command: { ...previewCommand, fullName: "teamsfx" },
        optionValues: { env: "local" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await previewCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });
  describe("entraAppUpdateCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "deployAadManifest").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...entraAppUpdateCommand, fullName: "teamsapp entraapp update" },
        optionValues: {
          env: "local",
          projectPath: "./",
          "manifest-file-path": "./aad.manifest.json",
        },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await entraAppUpdateCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("upgradeCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "phantomMigrationV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...upgradeCommand, fullName: "teamsfx" },
        optionValues: { force: true },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await upgradeCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("validateCommand", async () => {
    it("conflict", async () => {
      sandbox.stub(FxCore.prototype, "validateApplication").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...validateCommand, fullName: "teamsfx" },
        optionValues: { "manifest-path": "aaa", "app-package-file-path": "bbb" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await validateCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("none", async () => {
      sandbox.stub(FxCore.prototype, "validateApplication").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...validateCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await validateCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("manifest", async () => {
      sandbox.stub(FxCore.prototype, "validateApplication").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...validateCommand, fullName: "teamsfx" },
        optionValues: { "manifest-path": "aaa", env: "dev" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await validateCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("manifest missing env", async () => {
      sandbox.stub(FxCore.prototype, "validateApplication").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...validateCommand, fullName: "teamsfx" },
        optionValues: { "manifest-path": "aaa" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await validateCommand.handler!(ctx);
      assert.isTrue(res.isErr() && res.error instanceof MissingRequiredOptionError);
    });
    it("package", async () => {
      sandbox.stub(FxCore.prototype, "validateApplication").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...validateCommand, fullName: "teamsfx" },
        optionValues: { "app-package-file-path": "bbb" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await validateCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });

  describe("m365LaunchInfoCommand", async () => {
    beforeEach(() => {
      sandbox.stub(logger, "warning");
    });
    it("success retrieveTitleId", async () => {
      sandbox.stub(m365utils, "getTokenAndUpn").resolves(["token", "upn"]);
      sandbox.stub(PackageService.prototype, "retrieveTitleId").resolves("id");
      sandbox.stub(PackageService.prototype, "getLaunchInfoByTitleId").resolves("id");
      const ctx: CLIContext = {
        command: { ...m365LaunchInfoCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365LaunchInfoCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("success", async () => {
      sandbox.stub(m365utils, "getTokenAndUpn").resolves(["token", "upn"]);
      sandbox.stub(PackageService.prototype, "getLaunchInfoByTitleId").resolves("id");
      const ctx: CLIContext = {
        command: { ...m365LaunchInfoCommand, fullName: "teamsfx" },
        optionValues: { "title-id": "aaa" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365LaunchInfoCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("MissingRequiredOptionError", async () => {
      sandbox.stub(m365utils, "getTokenAndUpn").resolves(["token", "upn"]);
      const ctx: CLIContext = {
        command: { ...m365LaunchInfoCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365LaunchInfoCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });

  describe("m365SideloadingCommand", async () => {
    beforeEach(() => {
      sandbox.stub(logger, "warning");
    });

    describe("M365Utils - getTokenAndUpn", async () => {
      it("getAccessToken fail", async () => {
        sandbox.stub(M365TokenProvider, "getAccessToken").resolves(err(new UserCancelError()));
        try {
          await m365utils.getTokenAndUpn();
          assert.fail("should not reach here");
        } catch (e) {
          assert.isTrue(e instanceof UserCancelError);
        }
      });
      it("getStatus fail", async () => {
        sandbox.stub(M365TokenProvider, "getAccessToken").resolves(ok("token"));
        sandbox.stub(M365TokenProvider, "getStatus").resolves(err(new UserCancelError()));
        const res = await m365utils.getTokenAndUpn();
        assert.deepEqual(res, ["token", undefined]);
      });
      it("getStatus ok", async () => {
        sandbox.stub(M365TokenProvider, "getAccessToken").resolves(ok("token"));
        sandbox
          .stub(M365TokenProvider, "getStatus")
          .resolves(ok({ accountInfo: { upn: "test" } } as any));
        const res = await m365utils.getTokenAndUpn();
        assert.deepEqual(res, ["token", "test"]);
      });
      it("getStatus throw error", async () => {
        sandbox.stub(M365TokenProvider, "getAccessToken").resolves(ok("token"));
        sandbox.stub(M365TokenProvider, "getStatus").rejects(new Error());
        const res = await m365utils.getTokenAndUpn();
        assert.deepEqual(res, ["token", undefined]);
      });
    });

    it("should success with zip package", async () => {
      sandbox.stub(m365utils, "getTokenAndUpn").resolves(["token", "upn"]);
      sandbox.stub(PackageService.prototype, "sideLoading").resolves();
      const ctx: CLIContext = {
        command: { ...m365SideloadingCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa", "file-path": "./" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365SideloadingCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("should success with xml", async () => {
      sandbox.stub(m365utils, "getTokenAndUpn").resolves(["token", "upn"]);
      sandbox.stub(PackageService.prototype, "sideLoadXmlManifest").resolves();
      const ctx: CLIContext = {
        command: { ...m365SideloadingCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa", "xml-path": "./" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365SideloadingCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("should fail if both zip and xml are provided", async () => {
      const ctx: CLIContext = {
        command: { ...m365SideloadingCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa", "xml-path": "./", "file-path": "./" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365SideloadingCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("should fail if non of zip and xml are provided", async () => {
      const ctx: CLIContext = {
        command: { ...m365SideloadingCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365SideloadingCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });

  describe("m365UnacquireCommand", async () => {
    beforeEach(() => {
      sandbox.stub(logger, "warning");
    });
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "uninstall").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...m365UnacquireCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365UnacquireCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("failed", async () => {
      sandbox.stub(FxCore.prototype, "uninstall").resolves(err(new SystemError("", "", "")));
      const ctx: CLIContext = {
        command: { ...m365UnacquireCommand, fullName: "teamsfx" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365UnacquireCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });

  describe("v3 commands", async () => {
    beforeEach(() => {
      sandbox.stub(logger, "warning");
    });
    afterEach(() => {
      sandbox.restore();
    });
    it("update", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "updateTeamsAppCLIV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...teamsappUpdateCommand, fullName: "teamsapp update" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappUpdateCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("update conflict", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "updateTeamsAppCLIV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...teamsappUpdateCommand, fullName: "teamsapp update" },
        optionValues: { "manifest-file": "manifest.json", "package-file": "package.zip" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappUpdateCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("package", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "packageTeamsAppCLIV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...teamsappPackageCommand, fullName: "teamsapp package" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappPackageCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("validate", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "validateTeamsAppCLIV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...teamsappValidateCommand, fullName: "teamsapp validate" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappValidateCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("validate conflict", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "validateTeamsAppCLIV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...teamsappValidateCommand, fullName: "teamsapp validate" },
        optionValues: { "manifest-file": "manifest.json", "package-file": "package.zip" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappValidateCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
    it("publish", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "publishTeamsAppCLIV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...teamsappPublishCommand, fullName: "teamsapp publish" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappPublishCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("publish conflict", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "publishTeamsAppCLIV3").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...teamsappPublishCommand, fullName: "teamsapp publish" },
        optionValues: { "manifest-file": "manifest.json", "package-file": "package.zip" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappPublishCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });

  describe("addAuthConfigCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "addAuthAction").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...addAuthConfigCommand, fullName: "add auth-config" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await addAuthConfigCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
});

describe("CLI read-only commands", () => {
  const sandbox = sinon.createSandbox();

  let messages: string[] = [];

  beforeEach(() => {
    sandbox.stub(logger, "info").callsFake(async (message: string) => {
      messages.push(message);
      return true;
    });
    sandbox.stub(logger, "error").callsFake(async (message: string) => {
      messages.push(message);
      return true;
    });
    sandbox.stub(logger, "outputInfo").callsFake(async (message: string) => {
      messages.push(message);
      return true;
    });
    sandbox.stub(logger, "outputError").callsFake(async (message: string) => {
      messages.push(message);
      return true;
    });
  });

  afterEach(() => {
    messages = [];
    sandbox.restore();
  });
  describe("AccountUtils", async () => {
    it("outputAccountInfoOffline", async () => {
      const res = accountUtils.outputAccountInfoOffline("m365", "xxx");
      assert.isTrue(res);
    });
    it("outputM365Info login success", async () => {
      sandbox.stub(M365TokenProvider, "getJsonObject").resolves(ok({ upn: "fakename" }));
      sandbox.stub(M365TokenProvider, "getTenant").resolves(undefined);
      const res = await accountUtils.outputM365Info("login");
      assert.isTrue(res);
    });
    context("outputM365Info login under hosting tenant", () => {
      let mocks: RestoreFn;
      beforeEach(() => {
        mocks = mockedEnv({ TEAMSFX_MULTI_TENANT: "true" });
        sandbox.stub(M365TokenProvider, "getJsonObject").resolves(ok({ unique_name: "fakename" }));
        sandbox.stub(M365TokenProvider, "getTenant").resolves("faked_tenant_id");
      });

      afterEach(() => {
        mocks();
      });

      it("specified tenant name displayed", async () => {
        sandbox.stub(M365TokenProvider, "getAccessToken").resolves(ok("token"));
        sandbox
          .stub(tools, "listAllTenants")
          .resolves([
            { tenantId: "faked_tid_1" },
            { tenantId: "faked_tenant_id", displayName: "Test tenant" },
          ]);
        const res = await accountUtils.outputM365Info("login", "faked_tenant_id");
        assert.isTrue(res);
      });

      it("specified tenant not match", async () => {
        sandbox.stub(M365TokenProvider, "getAccessToken").resolves(ok("token"));
        sandbox
          .stub(tools, "listAllTenants")
          .resolves([{ tenantId: "faked_tid_1" }, { tenantId: "faked_tid_2" }]);
        const res = await accountUtils.outputM365Info("login", "faked_tenant_id");
        assert.isTrue(res);
      });

      it("failed to retrieve access token", async () => {
        sandbox
          .stub(M365TokenProvider, "getAccessToken")
          .resolves(err("failed to get access token" as any));
        const res = await accountUtils.outputM365Info("login", "faked_tenant_id");
        assert.isTrue(res);
      });
    });
    it("outputM365Info login fail", async () => {
      sandbox.stub(M365TokenProvider, "getJsonObject").resolves(err(new UserCancelError()));
      const res = await accountUtils.outputM365Info("login");
      assert.isFalse(res);
    });
    it("outputM365Info show success", async () => {
      sandbox.stub(M365TokenProvider, "getJsonObject").resolves(ok({ upn: "fakename" }));
      sandbox.stub(M365TokenProvider, "getTenant").resolves("faked_tenant_id");
      sandbox.stub(M365TokenProvider, "getAccessToken").resolves(ok("token"));
      sandbox
        .stub(tools, "listAllTenants")
        .resolves([{ tenantId: "faked_tid_1" }, { tenantId: "faked_tenant_id" }]);
      const res = await accountUtils.outputM365Info("show");
      assert.isTrue(res);
    });
    it("outputM365Info show fail", async () => {
      sandbox.stub(M365TokenProvider, "getJsonObject").resolves(err(new UserCancelError()));
      const res = await accountUtils.outputM365Info("show");
      assert.isFalse(res);
    });
    it("outputAzureInfo login", async () => {
      sandbox.stub(AzureTokenCIProvider, "load").resolves();
      sandbox.stub(AzureTokenCIProvider, "init").resolves();
      sandbox.stub(AzureTokenCIProvider, "getJsonObject").resolves({ upn: "test" });
      sandbox.stub(AzureTokenCIProvider, "listSubscriptions").resolves([]);
      const res = await accountUtils.outputAzureInfo("login", undefined, true);
      assert.isTrue(res);
    });
    it("outputAzureInfo login with tenant parameter", async () => {
      const mockedEnvRestore = mockedEnv({ TEAMSFX_MULTI_TENANT: "true" });
      sandbox.stub(AzureTokenCIProvider, "load").resolves();
      sandbox.stub(AzureTokenCIProvider, "init").resolves();
      sandbox.stub(AzureTokenCIProvider, "switchTenant").resolves();
      sandbox.stub(AzureTokenCIProvider, "getJsonObject").resolves({ unique_name: "test" });
      sandbox.stub(AzureTokenCIProvider, "listSubscriptions").resolves([]);
      sandbox.stub(AzureTokenCIProvider, "getTenant").resolves("faked_tenant_id");
      sandbox.stub(AzureTokenCIProvider, "getIdentityCredentialAsync").resolves({
        getToken: async () => {
          return Promise.resolve({ token: "faked_token" });
        },
      } as any);
      sandbox
        .stub(tools, "listAllTenants")
        .resolves([{ tenantId: "faked_tid_1" }, { tenantId: "faked_tenant_id" }]);
      const res = await accountUtils.outputAzureInfo("login", "faked_tenant_id", true);
      assert.isTrue(res);
      mockedEnvRestore();
    });
    it("outputAzureInfo login fail with tenant parameter - invalid token", async () => {
      const mockedEnvRestore = mockedEnv({ TEAMSFX_MULTI_TENANT: "true" });
      sandbox.stub(AzureTokenCIProvider, "load").resolves();
      sandbox.stub(AzureTokenCIProvider, "init").resolves();
      sandbox.stub(AzureTokenCIProvider, "switchTenant").resolves();
      sandbox.stub(AzureTokenCIProvider, "getJsonObject").resolves({ unique_name: "test" });
      sandbox.stub(AzureTokenCIProvider, "listSubscriptions").resolves([]);
      sandbox.stub(AzureTokenCIProvider, "getTenant").resolves("faked_tenant_id");
      sandbox.stub(AzureTokenCIProvider, "getIdentityCredentialAsync").resolves(undefined);
      const res = await accountUtils.outputAzureInfo("login", "faked_tenant_id", true);
      assert.isTrue(res);
      mockedEnvRestore();
    });
    it("outputAzureInfo login fail with tenant parameter - tenant mismatch", async () => {
      const mockedEnvRestore = mockedEnv({ TEAMSFX_MULTI_TENANT: "true" });
      sandbox.stub(AzureTokenCIProvider, "load").resolves();
      sandbox.stub(AzureTokenCIProvider, "init").resolves();
      sandbox.stub(AzureTokenCIProvider, "switchTenant").resolves();
      sandbox.stub(AzureTokenCIProvider, "getJsonObject").resolves({ unique_name: "test" });
      sandbox.stub(AzureTokenCIProvider, "listSubscriptions").resolves([]);
      sandbox.stub(AzureTokenCIProvider, "getTenant").resolves("faked_tenant_id");
      sandbox.stub(AzureTokenCIProvider, "getIdentityCredentialAsync").resolves({
        getToken: async () => {
          return Promise.resolve({ token: "faked_token" });
        },
      } as any);
      sandbox
        .stub(tools, "listAllTenants")
        .resolves([{ tenantId: "faked_tid_1" }, { tenantId: "faked_tid_2" }]);
      const res = await accountUtils.outputAzureInfo("login", "faked_tenant_id", true);
      assert.isTrue(res);
      mockedEnvRestore();
    });
    it("outputAzureInfo login fail", async () => {
      sandbox.stub(AzureTokenProvider, "getJsonObject").resolves(undefined);
      const res = await accountUtils.outputAzureInfo("login");
      assert.isFalse(res);
    });
    it("outputAzureInfo show", async () => {
      sandbox.stub(AzureTokenProvider, "getJsonObject").resolves({ upn: "test" });
      sandbox.stub(AzureTokenProvider, "listSubscriptions").resolves([]);
      const res = await accountUtils.outputAzureInfo("show");
      assert.isTrue(res);
    });
    it("outputAzureInfo show fail", async () => {
      sandbox.stub(AzureTokenProvider, "getJsonObject").resolves(undefined);
      const res = await accountUtils.outputAzureInfo("show");
      assert.isFalse(res);
    });
  });
  describe("accountShowCommand", async () => {
    it("both signedOut", async () => {
      sandbox.stub(M365TokenProvider, "getStatus").resolves(ok({ status: signedOut }));
      sandbox.stub(AzureTokenProvider, "getStatus").resolves({ status: signedOut });
      messages = [];
      const ctx: CLIContext = {
        command: { ...accountShowCommand, fullName: "teamsapp auth list" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountShowCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("both signedIn and checkIsOnline = true", async () => {
      sandbox.stub(M365TokenProvider, "getStatus").resolves(ok({ status: signedIn }));
      sandbox.stub(AzureTokenProvider, "getStatus").resolves({ status: signedIn });
      sandbox.stub(accountUtils, "checkIsOnline").resolves(true);
      const outputM365Info = sandbox.stub(accountUtils, "outputM365Info").resolves();
      const outputAzureInfo = sandbox.stub(accountUtils, "outputAzureInfo").resolves();
      messages = [];
      const ctx: CLIContext = {
        command: { ...accountShowCommand, fullName: "teamsapp auth list" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountShowCommand.handler!(ctx);
      assert.isTrue(res.isOk());
      assert.isTrue(outputM365Info.calledOnce);
      assert.isTrue(outputAzureInfo.calledOnce);
    });
    it("both signedIn and checkIsOnline = false", async () => {
      sandbox
        .stub(M365TokenProvider, "getStatus")
        .resolves(ok({ status: signedIn, accountInfo: { upn: "xxx" } }));
      sandbox
        .stub(AzureTokenProvider, "getStatus")
        .resolves({ status: signedIn, accountInfo: { upn: "xxx" } });
      sandbox.stub(accountUtils, "checkIsOnline").resolves(false);
      const outputAccountInfoOffline = sandbox.stub(accountUtils, "outputAccountInfoOffline");
      messages = [];
      const ctx: CLIContext = {
        command: { ...accountShowCommand, fullName: "teamsapp auth list" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountShowCommand.handler!(ctx);
      assert.isTrue(res.isOk());
      assert.isTrue(outputAccountInfoOffline.calledTwice);
    });
    it("M365TokenProvider.getStatus() returns error", async () => {
      sandbox.stub(M365TokenProvider, "getStatus").resolves(err(new UserCancelError()));
      messages = [];
      const ctx: CLIContext = {
        command: { ...accountShowCommand, fullName: "teamsapp auth list" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await accountShowCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });

  describe("accountLogoutCommand", async () => {
    it("azure success", async () => {
      sandbox.stub(AzureTokenProvider, "signout").resolves(true);
      const ctx: CLIContext = {
        command: { ...accountLogoutCommand, fullName: "teamsapp auth logout" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["azure"],
        telemetryProperties: {},
      };
      messages = [];
      const res = await accountLogoutCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("azure fail", async () => {
      sandbox.stub(AzureTokenProvider, "signout").resolves(false);
      const ctx: CLIContext = {
        command: { ...accountLogoutCommand, fullName: "teamsapp auth logout" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["azure"],
        telemetryProperties: {},
      };
      messages = [];
      const res = await accountLogoutCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("m365 success", async () => {
      sandbox.stub(M365TokenProvider, "signout").resolves(true);
      const ctx: CLIContext = {
        command: { ...accountLogoutCommand, fullName: "teamsapp auth logout" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["m365"],
        telemetryProperties: {},
      };
      const res = await accountLogoutCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("m365 fail", async () => {
      sandbox.stub(M365TokenProvider, "signout").resolves(false);
      const ctx: CLIContext = {
        command: { ...accountLogoutCommand, fullName: "teamsapp auth logout" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: ["m365"],
        telemetryProperties: {},
      };
      messages = [];
      const res = await accountLogoutCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("listTemplatesCommand", async () => {
    let mockedEnvRestore: RestoreFn;
    afterEach(() => {
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });
    it("happy path", async () => {
      const ctx: CLIContext = {
        command: { ...listTemplatesCommand, fullName: "list" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await listTemplatesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("table with description", async () => {
      const ctx: CLIContext = {
        command: { ...listTemplatesCommand, fullName: "..." },
        optionValues: { format: "table", description: true },
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await listTemplatesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("table without description", async () => {
      const ctx: CLIContext = {
        command: { ...listTemplatesCommand, fullName: "..." },
        optionValues: { format: "table", description: false },
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await listTemplatesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("listSamplesCommand", async () => {
    it("json", async () => {
      sandbox.stub(utils, "getTemplates").resolves([]);
      const ctx: CLIContext = {
        command: { ...listSamplesCommand, fullName: "..." },
        optionValues: { format: "json" },
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await listSamplesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("table with filter + description", async () => {
      sandbox.stub(utils, "getTemplates").resolves([]);
      const ctx: CLIContext = {
        command: { ...listSamplesCommand, fullName: "..." },
        optionValues: { tag: "tab", format: "table", description: true },
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await listSamplesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("table without description", async () => {
      sandbox.stub(utils, "getTemplates").resolves([]);
      const ctx: CLIContext = {
        command: { ...listSamplesCommand, fullName: "..." },
        optionValues: { format: "table", description: false },
        globalOptionValues: {},
        argumentValues: ["key", "value"],
        telemetryProperties: {},
      };
      const res = await listSamplesCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("helpCommand", async () => {
    it("happy", async () => {
      const ctx: CLIContext = {
        command: { ...helpCommand, fullName: "..." },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await helpCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });

  describe("doctor", async () => {
    describe("checkAccount", async () => {
      it("checkAccount error", async () => {
        sandbox
          .stub(DoctorChecker.prototype, "checkM365Account")
          .resolves(err(new UserCancelError()));
        const checker = new DoctorChecker();
        await checker.checkAccount();
      });
      it("checkAccount success", async () => {
        sandbox.stub(DoctorChecker.prototype, "checkM365Account").resolves(ok("success"));
        const checker = new DoctorChecker();
        await checker.checkAccount();
      });
    });
    describe("checkM365Account", async () => {
      it("checkM365Account - signin", async () => {
        const token = "test-token";
        const tenantId = "test-tenant-id";
        const upn = "test-user";
        sandbox.stub(M365TokenProvider, "getStatus").returns(
          Promise.resolve(
            ok({
              status: signedIn,
              token: token,
              accountInfo: {
                tid: tenantId,
                upn: upn,
              },
            })
          )
        );
        sandbox.stub(tools, "getSideloadingStatus").resolves(true);
        const checker = new DoctorChecker();
        const accountRes = await checker.checkM365Account();
        assert.isTrue(accountRes.isOk());
        const account = (accountRes as any).value;
        assert.include(account, "is signed in and custom app upload permission is enabled");
      });
      it("checkM365Account - error", async () => {
        sandbox.stub(M365TokenProvider, "getStatus").resolves(err(new UserCancelError()));
        sandbox.stub(tools, "getSideloadingStatus").resolves(true);
        const checker = new DoctorChecker();
        const accountRes = await checker.checkM365Account();
        assert.isTrue(accountRes.isOk());
        const account = (accountRes as any).value;
        assert.include(account, "You've not signed into your Microsoft 365 account yet.");
      });
      it("checkM365Account - error2", async () => {
        sandbox.stub(M365TokenProvider, "getStatus").rejects(new Error("test"));
        sandbox.stub(tools, "getSideloadingStatus").resolves(true);
        const checker = new DoctorChecker();
        const accountRes = await checker.checkM365Account();
        assert.isTrue(accountRes.isErr());
      });
      it("checkM365Account - signout", async () => {
        const token = "test-token";
        const tenantId = "test-tenant-id";
        const upn = "test-user";
        const getStatusStub = sandbox.stub(M365TokenProvider, "getStatus");
        getStatusStub.onCall(0).resolves(
          ok({
            status: signedOut,
          })
        );
        getStatusStub.onCall(1).resolves(
          ok({
            status: signedIn,
            token: token,
            accountInfo: {
              tid: tenantId,
              upn: upn,
            },
          })
        );
        sandbox.stub(M365TokenProvider, "getAccessToken").resolves(ok(token));
        sandbox.stub(tools, "getSideloadingStatus").resolves(true);
        const checker = new DoctorChecker();
        const accountRes = await checker.checkM365Account();
        assert.isTrue(accountRes.isOk());
        const account = (accountRes as any).value;
        assert.include(account, "is signed in and custom app upload permission is enabled");
      });

      it("checkM365Account - no custom app upload permission", async () => {
        const token = "test-token";
        const tenantId = "test-tenant-id";
        const upn = "test-user";
        sandbox.stub(M365TokenProvider, "getStatus").returns(
          Promise.resolve(
            ok({
              status: signedIn,
              token: token,
              accountInfo: {
                tid: tenantId,
                upn: upn,
              },
            })
          )
        );
        sandbox.stub(tools, "getSideloadingStatus").resolves(false);
        const checker = new DoctorChecker();
        const accountRes = await checker.checkM365Account();
        assert.isTrue(accountRes.isOk());
        const value = (accountRes as any).value;
        assert.include(
          value,
          "Your Microsoft 365 tenant admin hasn't enabled custom app upload permission for your account"
        );
      });
    });

    describe("checkNodejs", async () => {
      it("installed", async () => {
        sandbox
          .stub(LtsNodeChecker.prototype, "getInstallationInfo")
          .resolves({ isInstalled: true } as any);
        const checker = new DoctorChecker();
        await checker.checkNodejs();
      });
      it("error", async () => {
        sandbox
          .stub(LtsNodeChecker.prototype, "getInstallationInfo")
          .resolves({ isInstalled: true, error: new UserCancelError() } as any);
        const checker = new DoctorChecker();
        await checker.checkNodejs();
      });
      it("not installed", async () => {
        sandbox
          .stub(LtsNodeChecker.prototype, "getInstallationInfo")
          .resolves({ isInstalled: false } as any);
        const checker = new DoctorChecker();
        await checker.checkNodejs();
      });
    });
    describe("checkFuncCoreTool", async () => {
      it("installed", async () => {
        sandbox
          .stub(FuncToolChecker.prototype, "queryFuncVersion")
          .resolves({ versionStr: "3.0" } as any);
        const checker = new DoctorChecker();
        await checker.checkFuncCoreTool();
      });
      it("not installed", async () => {
        sandbox.stub(FuncToolChecker.prototype, "queryFuncVersion").rejects(new Error());
        const checker = new DoctorChecker();
        await checker.checkFuncCoreTool();
      });
    });
    describe("checkCert", async () => {
      it("not found", async () => {
        sandbox
          .stub(LocalCertificateManager.prototype, "setupCertificate")
          .resolves({ found: false } as any);
        const checker = new DoctorChecker();
        await checker.checkCert();
      });
      it("found trusted", async () => {
        sandbox
          .stub(LocalCertificateManager.prototype, "setupCertificate")
          .resolves({ found: true, alreadyTrusted: true } as any);
        const checker = new DoctorChecker();
        await checker.checkCert();
      });
      it("found not trusted", async () => {
        sandbox
          .stub(LocalCertificateManager.prototype, "setupCertificate")
          .resolves({ found: true, alreadyTrusted: false } as any);
        const checker = new DoctorChecker();
        await checker.checkCert();
      });
    });
    it("happy", async () => {
      sandbox.stub(DoctorChecker.prototype, "checkAccount").resolves();
      sandbox.stub(DoctorChecker.prototype, "checkNodejs").resolves();
      sandbox.stub(DoctorChecker.prototype, "checkFuncCoreTool").resolves();
      sandbox.stub(DoctorChecker.prototype, "checkCert").resolves();
      const ctx: CLIContext = {
        command: { ...teamsappDoctorCommand, fullName: "teamsapp doctor" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await teamsappDoctorCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
});
