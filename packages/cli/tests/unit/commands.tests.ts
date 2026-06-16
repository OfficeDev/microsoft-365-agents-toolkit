import { CLIContext, SystemError, err, ok } from "@microsoft/teamsfx-api";
import {
  CliQuestionName,
  CollaborationConstants,
  CollaborationStateResult,
  FxCore,
  ListCollaboratorResult,
  PackageService,
  PermissionGrantInputs,
  PermissionListInputs,
  PermissionsResult,
  QuestionNames,
  UserCancelError,
  envUtil,
  featureFlagManager,
} from "@microsoft/teamsfx-core";
import { assert } from "chai";
import { RestoreFn } from "mocked-env";
import * as sinon from "sinon";
import * as activate from "../../src/activate";
import { localTelemetryReporter } from "../../src/cmds/preview/localTelemetryReporter";
import {
  accountLoginAzureCommand,
  accountLoginM365Command,
  accountUtils,
  addCommand,
  addSPFxWebpartCommand,
  createSampleCommand,
  deployCommand,
  envAddCommand,
  envListCommand,
  getCreateCommand,
  listSamplesCommand,
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
  validateCommand,
} from "../../src/commands/models";
import { addAuthConfigCommand } from "../../src/commands/models/addAuthConfig";
import { addCapabilityCommand } from "../../src/commands/models/addCapability";
import { addPluginCommand } from "../../src/commands/models/addPlugin";
import { exportOpenPluginCommand } from "../../src/commands/models/exportOpenPlugin";
import { importOpenPluginCommand } from "../../src/commands/models/importOpenPlugin";
import { entraAppUpdateCommand } from "../../src/commands/models/entraAppUpdate";
import { envAddDeps } from "../../src/commands/models/envAdd";
import { envListDeps } from "../../src/commands/models/envList";
import { envResetCommand } from "../../src/commands/models/envReset";
import * as listTemplatesModule from "../../src/commands/models/listTemplates";
import { regeneratePluginCommand } from "../../src/commands/models/regeneratePlugin";
import { shareCommand } from "../../src/commands/models/share";
import { shareRemoveCommand } from "../../src/commands/models/shareRemove";
import { teamsappPackageCommand } from "../../src/commands/models/teamsapp/package";
import { teamsappPublishCommand } from "../../src/commands/models/teamsapp/publish";
import { teamsappUpdateCommand } from "../../src/commands/models/teamsapp/update";
import { teamsappValidateCommand } from "../../src/commands/models/teamsapp/validate";
import AzureTokenProvider from "../../src/commonlib/azureLogin";
import { logger } from "../../src/commonlib/logger";
import M365TokenProvider from "../../src/commonlib/M365TokenProviderWrapper";
import { MissingRequiredOptionError } from "../../src/error";
import * as utils from "../../src/utils";

describe("CLI commands", () => {
  const sandbox = sinon.createSandbox();

  const mockedEnvRestore: RestoreFn = () => {};

  process.env.TEAMSFX_CLI_BIN_NAME = "atk";
  beforeEach(() => {
    sandbox.stub(process.stdout, "write").returns(true as any);
    sandbox.stub(process.stderr, "write").returns(true as any);
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

    it("uses template alias and preset language in non-interactive mode", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      const createProjectStub = sandbox
        .stub(FxCore.prototype, "createProject")
        .resolves(ok({ projectPath: "..." }));
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(listTemplatesModule, "listAllTemplates").returns([
        {
          name: "api-plugin",
          alias: "api-plugin-from-scratch",
          displayName: "API Plugin",
          description: "desc",
          language: "typescript",
        },
      ] as any);

      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "new" },
        optionValues: {
          capabilities: "api-plugin-from-scratch",
          nonInteractive: true,
        },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };

      const res = await getCreateCommand().handler!(ctx);

      assert.isTrue(res.isOk());
      assert.isTrue(createProjectStub.calledOnce);
      const inputs = createProjectStub.firstCall.args[0] as any;
      assert.equal(inputs["template-name"], "api-plugin-from-scratch");
      assert.equal(inputs["programming-language"], "typescript");
    });

    it("keeps capability as template-name when template is not found", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      const createProjectStub = sandbox
        .stub(FxCore.prototype, "createProject")
        .resolves(ok({ projectPath: "..." }));
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(listTemplatesModule, "listAllTemplates").returns([] as any);

      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "new" },
        optionValues: {
          capabilities: "unknown-template",
          nonInteractive: true,
          "programming-language": "javascript",
        },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };

      const res = await getCreateCommand().handler!(ctx);

      assert.isTrue(res.isOk());
      const inputs = createProjectStub.firstCall.args[0] as any;
      assert.equal(inputs["template-name"], "unknown-template");
      assert.equal(inputs["programming-language"], "javascript");
    });

    it("includes template alias in capability choices", async () => {
      sandbox.stub(listTemplatesModule, "listAllTemplates").returns([
        {
          name: "api-plugin",
          alias: "api-plugin-from-scratch",
          displayName: "API Plugin",
          description: "desc",
          language: "typescript",
        },
      ] as any);

      const command = getCreateCommand();
      const capabilityOption = command.options?.find((o) => o.name === CliQuestionName.Capability);

      assert.include((capabilityOption as any)?.choices, "api-plugin-from-scratch");
    });

    it("with-plugin=yes and api-plugin-type matches a sub-template → uses subTemplate name", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      const createProjectStub = sandbox
        .stub(FxCore.prototype, "createProject")
        .resolves(ok({ projectPath: "..." }));
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(listTemplatesModule, "listAllTemplates").returns([
        {
          name: "declarative-agent",
          alias: "da",
          displayName: "Declarative Agent",
          description: "desc",
          language: "typescript",
        },
        {
          name: "declarative-agent-with-action-from-mcp",
          alias: "da-mcp",
          displayName: "DA+MCP",
          description: "desc",
          language: "typescript",
        },
      ] as any);

      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "new" },
        optionValues: {
          capabilities: "declarative-agent",
          "with-plugin": "yes",
          "api-plugin-type": "declarative-agent-with-action-from-mcp",
          nonInteractive: true,
        },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };

      const res = await getCreateCommand().handler!(ctx);
      assert.isTrue(res.isOk());
      const inputs = createProjectStub.firstCall.args[0] as any;
      assert.equal(inputs["template-name"], "declarative-agent-with-action-from-mcp");
    });

    it("with-plugin=yes and api-plugin-type=mcp falls back to actionTemplateMap", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      const createProjectStub = sandbox
        .stub(FxCore.prototype, "createProject")
        .resolves(ok({ projectPath: "..." }));
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      // Only parent template exists; 'mcp' action type is NOT in templates list
      sandbox.stub(listTemplatesModule, "listAllTemplates").returns([
        {
          name: "declarative-agent",
          alias: "da",
          displayName: "Declarative Agent",
          description: "desc",
          language: "typescript",
        },
      ] as any);

      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "new" },
        optionValues: {
          capabilities: "declarative-agent",
          "with-plugin": "yes",
          "api-plugin-type": "mcp",
          nonInteractive: true,
        },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };

      const res = await getCreateCommand().handler!(ctx);
      assert.isTrue(res.isOk());
      const inputs = createProjectStub.firstCall.args[0] as any;
      assert.equal(inputs["template-name"], "declarative-agent-with-action-from-mcp");
    });

    it("createProject result with warnings logs each warning", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      sandbox.stub(FxCore.prototype, "createProject").resolves(
        ok({
          projectPath: "...",
          warnings: [
            { type: "general", content: "warn1" },
            { type: "general", content: "warn2" },
          ],
        } as any)
      );
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      const warnStub = sandbox.stub(logger, "warning").resolves();

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
      assert.equal(warnStub.callCount, 2);
      assert.equal(warnStub.firstCall.args[0], "warn1");
      assert.equal(warnStub.secondCall.args[0], "warn2");
    });

    it("isTdpTemplate=true triggers createProjectFromTdp instead of createProject", async () => {
      sandbox.stub(activate, "getFxCore").returns(new FxCore({} as any));
      const createProjectFromTdpStub = sandbox
        .stub(FxCore.prototype, "createProjectFromTdp")
        .resolves(ok({ projectPath: "..." }));
      const createProjectStub = sandbox.stub(FxCore.prototype, "createProject");
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(listTemplatesModule, "listAllTemplates").returns([] as any);

      const ctx: CLIContext = {
        command: { ...getCreateCommand(), fullName: "new" },
        optionValues: {
          // Providing teamsAppFromTdp with a staticTab makes isTdpTemplate() return true
          teamsAppFromTdp: {
            teamsAppId: "test-app-id",
            staticTabs: [
              {
                objectId: "objId",
                entityId: "entityId",
                name: "tab",
                contentUrl: "https://example.com",
                websiteUrl: "https://example.com",
                scopes: [],
                context: [],
              },
            ],
          } as any,
          nonInteractive: true,
        } as any,
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };

      const res = await getCreateCommand().handler!(ctx);
      assert.isTrue(res.isOk());
      assert.isTrue(createProjectFromTdpStub.calledOnce);
      assert.isTrue(createProjectStub.notCalled);
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
        command: {
          ...listSamplesCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} list samples`,
        },
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
        command: {
          ...accountLoginAzureCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} auth login azure`,
        },
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
        command: {
          ...accountLoginAzureCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} auth login azure`,
        },
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
        command: {
          ...accountLoginAzureCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} auth login azure`,
        },
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
        command: {
          ...accountLoginM365Command,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} auth login m365`,
        },
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

  describe("importOpenPluginCommand", async () => {
    it("success", async () => {
      sandbox
        .stub(FxCore.prototype, "importOpenPlugin")
        .resolves(ok({ projectPath: "/tmp/imported", warnings: [] }));
      const ctx: CLIContext = {
        command: { ...importOpenPluginCommand, fullName: "import openplugin" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await importOpenPluginCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });

    it("logs warnings returned by importOpenPlugin", async () => {
      sandbox.stub(FxCore.prototype, "importOpenPlugin").resolves(
        ok({
          projectPath: "/tmp/imported",
          warnings: [{ type: "openPluginImport", content: "test warning" }],
        })
      );
      const ctx: CLIContext = {
        command: { ...importOpenPluginCommand, fullName: "import openplugin" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await importOpenPluginCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });

    it("propagates errors from importOpenPlugin", async () => {
      sandbox
        .stub(FxCore.prototype, "importOpenPlugin")
        .resolves(err(new SystemError("OpenPluginImport", "Boom", "boom")));
      const ctx: CLIContext = {
        command: { ...importOpenPluginCommand, fullName: "import openplugin" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await importOpenPluginCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });

  describe("exportOpenPluginCommand", async () => {
    it("success", async () => {
      sandbox
        .stub(FxCore.prototype, "exportOpenPlugin")
        .resolves(ok({ outputPath: "/tmp/exported", warnings: [] }));
      const ctx: CLIContext = {
        command: { ...exportOpenPluginCommand, fullName: "export openplugin" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await exportOpenPluginCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });

    it("logs warnings returned by exportOpenPlugin", async () => {
      sandbox.stub(FxCore.prototype, "exportOpenPlugin").resolves(
        ok({
          outputPath: "/tmp/exported",
          warnings: [{ type: "openPluginExport", content: "test warning" }],
        })
      );
      const ctx: CLIContext = {
        command: { ...exportOpenPluginCommand, fullName: "export openplugin" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await exportOpenPluginCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });

    it("propagates errors from exportOpenPlugin", async () => {
      sandbox
        .stub(FxCore.prototype, "exportOpenPlugin")
        .resolves(err(new SystemError("OpenPluginExport", "Boom", "boom")));
      const ctx: CLIContext = {
        command: { ...exportOpenPluginCommand, fullName: "export openplugin" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await exportOpenPluginCommand.handler!(ctx);
      assert.isTrue(res.isErr());
    });
  });

  describe("regeneratePlguinCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "regeneratePlugin").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...regeneratePluginCommand, fullName: "regenerate plugin" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await regeneratePluginCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });

  describe("addCapabilityCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "addKnowledge").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...addCapabilityCommand, fullName: "add capability" },
        optionValues: {},
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await addCapabilityCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });

  describe("getAddCommand", async () => {
    it("customize GPT is enabled", async () => {
      const commands = addCommand();
      assert.isTrue(commands.commands?.length === 4);
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
      sandbox.stub(envAddDeps, "isValidProjectV3").returns(true);
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
      sandbox.stub(envAddDeps, "isValidProjectV3").returns(false);
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
      sandbox.stub(envListDeps, "isValidProjectV3").returns(true);
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
      sandbox.stub(envListDeps, "isValidProjectV3").returns(false);
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
      sandbox.stub(envListDeps, "isValidProjectV3").returns(true);
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
        command: { ...envAddCommand, fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} env reset` },
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
        command: { ...envAddCommand, fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} env reset` },
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
    afterEach(() => {
      sandbox.restore();
    });

    it("success with agent option", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox
        .stub(FxCore.prototype, "grantPermission")
        .resolves(ok({ state: "OK" } as PermissionsResult));
      const ctx: CLIContext = {
        command: { ...permissionGrantCommand, fullName: "teamsfx" },
        optionValues: { agent: true, email: "email", env: "dev" },
        globalOptionValues: { interactive: false },
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionGrantCommand.handler!(ctx);
      assert.isTrue(res.isOk());
      const inputs = ctx.optionValues as PermissionGrantInputs;
      assert.deepEqual(inputs[QuestionNames.collaborationAppType], [
        CollaborationConstants.AgentOptionId,
      ]);
    });

    it("success with agent option in interactive mode", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox
        .stub(FxCore.prototype, "grantPermission")
        .resolves(ok({ state: "OK" } as PermissionsResult));
      const ctx: CLIContext = {
        command: { ...permissionGrantCommand, fullName: "teamsfx" },
        optionValues: { agent: true },
        globalOptionValues: { interactive: true },
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionGrantCommand.handler!(ctx);
      assert.isTrue(res.isOk());
      const inputs = ctx.optionValues as PermissionGrantInputs;
      assert.deepEqual(inputs[QuestionNames.collaborationAppType], [
        CollaborationConstants.AgentOptionId,
      ]);
    });

    it("missing manifest options with agent = false", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox
        .stub(FxCore.prototype, "grantPermission")
        .resolves(ok({ state: "OK" } as PermissionsResult));
      const ctx: CLIContext = {
        command: { ...permissionGrantCommand, fullName: "teamsfx" },
        optionValues: { env: "dev", email: "email", agent: false },
        globalOptionValues: { interactive: false },
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionGrantCommand.handler!(ctx);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.isTrue(res.error instanceof MissingRequiredOptionError);
      }
    });

    it("success interactive = false", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox
        .stub(FxCore.prototype, "grantPermission")
        .resolves(ok({ state: "OK" } as PermissionsResult));
      const ctx: CLIContext = {
        command: { ...permissionGrantCommand, fullName: "teamsfx" },
        optionValues: { "manifest-path": "abc" },
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
    afterEach(() => {
      sandbox.restore();
    });

    it("listCollaborator with agent option", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox
        .stub(FxCore.prototype, "listCollaborator")
        .resolves(ok({ state: "OK" } as ListCollaboratorResult));
      const ctx: CLIContext = {
        command: { ...permissionStatusCommand, fullName: "teamsfx" },
        optionValues: { all: true, agent: true },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionStatusCommand.handler!(ctx);
      assert.isTrue(res.isOk());
      const inputs = ctx.optionValues as PermissionListInputs;
      assert.deepEqual(inputs[QuestionNames.collaborationAppType], [
        CollaborationConstants.AgentOptionId,
      ]);
    });

    it("checkPermission with agent option", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox
        .stub(FxCore.prototype, "checkPermission")
        .resolves(ok({ state: "OK" } as CollaborationStateResult));
      const ctx: CLIContext = {
        command: { ...permissionStatusCommand, fullName: "teamsfx" },
        optionValues: { all: false, agent: true },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await permissionStatusCommand.handler!(ctx);
      assert.isTrue(res.isOk());
      const inputs = ctx.optionValues as PermissionListInputs;
      assert.deepEqual(inputs[QuestionNames.collaborationAppType], [
        CollaborationConstants.AgentOptionId,
      ]);
    });

    it("listCollaborator", async () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
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
  describe("shareCommand", async () => {
    it("success", async () => {
      sandbox.stub(FxCore.prototype, "shareApplication").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...shareCommand, fullName: "teamsfx" },
        optionValues: { env: "dev" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await shareCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
  });
  describe("shareRemoveCommand", async () => {
    it("share with owners", async () => {
      sandbox.stub(FxCore.prototype, "removeSharedAccess").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...shareRemoveCommand, fullName: "teamsfx" },
        optionValues: { env: "dev" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await shareRemoveCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("share with users", async () => {
      sandbox.stub(FxCore.prototype, "shareApplication").resolves(ok(undefined));
      const ctx: CLIContext = {
        command: { ...shareRemoveCommand, fullName: "teamsfx" },
        optionValues: { env: "dev", users: "test@example.com" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await shareRemoveCommand.handler!(ctx);
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
        command: {
          ...entraAppUpdateCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} entraapp update`,
        },
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
      sandbox.stub(PackageService.prototype, "sideLoading").resolves(["", "", ""]);
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
    it("should success with zip package with Personal scope", async () => {
      sandbox.stub(m365utils, "getTokenAndUpn").resolves(["token", "upn"]);
      sandbox.stub(PackageService.prototype, "sideLoading").resolves(["", "", ""]);
      const ctx: CLIContext = {
        command: { ...m365SideloadingCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa", "file-path": "./", scope: "Personal" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365SideloadingCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("should success with zip package with Shared scope", async () => {
      sandbox.stub(m365utils, "getTokenAndUpn").resolves(["token", "upn"]);
      sandbox.stub(PackageService.prototype, "sideLoading").resolves(["", "", "share link"]);
      const ctx: CLIContext = {
        command: { ...m365SideloadingCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa", "file-path": "./", scope: "Shared" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365SideloadingCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("should success with zip package with Shared scope - lower case", async () => {
      sandbox.stub(m365utils, "getTokenAndUpn").resolves(["token", "upn"]);
      sandbox.stub(PackageService.prototype, "sideLoading").resolves(["", "", "share link"]);
      const ctx: CLIContext = {
        command: { ...m365SideloadingCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa", "file-path": "./", scope: "shared" },
        globalOptionValues: {},
        argumentValues: [],
        telemetryProperties: {},
      };
      const res = await m365SideloadingCommand.handler!(ctx);
      assert.isTrue(res.isOk());
    });
    it("should success with zip package with unknown scope", async () => {
      sandbox.stub(m365utils, "getTokenAndUpn").resolves(["token", "upn"]);
      sandbox.stub(PackageService.prototype, "sideLoading").resolves(["", "", ""]);
      const ctx: CLIContext = {
        command: { ...m365SideloadingCommand, fullName: "teamsfx" },
        optionValues: { "manifest-id": "aaa", "file-path": "./", scope: "unknown" },
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
        command: {
          ...teamsappUpdateCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} update`,
        },
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
        command: {
          ...teamsappUpdateCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} update`,
        },
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
        command: {
          ...teamsappPackageCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} package`,
        },
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
        command: {
          ...teamsappValidateCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} validate`,
        },
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
        command: {
          ...teamsappValidateCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} validate`,
        },
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
        command: {
          ...teamsappPublishCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} publish`,
        },
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
        command: {
          ...teamsappPublishCommand,
          fullName: `${process.env.TEAMSFX_CLI_BIN_NAME} publish`,
        },
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
