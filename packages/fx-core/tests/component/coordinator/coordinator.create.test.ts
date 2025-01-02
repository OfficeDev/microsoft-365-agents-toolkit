import "mocha";

import { err, Inputs, ok, Platform, SystemError, UserError } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import { glob } from "glob";
import * as sinon from "sinon";
import { createContext, setTools } from "../../../src/common/globalVars";
import { coordinator } from "../../../src/component/coordinator";
import { developerPortalScaffoldUtils } from "../../../src/component/developerPortalScaffoldUtils";
import { AppDefinition } from "../../../src/component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import { SpecGenerator } from "../../../src/component/generator/apiSpec/generator";
import { Generator } from "../../../src/component/generator/generator";
import { OfficeAddinGeneratorNew } from "../../../src/component/generator/officeAddin/generator";
import { SPFxGeneratorNew } from "../../../src/component/generator/spfx/spfxGenerator";
import { DefaultTemplateGenerator } from "../../../src/component/generator/templates/templateGenerator";
import { TemplateNames } from "../../../src/component/generator/templates/templateNames";
import { FxCore } from "../../../src/core/FxCore";
import { InputValidationError, MissingRequiredInputError } from "../../../src/error/common";
import { CreateSampleProjectInputs } from "../../../src/question";
import {
  ApiAuthOptions,
  ApiPluginStartOptions,
  CapabilityOptions,
  CustomCopilotAssistantOptions,
  CustomCopilotRagOptions,
  MeArchitectureOptions,
  ProjectTypeOptions,
  QuestionNames,
  ScratchOptions,
} from "../../../src/question/constants";
import { validationUtils } from "../../../src/ui/validationUtils";
import { MockTools, randomAppName } from "../../core/utils";
import { MockedUserInteraction } from "../../plugins/solution/util";
import mockedEnv, { RestoreFn } from "mocked-env";
import { FeatureFlagName } from "../../../src/common/featureFlags";
import { manifestUtils } from "../../../src/component/driver/teamsApp/utils/ManifestUtils";

describe("coordinator create", () => {
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  let generator: sinon.SinonStub;
  setTools(tools);
  let mockedEnvRestore: RestoreFn;
  beforeEach(() => {
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(manifestUtils, "trimManifestShortName").resolves(ok(undefined));
    generator = sandbox
      .stub(DefaultTemplateGenerator.prototype, <any>"scaffolding")
      .resolves(ok(undefined));
  });
  afterEach(() => {
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
    sandbox.restore();
  });

  describe("createSampleProject", () => {
    it("create project from sample", async () => {
      sandbox.stub(Generator, "generateSample").resolves(ok(undefined));
      sandbox.stub(fs, "pathExists").resolves(false);
      const inputs: CreateSampleProjectInputs = {
        platform: Platform.CLI,
        folder: ".",
        samples: "hello-world-tab-with-backend",
      };
      const fxCore = new FxCore(tools);
      const res = await fxCore.createSampleProject(inputs);
      assert.isTrue(res.isOk());
    });
    it("create project from sample: todo-list-SPFx", async () => {
      sandbox.stub(Generator, "generateSample").resolves(ok(undefined));
      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(glob, "glob").resolves();
      sandbox.stub(fs, "readFile").resolves("test" as any);
      sandbox.stub(fs, "writeFile").resolves("");
      const inputs: CreateSampleProjectInputs = {
        platform: Platform.CLI,
        folder: ".",
        samples: "todo-list-SPFx",
      };
      const fxCore = new FxCore(tools);
      const res = await fxCore.createSampleProject(inputs);
      assert.isTrue(res.isOk());
    });
    it("fail to create project from sample", async () => {
      sandbox.stub(Generator, "generateSample").resolves(err(new UserError({})));
      sandbox.stub(fs, "pathExists").resolves(false);
      const inputs: CreateSampleProjectInputs = {
        platform: Platform.CLI,
        folder: ".",
        samples: "hello-world-tab-with-backend",
      };
      const fxCore = new FxCore(tools);
      const res = await fxCore.createSampleProject(inputs);
      assert.isTrue(res.isErr());
    });
    it("create project from sample rename folder", async () => {
      sandbox.stub(Generator, "generateSample").resolves(ok(undefined));
      sandbox
        .stub(fs, "pathExists")
        .onFirstCall()
        .resolves(true)
        .onSecondCall()
        .resolves(false)
        .onThirdCall()
        .resolves(false);
      sandbox
        .stub(fs, "readdir")
        .onFirstCall()
        .resolves(["abc"] as any)
        .onSecondCall()
        .resolves([]);
      const inputs: CreateSampleProjectInputs = {
        platform: Platform.CLI,
        folder: ".",
        samples: "hello-world-tab-with-backend",
      };
      const fxCore = new FxCore(tools);
      const res = await fxCore.createSampleProject(inputs);
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.isTrue(res.value.projectPath.endsWith("_1"));
      }
    });
    it("MissingRequiredInputError missing sample id", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        ignoreLockByUT: true,
        folder: ".",
        [QuestionNames.Scratch]: ScratchOptions.no().id,
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.isTrue(res.error instanceof MissingRequiredInputError);
      }
    });
  });

  describe("create from scratch", async () => {
    it("MissingRequiredInputError missing folder", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.isTrue(res.error instanceof MissingRequiredInputError);
      }
    });
    it("MissingRequiredInputError missing App name", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        ignoreLockByUT: true,
        folder: ".",
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.isTrue(res.error instanceof MissingRequiredInputError);
      }
    });
    it("MissingRequiredInputError invalid App name", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        ignoreLockByUT: true,
        folder: ".",
        "app-name": "__#$%___",
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.isTrue(res.error instanceof InputValidationError);
      }
    });
    it("fail to create SPFx project", async () => {
      sandbox.stub(SPFxGeneratorNew.prototype, "run").resolves(err(new UserError({})));
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.Capabilities]: CapabilityOptions.SPFxTab().id,
        [QuestionNames.ProgrammingLanguage]: "javascript",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFramework]: "none",
        [QuestionNames.SPFxWebpartName]: "test",
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isErr());
    });

    it("ensureTrackingId fails", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(SPFxGeneratorNew.prototype, "run").resolves(ok({}));
      sandbox.stub(coordinator, "ensureTrackingId").resolves(err(new UserError({})));
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.Capabilities]: CapabilityOptions.SPFxTab().id,
        [QuestionNames.ProgrammingLanguage]: "typescript",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFramework]: "none",
        [QuestionNames.SPFxWebpartName]: "test",
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isErr());
    });
    it("success", async () => {
      sandbox.stub(SPFxGeneratorNew.prototype, "run").resolves(ok({}));
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(coordinator, "ensureTrackingId").resolves(ok("mock-id"));
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.Capabilities]: CapabilityOptions.SPFxTab().id,
        [QuestionNames.ProgrammingLanguage]: "typescript",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFramework]: "none",
        [QuestionNames.SPFxWebpartName]: "test",
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isOk());
    });

    it("create project for app with tab features from Developer Portal", async () => {
      sandbox.stub(coordinator, "ensureTrackingId").resolves(ok("mock-id"));
      sandbox.stub(developerPortalScaffoldUtils, "updateFilesForTdp").resolves(ok(undefined));
      const appDefinition: AppDefinition = {
        teamsAppId: "mock-id",
        appId: "mock-id",
        staticTabs: [
          {
            name: "tab1",
            entityId: "tab1",
            contentUrl: "mock-contentUrl",
            websiteUrl: "mock-websiteUrl",
            context: [],
            scopes: [],
          },
        ],
      };
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "javascript",
        teamsAppFromTdp: appDefinition,
        [QuestionNames.ProjectType]: "tab-type",
        [QuestionNames.Capabilities]: CapabilityOptions.nonSsoTab().id,
        [QuestionNames.ReplaceWebsiteUrl]: ["tab1"],
        [QuestionNames.ReplaceContentUrl]: [],
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isOk());
      assert.equal(generator.args[0][1].templateName, TemplateNames.Tab);
    });
    it("create project for app with bot feature from Developer Portal with updating files failed", async () => {
      sandbox.stub(coordinator, "ensureTrackingId").resolves(ok("mock-id"));
      sandbox
        .stub(developerPortalScaffoldUtils, "updateFilesForTdp")
        .resolves(err(new UserError("coordinator", "error", "msg", "msg")));
      const appDefinition: AppDefinition = {
        teamsAppId: "mock-id",
        appId: "mock-id",
        bots: [
          {
            botId: "mock-bot-id",
            isNotificationOnly: false,
            needsChannelSelector: false,
            supportsCalling: false,
            supportsFiles: false,
            supportsVideo: false,
            scopes: [],
            teamCommands: [],
            groupChatCommands: [],
            personalCommands: [],
          },
        ],
      };

      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "javascript",
        [QuestionNames.ProjectType]: ProjectTypeOptions.bot().id,
        [QuestionNames.Capabilities]: CapabilityOptions.basicBot().id,
        [QuestionNames.ReplaceBotIds]: ["bot"],
        teamsAppFromTdp: appDefinition,
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "error");
      }
      assert.equal(generator.args[0][1].templateName, TemplateNames.DefaultBot);
    });
    it("create project for app with tab and bot features from Developer Portal", async () => {
      sandbox.stub(coordinator, "ensureTrackingId").resolves(ok("mock-id"));
      sandbox.stub(developerPortalScaffoldUtils, "updateFilesForTdp").resolves(ok(undefined));
      const appDefinition: AppDefinition = {
        teamsAppId: "mock-id",
        appId: "mock-id",
        staticTabs: [
          {
            name: "tab1",
            entityId: "tab1",
            contentUrl: "mock-contentUrl",
            websiteUrl: "mock-websiteUrl",
            context: [],
            scopes: [],
          },
        ],
        bots: [
          {
            botId: "mock-bot-id",
            isNotificationOnly: false,
            needsChannelSelector: false,
            supportsCalling: false,
            supportsFiles: false,
            supportsVideo: false,
            scopes: [],
            teamCommands: [],
            groupChatCommands: [],
            personalCommands: [],
          },
        ],
      };

      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "javascript",
        teamsAppFromTdp: appDefinition,
        [QuestionNames.ProjectType]: "tab-bot-type",
        [QuestionNames.Capabilities]: "TabNonSsoAndBot",
        [QuestionNames.ReplaceWebsiteUrl]: ["tab1"],
        [QuestionNames.ReplaceContentUrl]: [],
        [QuestionNames.ReplaceBotIds]: ["bot"],
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isOk());
      assert.isTrue(generator.calledOnce);
      assert.equal(generator.args[0][1].templateName, TemplateNames.TabAndDefaultBot);
    });
    it("create project for app with tab and message extension features from Developer Portal", async () => {
      sandbox.stub(coordinator, "ensureTrackingId").resolves(ok("mock-id"));
      sandbox.stub(developerPortalScaffoldUtils, "updateFilesForTdp").resolves(ok(undefined));
      const appDefinition: AppDefinition = {
        teamsAppId: "mock-id",
        appId: "mock-id",
        staticTabs: [
          {
            name: "tab1",
            entityId: "tab1",
            contentUrl: "mock-contentUrl",
            websiteUrl: "mock-websiteUrl",
            context: [],
            scopes: [],
          },
        ],
        messagingExtensions: [
          {
            botId: "mock-bot-id",
            canUpdateConfiguration: false,
            commands: [],
            messageHandlers: [],
          },
        ],
      };
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "javascript",
        teamsAppFromTdp: appDefinition,
        [QuestionNames.ProjectType]: "tab-bot-type",
        [QuestionNames.Capabilities]: "TabNonSsoAndBot",
        [QuestionNames.ReplaceWebsiteUrl]: ["tab1"],
        [QuestionNames.ReplaceContentUrl]: [],
        [QuestionNames.ReplaceBotIds]: ["messageExtension"],
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isOk());
      assert.isTrue(generator.calledOnce);
      assert.equal(generator.args[0][1].templateName, TemplateNames.TabAndDefaultBot);
    });
    it("create project for app with no features from Developer Portal - failed expecting inputs", async () => {
      sandbox.stub(coordinator, "ensureTrackingId").resolves(ok("mock-id"));
      sandbox.stub(developerPortalScaffoldUtils, "updateFilesForTdp").resolves(ok(undefined));
      const appDefinition: AppDefinition = {
        teamsAppId: "mock-id",
        appId: "mock-id",
        staticTabs: [],
      };

      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "javascript",
        teamsAppFromTdp: appDefinition,
      };
      const fxCore = new FxCore(tools);
      const res = await fxCore.createProject(inputs);
      assert.isTrue(res.isErr());
    });

    it("create project for app from Developer Portal - not overwrite already set project type and capability", async () => {
      sandbox.stub(coordinator, "ensureTrackingId").resolves(ok("mock-id"));
      sandbox.stub(developerPortalScaffoldUtils, "updateFilesForTdp").resolves(ok(undefined));
      const appDefinition: AppDefinition = {
        teamsAppId: "mock-id",
        appId: "mock-id",
      };

      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "javascript",
        teamsAppFromTdp: appDefinition,
        [QuestionNames.ReplaceWebsiteUrl]: ["tab1"],
        [QuestionNames.ReplaceContentUrl]: [],
        [QuestionNames.ProjectType]: ProjectTypeOptions.tab().id,
        [QuestionNames.Capabilities]: CapabilityOptions.nonSsoTab().id,
      };
      const fxCore = new FxCore(tools);
      const res = await fxCore.createProject(inputs);
      assert.isTrue(res.isOk());
      assert.equal(generator.args[0][1].templateName, TemplateNames.Tab);
    });

    it("create API ME (no auth) from new api sucessfully", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.ProjectType]: ProjectTypeOptions.me().id,
        [QuestionNames.Capabilities]: CapabilityOptions.m365SearchMe().id,
        [QuestionNames.MeArchitectureType]: MeArchitectureOptions.newApi().id,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.Scratch]: ScratchOptions.yes().id,
      };
      const res = await coordinator.create(v3ctx, inputs);
      assert.isTrue(res.isOk());
      assert.equal(generator.args[0][1].templateName, TemplateNames.CopilotPluginFromScratch);
    });

    it("create API ME (key auth) from new api sucessfully", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();

      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.ProjectType]: ProjectTypeOptions.me().id,
        [QuestionNames.Capabilities]: CapabilityOptions.m365SearchMe().id,
        [QuestionNames.MeArchitectureType]: MeArchitectureOptions.newApi().id,
        [QuestionNames.ApiAuth]: ApiAuthOptions.bearerToken().id,
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.Scratch]: ScratchOptions.yes().id,
      };
      const res = await coordinator.create(v3ctx, inputs);
      assert.isTrue(res.isOk());
      assert.equal(generator.args[0][1].templateName, TemplateNames.CopilotPluginFromScratchApiKey);
    });

    it("create API ME from existing api successfully", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();
      sandbox
        .stub(SpecGenerator.prototype, "run")
        .resolves(ok({ warnings: [{ type: "", content: "", data: {} } as any] }));
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.ProjectType]: ProjectTypeOptions.me().id,
        [QuestionNames.Capabilities]: CapabilityOptions.m365SearchMe().id,
        [QuestionNames.MeArchitectureType]: MeArchitectureOptions.apiSpec().id,
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.Scratch]: ScratchOptions.yes().id,
      };
      const res = await coordinator.create(v3ctx, inputs);
      assert.isTrue(res.isOk());
    });

    it("create non-sso tab earlier than .Net8", async () => {
      const inputs: Inputs = {
        platform: Platform.VS,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "csharp",
        [QuestionNames.SafeProjectName]: "safeprojectname",
        ["targetFramework"]: "net6.0",
        [QuestionNames.ProjectType]: ProjectTypeOptions.tab().id,
        [QuestionNames.Capabilities]: CapabilityOptions.nonSsoTab().id,
      };
      const fxCore = new FxCore(tools);
      const res = await fxCore.createProject(inputs);

      assert.isTrue(res.isOk());
      assert.equal(generator.args[0][1].templateName, TemplateNames.Tab);
    });

    it("create sso tab earlier than .Net8", async () => {
      const inputs: Inputs = {
        platform: Platform.VS,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "csharp",
        [QuestionNames.SafeProjectName]: "safeprojectname",
        ["targetFramework"]: "net6.0",
        [QuestionNames.ProjectType]: ProjectTypeOptions.tab().id,
        [QuestionNames.Capabilities]: CapabilityOptions.tab().id,
      };
      const fxCore = new FxCore(tools);
      const res = await fxCore.createProject(inputs);

      assert.isTrue(res.isOk());
      assert.equal(generator.args[0][1].templateName, TemplateNames.SsoTab);
    });

    it("create non-sso tab from .NET 8", async () => {
      const inputs: Inputs = {
        platform: Platform.VS,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "csharp",
        [QuestionNames.SafeProjectName]: "safeprojectname",
        ["targetFramework"]: "net8.0",
        [QuestionNames.ProjectType]: ProjectTypeOptions.tab().id,
        [QuestionNames.Capabilities]: CapabilityOptions.nonSsoTab().id,
      };
      const fxCore = new FxCore(tools);
      const res = await fxCore.createProject(inputs);

      assert.isTrue(res.isOk());
      assert.equal(generator.args[0][1].templateName, TemplateNames.TabSSR);
    });

    it("create sso tab from .NET 8", async () => {
      const inputs: Inputs = {
        platform: Platform.VS,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "csharp",
        [QuestionNames.SafeProjectName]: "safeprojectname",
        ["targetFramework"]: "net8.0",
        [QuestionNames.ProjectType]: ProjectTypeOptions.tab().id,
        [QuestionNames.Capabilities]: CapabilityOptions.tab().id,
      };
      const fxCore = new FxCore(tools);
      const res = await fxCore.createProject(inputs);

      assert.isTrue(res.isOk());
      assert.equal(generator.args[0][1].templateName, TemplateNames.SsoTabSSR);
    });

    it("create custom copilot rag custom api success", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "typescript",
        [QuestionNames.SafeProjectName]: "safeprojectname",
        [QuestionNames.ProjectType]: ProjectTypeOptions.customCopilot().id,
        [QuestionNames.Capabilities]: CapabilityOptions.customCopilotRag().id,
        [QuestionNames.CustomCopilotRag]: CustomCopilotRagOptions.customApi().id,
        [QuestionNames.ApiSpecLocation]: "spec",
        [QuestionNames.ApiOperation]: "test",
        [QuestionNames.LLMService]: "llm-service-openAI",
        [QuestionNames.OpenAIKey]: "mockedopenaikey",
      };
      sandbox.stub(SpecGenerator.prototype, "post").resolves(ok({}));
      sandbox.stub(validationUtils, "validateInputs").resolves(undefined);

      const fxCore = new FxCore(tools);
      const res = await fxCore.createProject(inputs);

      assert.isTrue(res.isOk());
      assert.equal(generator.args[0][1].templateName, TemplateNames.CustomCopilotRagCustomApi);
    });

    it("create custom copilot rag custom api with azure open ai success", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "typescript",
        [QuestionNames.SafeProjectName]: "safeprojectname",
        [QuestionNames.ProjectType]: ProjectTypeOptions.customCopilot().id,
        [QuestionNames.Capabilities]: CapabilityOptions.customCopilotRag().id,
        [QuestionNames.CustomCopilotRag]: CustomCopilotRagOptions.customApi().id,
        [QuestionNames.ApiSpecLocation]: "spec",
        [QuestionNames.ApiOperation]: "test",
        [QuestionNames.LLMService]: "llm-service-azure-openai",
        [QuestionNames.AzureOpenAIKey]: "mockedAzureOpenAIKey",
        [QuestionNames.AzureOpenAIEndpoint]: "mockedAzureOpenAIEndpoint",
        [QuestionNames.AzureOpenAIDeploymentName]: "mockedAzureOpenAIDeploymentName",
      };
      sandbox.stub(SpecGenerator.prototype, "post").resolves(ok({}));
      sandbox.stub(validationUtils, "validateInputs").resolves(undefined);

      const fxCore = new FxCore(tools);
      const res = await fxCore.createProject(inputs);

      assert.isTrue(res.isOk());
      assert.equal(generator.args[0][1].templateName, TemplateNames.CustomCopilotRagCustomApi);
    });

    it("create custom agent api with azure open ai success", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "typescript",
        [QuestionNames.SafeProjectName]: "safeprojectname",
        [QuestionNames.ProjectType]: ProjectTypeOptions.customCopilot().id,
        [QuestionNames.Capabilities]: CapabilityOptions.customCopilotAssistant().id,
        [QuestionNames.CustomCopilotAssistant]: CustomCopilotAssistantOptions.new().id,
        [QuestionNames.ApiSpecLocation]: "spec",
        [QuestionNames.ApiOperation]: "test",
        [QuestionNames.AzureOpenAIKey]: "mockedAzureOpenAIKey",
        [QuestionNames.AzureOpenAIEndpoint]: "mockedAzureOpenAIEndpoint",
        [QuestionNames.AzureOpenAIDeploymentName]: "mockedAzureOpenAIDeploymentName",
      };
      sandbox.stub(SpecGenerator.prototype, "post").resolves(ok({}));
      sandbox.stub(validationUtils, "validateInputs").resolves(undefined);

      const fxCore = new FxCore(tools);
      const res = await fxCore.createProject(inputs);

      assert.isTrue(res.isOk());
      assert.equal(generator.args[0][1].templateName, TemplateNames.CustomCopilotAssistantNew);
    });

    it("create custom copilot rag custom api failed", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "typescript",
        [QuestionNames.SafeProjectName]: "safeprojectname",
        [QuestionNames.ProjectType]: ProjectTypeOptions.customCopilot().id,
        [QuestionNames.Capabilities]: CapabilityOptions.customCopilotRag().id,
        [QuestionNames.CustomCopilotRag]: CustomCopilotRagOptions.customApi().id,
        [QuestionNames.ApiSpecLocation]: "spec",
        [QuestionNames.ApiOperation]: "test",
        [QuestionNames.LLMService]: "llm-service-openAI",
        [QuestionNames.OpenAIKey]: "mockedopenaikey",
      };
      sandbox
        .stub(SpecGenerator.prototype, "run")
        .resolves(err(new SystemError("test", "test", "test")));
      sandbox.stub(validationUtils, "validateInputs").resolves(undefined);

      const fxCore = new FxCore(tools);
      const res = await fxCore.createProject(inputs);

      assert.isTrue(res.isErr() && res.error.name === "test");
    });

    it("create API Plugin with No authentication (feature flag enabled)", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();

      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.ProjectType]: ProjectTypeOptions.Agent().id,
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.newApi().id,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.ProgrammingLanguage]: "javascript",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.Scratch]: ScratchOptions.yes().id,
      };
      const res = await coordinator.create(v3ctx, inputs);
      assert.isTrue(res.isOk());
    });

    it("create API Plugin with api-key auth (feature flag enabled)", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();

      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.ProjectType]: ProjectTypeOptions.Agent().id,
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.newApi().id,
        [QuestionNames.ApiAuth]: ApiAuthOptions.apiKey().id,
        [QuestionNames.ProgrammingLanguage]: "javascript",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.Scratch]: ScratchOptions.yes().id,
      };
      const res = await coordinator.create(v3ctx, inputs);
      assert.isTrue(res.isOk());
    });

    it("create API Plugin with OAuth (feature flag enabled)", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();

      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.ProjectType]: ProjectTypeOptions.Agent().id,
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.newApi().id,
        [QuestionNames.ApiAuth]: ApiAuthOptions.oauth().id,
        [QuestionNames.ProgrammingLanguage]: "javascript",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.Scratch]: ScratchOptions.yes().id,
      };
      const res = await coordinator.create(v3ctx, inputs);
      assert.isTrue(res.isOk());
    });

    it("should scaffold taskpane successfully", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();
      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(OfficeAddinGeneratorNew.prototype, "run").resolves(ok({}));
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.ProjectType]: ProjectTypeOptions.outlookAddin().id,
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.Scratch]: ScratchOptions.yes().id,
      };
      const res = await coordinator.create(v3ctx, inputs);
      assert.isTrue(res.isOk());
    });

    it("should scaffold from API spec successfully", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();

      sandbox
        .stub(SpecGenerator.prototype, "run")
        .resolves(ok({ warnings: [{ type: "", content: "", data: {} } as any] }));

      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.ProjectType]: ProjectTypeOptions.Agent().id,
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.Scratch]: ScratchOptions.yes().id,
      };
      const res = await coordinator.create(v3ctx, inputs);
      assert.isTrue(res.isOk());
    });

    it("scaffold from API spec error", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();

      sandbox
        .stub(SpecGenerator.prototype, "run")
        .resolves(err(new SystemError("mockedSource", "mockedError", "mockedMessage", "")));
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.ProjectType]: ProjectTypeOptions.Agent().id,
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.Scratch]: ScratchOptions.yes().id,
      };
      const res = await coordinator.create(v3ctx, inputs);
      assert.isTrue(res.isErr());
    });

    it("success for kiota integration: plugin", async () => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.KiotaIntegration]: "true",
      });
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(coordinator, "ensureTrackingId").resolves(ok("mock-id"));
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.ProjectType]: ProjectTypeOptions.Agent().id,
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.isNotNull(res.value.lastCommand);
        assert.equal(res.value.projectPath, "");
      }
    });

    it("success for kiota integration: declarative copilot", async () => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.KiotaIntegration]: "true",
      });
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(coordinator, "ensureTrackingId").resolves(ok("mock-id"));
      const inputs: Inputs = {
        platform: Platform.VSCode,
        [QuestionNames.ProjectType]: ProjectTypeOptions.Agent().id,
        [QuestionNames.Capabilities]: CapabilityOptions.declarativeAgent().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.WithPlugin]: "yes",
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.isNotNull(res.value.lastCommand);
        assert.equal(res.value.projectPath, "");
      }
    });
  });
});
