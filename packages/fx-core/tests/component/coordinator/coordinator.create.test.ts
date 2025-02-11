import "mocha";

import { err, Inputs, ok, Platform, SystemError, UserError } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import { glob } from "glob";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as sinon from "sinon";
import { FeatureFlagName } from "../../../src/common/featureFlags";
import { createContext, setTools } from "../../../src/common/globalVars";
import { coordinator } from "../../../src/component/coordinator";
import { developerPortalScaffoldUtils } from "../../../src/component/developerPortalScaffoldUtils";
import { AppDefinition } from "../../../src/component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import { manifestUtils } from "../../../src/component/driver/teamsApp/utils/ManifestUtils";
import { SpecGenerator } from "../../../src/component/generator/apiSpec/generator";
import { DefaultTemplateGenerator } from "../../../src/component/generator/defaultGenerator";
import { Generator } from "../../../src/component/generator/generator";
import { OfficeAddinGeneratorNew } from "../../../src/component/generator/officeAddin/generator";
import { SPFxGeneratorNew } from "../../../src/component/generator/spfx/spfxGenerator";
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
import { TdpGenerator } from "../../../src/component/generator/other/tdpGenerator";
import { SsrTabGenerator } from "../../../src/component/generator/other/ssrTabGenerator";
import { CopilotExtensionGenerator } from "../../../src/component/generator/copilotExtension/generator";

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
      sandbox.stub(TdpGenerator.prototype, "run").resolves(ok({}));
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
        [QuestionNames.TemplateName]: TemplateNames.Tab,
        [QuestionNames.ReplaceWebsiteUrl]: ["tab1"],
        [QuestionNames.ReplaceContentUrl]: [],
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isOk());
    });
    it("create project for app with bot feature from Developer Portal with updating files failed", async () => {
      sandbox.stub(coordinator, "ensureTrackingId").resolves(ok("mock-id"));
      sandbox
        .stub(TdpGenerator.prototype, "run")
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
        [QuestionNames.TemplateName]: TemplateNames.DefaultBot,
        [QuestionNames.ReplaceBotIds]: ["bot"],
        teamsAppFromTdp: appDefinition,
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "error");
      }
    });
    it("create project for app with tab and bot features from Developer Portal", async () => {
      sandbox.stub(coordinator, "ensureTrackingId").resolves(ok("mock-id"));
      sandbox.stub(TdpGenerator.prototype, "run").resolves(ok({}));
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
        [QuestionNames.TemplateName]: TemplateNames.TabAndDefaultBot,
        [QuestionNames.ReplaceWebsiteUrl]: ["tab1"],
        [QuestionNames.ReplaceContentUrl]: [],
        [QuestionNames.ReplaceBotIds]: ["bot"],
      };
      const context = createContext();
      const res = await coordinator.create(context, inputs);
      assert.isTrue(res.isOk());
    });

    it("create non-sso tab from .NET 8", async () => {
      sandbox.stub(SsrTabGenerator.prototype, "run").resolves(ok({}));
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();
      const inputs: Inputs = {
        platform: Platform.VS,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "csharp",
        [QuestionNames.SafeProjectName]: "safeprojectname",
        ["targetFramework"]: "net8.0",
        [QuestionNames.TemplateName]: TemplateNames.TabSSR,
      };
      const res = await coordinator.create(v3ctx, inputs);

      assert.isTrue(res.isOk());
    });

    it("create sso tab from .NET 8", async () => {
      const v3ctx = createContext();
      sandbox.stub(SsrTabGenerator.prototype, "run").resolves(ok({}));
      v3ctx.userInteraction = new MockedUserInteraction();
      const inputs: Inputs = {
        platform: Platform.VS,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "csharp",
        [QuestionNames.SafeProjectName]: "safeprojectname",
        ["targetFramework"]: "net8.0",
        [QuestionNames.TemplateName]: TemplateNames.SsoTabSSR,
      };
      const res = await coordinator.create(v3ctx, inputs);

      assert.isTrue(res.isOk());
    });

    it("create custom copilot rag custom api success", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "typescript",
        [QuestionNames.SafeProjectName]: "safeprojectname",
        [QuestionNames.TemplateName]: TemplateNames.CustomCopilotRagCustomApi,
        [QuestionNames.CustomCopilotRag]: CustomCopilotRagOptions.customApi().id,
        [QuestionNames.ApiSpecLocation]: "spec",
        [QuestionNames.ApiOperation]: "test",
        [QuestionNames.LLMService]: "llm-service-openAI",
        [QuestionNames.OpenAIKey]: "mockedopenaikey",
      };
      sandbox.stub(SpecGenerator.prototype, "run").resolves(ok({}));
      sandbox.stub(validationUtils, "validateInputs").resolves(undefined);

      const res = await coordinator.create(v3ctx, inputs);

      assert.isTrue(res.isOk());
    });

    it("create custom copilot rag custom api with azure open ai success", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "typescript",
        [QuestionNames.SafeProjectName]: "safeprojectname",
        [QuestionNames.TemplateName]: TemplateNames.CustomCopilotRagCustomApi,
        [QuestionNames.CustomCopilotRag]: CustomCopilotRagOptions.customApi().id,
        [QuestionNames.ApiSpecLocation]: "spec",
        [QuestionNames.ApiOperation]: "test",
        [QuestionNames.LLMService]: "llm-service-azure-openai",
        [QuestionNames.AzureOpenAIKey]: "mockedAzureOpenAIKey",
        [QuestionNames.AzureOpenAIEndpoint]: "mockedAzureOpenAIEndpoint",
        [QuestionNames.AzureOpenAIDeploymentName]: "mockedAzureOpenAIDeploymentName",
      };
      sandbox.stub(SpecGenerator.prototype, "run").resolves(ok({}));
      sandbox.stub(validationUtils, "validateInputs").resolves(undefined);

      const res = await coordinator.create(v3ctx, inputs);

      assert.isTrue(res.isOk());
    });

    it("create custom agent api with azure open ai success", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "typescript",
        [QuestionNames.SafeProjectName]: "safeprojectname",
        [QuestionNames.TemplateName]: TemplateNames.CustomCopilotAssistantNew,
        [QuestionNames.CustomCopilotAssistant]: CustomCopilotAssistantOptions.new().id,
        [QuestionNames.ApiSpecLocation]: "spec",
        [QuestionNames.ApiOperation]: "test",
        [QuestionNames.AzureOpenAIKey]: "mockedAzureOpenAIKey",
        [QuestionNames.AzureOpenAIEndpoint]: "mockedAzureOpenAIEndpoint",
        [QuestionNames.AzureOpenAIDeploymentName]: "mockedAzureOpenAIDeploymentName",
      };
      sandbox.stub(DefaultTemplateGenerator.prototype, "run").resolves(ok({}));
      sandbox.stub(validationUtils, "validateInputs").resolves(undefined);

      const res = await coordinator.create(v3ctx, inputs);

      assert.isTrue(res.isOk());
    });

    it("create custom copilot rag custom api failed", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();
      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.ProgrammingLanguage]: "typescript",
        [QuestionNames.SafeProjectName]: "safeprojectname",
        [QuestionNames.TemplateName]: TemplateNames.CustomCopilotRagCustomApi,
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

      const res = await coordinator.create(v3ctx, inputs);

      assert.isTrue(res.isErr() && res.error.name === "test");
    });

    it("create API Plugin with No authentication (feature flag enabled)", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();
      sandbox.stub(CopilotExtensionGenerator.prototype, "run").resolves(ok({}));
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
        [QuestionNames.TemplateName]: TemplateNames.ApiPluginFromScratch,
      };
      const res = await coordinator.create(v3ctx, inputs);
      assert.isTrue(res.isOk());
    });

    it("create API Plugin with api-key auth (feature flag enabled)", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();
      sandbox.stub(CopilotExtensionGenerator.prototype, "run").resolves(ok({}));
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
        [QuestionNames.TemplateName]: TemplateNames.ApiPluginFromScratchBearer,
      };
      const res = await coordinator.create(v3ctx, inputs);
      assert.isTrue(res.isOk());
    });

    it("create API Plugin with OAuth (feature flag enabled)", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();
      sandbox.stub(CopilotExtensionGenerator.prototype, "run").resolves(ok({}));
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
        [QuestionNames.TemplateName]: TemplateNames.ApiPluginFromScratchOAuth,
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
        [QuestionNames.TemplateName]: TemplateNames.OutlookTaskpane,
      };
      const res = await coordinator.create(v3ctx, inputs);
      assert.isTrue(res.isOk());
    });

    it("should scaffold from API spec successfully", async () => {
      const v3ctx = createContext();
      v3ctx.userInteraction = new MockedUserInteraction();

      sandbox
        .stub(DefaultTemplateGenerator.prototype, "run")
        .resolves(ok({ warnings: [{ type: "", content: "", data: {} } as any] }));

      const inputs: Inputs = {
        platform: Platform.VSCode,
        folder: ".",
        [QuestionNames.ProjectType]: ProjectTypeOptions.Agent().id,
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.AppName]: randomAppName(),
        [QuestionNames.Scratch]: ScratchOptions.yes().id,
        [QuestionNames.TemplateName]: TemplateNames.ApiPluginFromScratch,
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
