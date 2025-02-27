import { ProjectType, SpecParser, ValidationStatus } from "@microsoft/m365-spec-parser";
import { ApiOperation, Inputs, Platform, SystemError } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import { err, ok } from "neverthrow";
import * as sinon from "sinon";
import { FeatureFlagName } from "../../../../src/common/featureFlags";
import { createContext, setTools } from "../../../../src/common/globalVars";
import { copilotGptManifestUtils } from "../../../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import { manifestUtils } from "../../../../src/component/driver/teamsApp/utils/ManifestUtils";
import { DeclarativeAgentWithExistingApiSpecGenerator } from "../../../../src/component/generator/openApiSpec/declarativeAgentGenerator";
import * as helper from "../../../../src/component/generator/openApiSpec/helper";
import { TemplateNames } from "../../../../src/component/generator/templates/templateNames";
import {
  ApiPluginStartOptions,
  CapabilityOptions,
  ProgrammingLanguage,
  QuestionNames,
} from "../../../../src/question";
import { MockTools } from "../../../core/utils";
import { teamsManifest } from "./fakeData";

const tools = new MockTools();

describe("DeclarativeAgentWithExistingApiSpecGenerator", async () => {
  const sandbox = sinon.createSandbox();
  before(() => {
    setTools(tools);
  });
  after(() => {
    sandbox.restore();
  });
  describe("activate", async () => {
    it("should activate and get correct template name", async () => {
      const generator = new DeclarativeAgentWithExistingApiSpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.TemplateName]: TemplateNames.ApiPluginWithExistingApiSpec,
      };
      const res = await generator.activate(context, inputs);
      assert.isTrue(res);
    });
  });

  describe("getTemplateInfos", async () => {
    const sandbox = sinon.createSandbox();
    let mockedEnvRestore: RestoreFn | undefined;
    afterEach(async () => {
      sandbox.restore();
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });

    it("happy path", async () => {
      const generator = new DeclarativeAgentWithExistingApiSpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.TemplateName]: TemplateNames.ApiPluginWithExistingApiSpec,
        [QuestionNames.AppName]: "testapp",
      };
      inputs[QuestionNames.ApiSpecLocation] = "test.yaml";
      inputs.apiAuthData = [
        { serverUrl: "https://test.com", authName: "test", authType: "apiKey" },
      ];
      const res = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.length, 1);
        assert.equal(res.value[0].templateName, TemplateNames.ApiPluginWithExistingApiSpec);
        assert.equal(res.value[0].replaceMap?.["DeclarativeCopilot"], "true");
      }
    });

    it("succeed even get yaml file failed", async () => {
      const generator = new DeclarativeAgentWithExistingApiSpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.AppName]: "testapp",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
        [QuestionNames.TemplateName]: TemplateNames.ApiPluginWithExistingApiSpec,
      };
      inputs[QuestionNames.ApiSpecLocation] = "test.yaml";
      inputs.apiAuthData = [
        { serverUrl: "https://test.com", authName: "test", authType: "apiKey" },
      ];
      sandbox.stub(JSON, "parse").throws();
      const res = await generator.getTemplateInfos(context, inputs, ".", { telemetryProps: {} });
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.length, 1);
        assert.equal(res.value[0].templateName, TemplateNames.ApiPluginWithExistingApiSpec);
        assert.equal(res.value[0].language, ProgrammingLanguage.CSharp);
      }
    });

    it("happy path for kiota integration with auth", async () => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.KiotaIntegration]: "true",
      });
      const generator = new DeclarativeAgentWithExistingApiSpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.AppName]: "testapp",
        [QuestionNames.ApiPluginManifestPath]: "ai-plugin.json",
        [QuestionNames.TemplateName]: TemplateNames.ApiPluginWithExistingApiSpec,
      };
      inputs[QuestionNames.ApiSpecLocation] = "test.yaml";
      sandbox.stub(helper, "parseAndUpdatePluginManifestForKiota").resolves([
        {
          authName: "mockedAuthName",
          authType: "apiKey",
          registrationId: "MOCKED_REGISTRATION_ID",
        },
      ]);
      const res = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.length, 1);
        assert.equal(res.value[0].templateName, "api-plugin-existing-api");
        assert.equal(res.value[0].replaceMap?.["DeclarativeCopilot"], "true");
      }
    });

    it("happy path for kiota integration without auth", async () => {
      mockedEnvRestore = mockedEnv({
        [FeatureFlagName.KiotaIntegration]: "true",
      });
      const generator = new DeclarativeAgentWithExistingApiSpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.AppName]: "testapp",
        [QuestionNames.ApiPluginManifestPath]: "ai-plugin.json",
        [QuestionNames.TemplateName]: TemplateNames.ApiPluginWithExistingApiSpec,
      };
      inputs[QuestionNames.ApiSpecLocation] = "test.yaml";
      sandbox.stub(helper, "parseAndUpdatePluginManifestForKiota").resolves([]);
      const res = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.length, 1);
        assert.equal(res.value[0].templateName, "api-plugin-existing-api");
        assert.equal(res.value[0].replaceMap?.["DeclarativeCopilot"], "true");
      }
    });
  });

  describe("post", function () {
    const sandbox = sinon.createSandbox();
    let mockedEnvRestore: RestoreFn | undefined;

    const apiOperations: ApiOperation[] = [
      {
        id: "operation1",
        label: "operation1",
        groupName: "1",
        data: {
          serverUrl: "https://server1",
        },
      },
      {
        id: "operation2",
        label: "operation2",
        groupName: "1",
        data: {
          serverUrl: "https://server1",
          authName: "auth",
        },
      },
    ];

    before(() => {
      setTools(tools);
    });

    afterEach(async () => {
      sandbox.restore();
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });

    it("success", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.ApiSpecLocation]: "https://test.com",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        templateState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generateForCopilot")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(copilotGptManifestUtils, "updateDeclarativeAgentManifest").resolves(ok(""));
      sandbox.stub(helper, "generateScaffoldingSummary").resolves("");

      const generator = new DeclarativeAgentWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
    });

    it("generate for oauth: success", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.AppName]: "test",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        supportedApisFromApiSpec: [
          {
            id: "operation1",
            label: "operation1",
            groupName: "1",
            data: {
              serverUrl: "https://server1",
              authName: "auth",
              authType: "oauth2",
            },
          },
        ] as ApiOperation[],
        templateState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();

      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      sandbox.stub(copilotGptManifestUtils, "updateDeclarativeAgentManifest").resolves(ok(""));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generateForCopilot")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(helper, "generateScaffoldingSummary").resolves("");

      const generator = new DeclarativeAgentWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");
      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
    });

    it("declarative copilot with plugin success", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.Capabilities]: CapabilityOptions.declarativeAgent().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.ApiSpecLocation]: "https://test.com",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        templateState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const addAction = sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generateForCopilot")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(helper, "generateScaffoldingSummary").resolves("");

      const generator = new DeclarativeAgentWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
      assert.isTrue(addAction.calledOnce);
    });

    it("declarative copilot with plugin error", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.Capabilities]: CapabilityOptions.declarativeAgent().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.ApiSpecLocation]: "https://test.com",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        templateState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const addAction = sandbox
        .stub(copilotGptManifestUtils, "addAction")
        .resolves(err(new SystemError("test", "test", "test", "test")));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generateForCopilot")
        .resolves({ allSuccess: true, warnings: [] });

      const generator = new DeclarativeAgentWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isErr() && result.error.name === "test");
      assert.isTrue(generateBasedOnSpec.calledOnce);
      assert.isTrue(addAction.calledOnce);
    });

    it("generate for kiota", async function () {
      mockedEnvRestore = mockedEnv({ [FeatureFlagName.KiotaIntegration]: "true" });
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.AppName]: "test",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.ApiPluginManifestPath]: "test.json",
        [QuestionNames.ProjectType]: "copilot-agent-type",
        templateState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "list").resolves({
        APIs: [
          {
            api: "api1",
            server: "https://test",
            operationId: "get",
            auth: {
              name: "test",
              authScheme: {
                type: "http",
                scheme: "bearer",
              },
            },
            isValid: true,
            reason: [],
          },
        ],
        allAPICount: 1,
        validAPICount: 1,
      });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      sandbox.stub(fs, "pathExists").onFirstCall().resolves(true).onSecondCall().resolves(false);
      sandbox.stub(fs, "copyFile").resolves();
      sandbox.stub(fs, "readJSON").resolves({});
      sandbox.stub(SpecParser.prototype, "generateAdaptiveCardInPlugin").resolves();
      sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));

      const copyKiotaFolder = sandbox.stub(fs, "copy").callsFake((src, dest, options) => {
        assert.isTrue(src.endsWith(".kiota"));
        assert.isTrue(dest.endsWith(".kiota"));
      });
      sandbox.stub(helper, "generateScaffoldingSummary").resolves("");

      const generator = new DeclarativeAgentWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");
      assert.isTrue(result.isOk());
      assert.isTrue(copyKiotaFolder.calledOnce);
    });

    it("generate for kiota without .kiota folder", async function () {
      mockedEnvRestore = mockedEnv({ [FeatureFlagName.KiotaIntegration]: "true" });
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.AppName]: "test",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.ApiPluginManifestPath]: "test.json",
        [QuestionNames.ProjectType]: "copilot-agent-type",
        templateState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "list").resolves({
        APIs: [
          {
            api: "api1",
            server: "https://test",
            operationId: "get",
            auth: {
              name: "test",
              authScheme: {
                type: "http",
                scheme: "bearer",
              },
            },
            isValid: true,
            reason: [],
          },
        ],
        allAPICount: 1,
        validAPICount: 1,
      });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      sandbox.stub(fs, "pathExists").onFirstCall().resolves(false).onSecondCall().resolves(false);
      sandbox.stub(fs, "copyFile").resolves();
      sandbox.stub(fs, "readJSON").resolves({});
      sandbox.stub(SpecParser.prototype, "generateAdaptiveCardInPlugin").resolves();
      sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));

      const copyKiotaFolder = sandbox.stub(fs, "copy").resolves();
      sandbox.stub(helper, "generateScaffoldingSummary").resolves("");

      const generator = new DeclarativeAgentWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");
      assert.isTrue(result.isOk());
      assert.isTrue(copyKiotaFolder.notCalled);
    });

    it("generate for kiota with .kiota folder already exists", async function () {
      mockedEnvRestore = mockedEnv({ [FeatureFlagName.KiotaIntegration]: "true" });
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.AppName]: "test",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.ApiPluginManifestPath]: "test.json",
        [QuestionNames.ProjectType]: "copilot-agent-type",
        templateState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "list").resolves({
        APIs: [
          {
            api: "api1",
            server: "https://test",
            operationId: "get",
            auth: {
              name: "test",
              authScheme: {
                type: "http",
                scheme: "bearer",
              },
            },
            isValid: true,
            reason: [],
          },
        ],
        allAPICount: 1,
        validAPICount: 1,
      });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      sandbox.stub(fs, "pathExists").onFirstCall().resolves(true).onSecondCall().resolves(true);
      sandbox.stub(fs, "copyFile").resolves();
      sandbox.stub(fs, "readJSON").resolves({});
      sandbox.stub(SpecParser.prototype, "generateAdaptiveCardInPlugin").resolves();
      sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));

      const copyKiotaFolder = sandbox.stub(fs, "copy").resolves();
      sandbox.stub(helper, "generateScaffoldingSummary").resolves("");

      const generator = new DeclarativeAgentWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");
      assert.isTrue(result.isOk());
      assert.isTrue(copyKiotaFolder.notCalled);
    });

    it("generate for kiota without adaptive card", async function () {
      mockedEnvRestore = mockedEnv({ [FeatureFlagName.KiotaIntegration]: "true" });
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.AppName]: "test",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.apiSpec().id,
        [QuestionNames.ApiPluginManifestPath]: "test.json",
        [QuestionNames.ProjectType]: "copilot-agent-type",
        templateState: {
          templateName: "api-plugin-existing-api",
          isPlugin: true,
          uri: "https://test.com",
          isYaml: true,
          type: ProjectType.Copilot,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "list").resolves({
        APIs: [
          {
            api: "api1",
            server: "https://test",
            operationId: "get",
            auth: {
              name: "test",
              authScheme: {
                type: "http",
                scheme: "bearer",
              },
            },
            isValid: true,
            reason: [],
          },
        ],
        allAPICount: 1,
        validAPICount: 1,
      });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      sandbox.stub(fs, "pathExists").onFirstCall().resolves(true).onSecondCall().resolves(false);
      sandbox.stub(fs, "copyFile").resolves();
      sandbox.stub(fs, "readJSON").resolves({});
      sandbox
        .stub(SpecParser.prototype, "generateAdaptiveCardInPlugin")
        .throws(new Error("generate ac failed"));
      sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));

      const copyKiotaFolder = sandbox.stub(fs, "copy").resolves();
      sandbox.stub(helper, "generateScaffoldingSummary").resolves("");

      const generator = new DeclarativeAgentWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");
      assert.isTrue(result.isOk());
      assert.isTrue(copyKiotaFolder.calledOnce);
    });
  });
});
