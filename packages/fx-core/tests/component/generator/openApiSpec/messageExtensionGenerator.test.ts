import {
  ErrorType,
  ProjectType,
  SpecParser,
  SpecParserError,
  ValidationStatus,
  WarningType,
} from "@microsoft/m365-spec-parser";
import {
  ApiOperation,
  err,
  Inputs,
  ok,
  Platform,
  ResponseTemplatesFolderName,
  SystemError,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import { RestoreFn } from "mocked-env";
import * as sinon from "sinon";
import { createContext, setTools } from "../../../../src/common/globalVars";
import { manifestUtils } from "../../../../src/component/driver/teamsApp/utils/ManifestUtils";
import * as helper from "../../../../src/component/generator/openApiSpec/helper";
import { MessageExtensionWithExistingApiSpecGenerator } from "../../../../src/component/generator/openApiSpec/messageExtensionGenerator";
import { TemplateNames } from "../../../../src/component/generator/templates/templateNames";
import { ActionStartOptions, ProgrammingLanguage, QuestionNames } from "../../../../src/question";
import { MockTools } from "../../../core/utils";
import { teamsManifest } from "./fakeData";

const tools = new MockTools();

describe("MessageExtensionWithExistingApiSpecGenerator", async () => {
  const sandbox = sinon.createSandbox();
  before(() => {
    setTools(tools);
  });
  after(() => {
    sandbox.restore();
  });
  describe("activate", async () => {
    it("should activate and get correct template name", async () => {
      const generator = new MessageExtensionWithExistingApiSpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.TemplateName]: TemplateNames.MessageExtensionWithExistingApiSpec,
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
      const generator = new MessageExtensionWithExistingApiSpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: ActionStartOptions.apiSpec().id,
        [QuestionNames.TemplateName]: TemplateNames.MessageExtensionWithExistingApiSpec,
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
        assert.equal(res.value[0].templateName, TemplateNames.MessageExtensionWithExistingApiSpec);
        assert.equal(res.value[0].replaceMap?.["DeclarativeCopilot"], "");
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
        [QuestionNames.ApiSpecLocation]: "https://test.com",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        templateState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generate")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(helper, "generateScaffoldingSummary").resolves("");

      const generator = new MessageExtensionWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath", {
        telemetryProps: {
          "project-id": "test",
        },
      });

      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
    });

    it("success with API Key authentication", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.AppName]: "test",
        [QuestionNames.ApiSpecLocation]: "test.json",
        [QuestionNames.ApiOperation]: ["operation2"],
        supportedApisFromApiSpec: apiOperations,
        apiAuthData: [
          {
            authType: "apiKey",
            serverUrl: "",
          },
        ],
        templateState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generate")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(helper, "generateScaffoldingSummary").resolves("");

      const generator = new MessageExtensionWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
    });

    it("success with api spec warning and generate warnings", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
        [QuestionNames.ApiSpecLocation]: "test.json",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        templateState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "validate").resolves({
        status: ValidationStatus.Warning,
        errors: [],
        warnings: [
          {
            type: WarningType.OperationIdMissing,
            content: "warning",
            data: ["operation1", " operation2"],
          },
          {
            type: WarningType.ConvertSwaggerToOpenAPI,
            content: "",
          },
        ],
        specHash: "xxx",
      });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok({ ...teamsManifest }));
      const generateParser = sandbox.stub(SpecParser.prototype, "generate").resolves({
        allSuccess: true,
        warnings: [
          { type: WarningType.GenerateCardFailed, content: "test", data: "getPets" },
          {
            type: WarningType.OperationOnlyContainsOptionalParam,
            content: "test",
            data: "getPets",
          },
        ],
      });
      sandbox.stub(helper, "generateScaffoldingSummary").resolves("warning message");

      const generator = new MessageExtensionWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      if (result.isOk()) {
        assert.isTrue(result.value.warnings!.length === 4);
        assert.isFalse(result.value.warnings![0].content.includes("operation2"));
        assert.isUndefined(result.value.warnings![0].data);
        assert.equal(result.value.warnings![1].type, WarningType.ConvertSwaggerToOpenAPI);
        assert.equal(result.value.warnings![2].type, WarningType.GenerateCardFailed);
        assert.equal(
          result.value.warnings![3].type,
          WarningType.OperationOnlyContainsOptionalParam
        );
        assert.equal(result.value.warnings![3].content, "");
        assert.isTrue(generateParser.args[0][3]?.includes(ResponseTemplatesFolderName));
      }
    });

    it("success without api spec warning after filtering", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
        [QuestionNames.ApiSpecLocation]: "https://test.com",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        templateState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "validate").resolves({
        status: ValidationStatus.Warning,
        errors: [],
        warnings: [
          { type: WarningType.OperationIdMissing, content: "warning", data: ["operation2"] },
        ],
        specHash: "xxx",
      });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok({ ...teamsManifest }));
      sandbox.stub(SpecParser.prototype, "generate").resolves({ allSuccess: true, warnings: [] });

      const generator = new MessageExtensionWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      if (result.isOk()) {
        assert.isTrue(result.value.warnings!.length === 0);
      }
    });

    it("success with warnings when CSharp", async function () {
      const inputs: Inputs = {
        platform: Platform.VS,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        templateState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "validate").resolves({
        status: ValidationStatus.Warning,
        errors: [],
        warnings: [{ type: WarningType.OperationIdMissing, content: "warning" }],
        specHash: "xxx",
      });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox
        .stub(manifestUtils, "_readAppManifest")
        .resolves(ok({ ...teamsManifest, name: { short: "", full: "" } }));
      sandbox.stub(SpecParser.prototype, "generate").resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(helper, "generateScaffoldingSummary").resolves("warn message");

      const generator = new MessageExtensionWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
    });

    it("invalid API spec", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        templateState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "validate").resolves({
        status: ValidationStatus.Error,
        errors: [{ type: ErrorType.NoServerInformation, content: "" }],
        warnings: [],
        specHash: "xxx",
      });

      sandbox.stub(SpecParser.prototype, "generate").resolves();

      const generator = new MessageExtensionWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.isTrue(result.error.name === "invalid-api-spec");
      }
    });

    it("read manifest error", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        templateState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox
        .stub(manifestUtils, "_readAppManifest")
        .resolves(err(new SystemError("readManifest", "name", "", "")));
      sandbox.stub(SpecParser.prototype, "generate").resolves({ allSuccess: true, warnings: [] });

      const generator = new MessageExtensionWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.equal(result.error.source, "readManifest");
      }
    });

    it("throws exception", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        templateState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox.stub(SpecParser.prototype, "validate").throws(new Error("test"));

      const generator = new MessageExtensionWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isErr());
    });

    it("throws specParser error", async function () {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ApiSpecLocation]: "https://test.com",
        [QuestionNames.ApiOperation]: ["operation1"],
        supportedApisFromApiSpec: apiOperations,
        templateState: {
          templateName: "copilot-plugin-existing-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.SME,
        },
      };
      const context = createContext();
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      sandbox
        .stub(SpecParser.prototype, "generate")
        .throws(new SpecParserError("test", ErrorType.Unknown));
      const generator = new MessageExtensionWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.equal(result.error.message, "test");
      }
    });
  });
});
