import { ErrorType, ProjectType, SpecParserError } from "@microsoft/m365-spec-parser";
import { Inputs, ok, Platform, SystemError } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import { RestoreFn } from "mocked-env";
import * as sinon from "sinon";
import { createContext, setTools } from "../../../../src/common/globalVars";
import {
  customEngineAgentGeneratorDeps,
  CustomEngineAgentWithExistingApiSpecGenerator,
} from "../../../../src/component/generator/openApiSpec/customEngineAgentGenerator";
import { TemplateNames } from "../../../../src/component/generator/templates/templateNames";
import { ActionStartOptions, ProgrammingLanguage, QuestionNames } from "../../../../src/question";
import { TeamsAgentCapabilityOptions } from "../../../../src/question/scaffold/vsc/CapabilityOptions";
import { MockTools } from "../../../core/utils";

const tools = new MockTools();

describe("CustomEngineAgentWithExistingApiSpecGenerator", async () => {
  const sandbox = sinon.createSandbox();
  before(() => {
    setTools(tools);
  });
  after(() => {
    sandbox.restore();
  });
  describe("activate", async () => {
    it("should activate and get correct template name", async () => {
      const generator = new CustomEngineAgentWithExistingApiSpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.TemplateName]: TemplateNames.CustomCopilotRagCustomApi,
      };
      const res = await generator.activate(context, inputs);
      assert.isTrue(res);
    });
  });

  describe("getTemplateInfos", async () => {
    let mockedEnvRestore: RestoreFn | undefined;
    afterEach(async () => {
      sandbox.restore();
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });
    it("happy path", async () => {
      const generator = new CustomEngineAgentWithExistingApiSpecGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: TeamsAgentCapabilityOptions.customCopilotRag().id,
        [QuestionNames.ActionType]: ActionStartOptions.apiSpec().id,
        [QuestionNames.TemplateName]: TemplateNames.CustomCopilotRagCustomApi,
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
        assert.equal(res.value[0].templateName, TemplateNames.CustomCopilotRagCustomApi);
        assert.equal(res.value[0].replaceMap?.["DeclarativeCopilot"], "");
      }
    });
  });

  describe("post", function () {
    const sandbox = sinon.createSandbox();
    let mockedEnvRestore: RestoreFn | undefined;

    before(() => {
      setTools(tools);
    });

    afterEach(async () => {
      sandbox.restore();
      if (mockedEnvRestore) {
        mockedEnvRestore();
      }
    });

    it("generateCustomCopilot: success", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        templateState: {
          templateName: "custom-copilot-rag-custom-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.TeamsAi,
        },
      };
      const context = createContext();
      const generateBasedOnSpec = sandbox
        .stub(customEngineAgentGeneratorDeps, "generateFilesFromApiSpec")
        .resolves(ok({}));

      const generator = new CustomEngineAgentWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
    });

    it("generateCustomCopilot for csharp: success", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.CSharp,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        templateState: {
          templateName: "custom-copilot-rag-custom-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.TeamsAi,
        },
      };
      const context = createContext();
      const generateBasedOnSpec = sandbox
        .stub(customEngineAgentGeneratorDeps, "generateFilesFromApiSpec")
        .resolves(ok({}));

      const generator = new CustomEngineAgentWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
    });

    it("generateCustomCopilot: CLI with warning", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        templateState: {
          templateName: "custom-copilot-rag-custom-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.TeamsAi,
        },
      };
      const context = createContext();
      const generateBasedOnSpec = sandbox
        .stub(customEngineAgentGeneratorDeps, "generateFilesFromApiSpec")
        .resolves(ok({}));

      const generator = new CustomEngineAgentWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isOk());
      assert.isTrue(generateBasedOnSpec.calledOnce);
    });

    it("generateCustomCopilot: error", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        templateState: {
          templateName: "custom-copilot-rag-custom-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.TeamsAi,
        },
      };
      const context = createContext();
      sandbox
        .stub(customEngineAgentGeneratorDeps, "generateFilesFromApiSpec")
        .rejects(new Error("test"));
      const assembleErrorStub = sandbox
        .stub(customEngineAgentGeneratorDeps, "assembleError")
        .returns(new SystemError("ut", "test", "", ""));

      const generator = new CustomEngineAgentWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isErr());
      assert.isTrue(assembleErrorStub.calledOnce);
    });

    it("generateCustomCopilot: SpecParserError", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "path",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.ApiSpecLocation]: "test.yaml",
        [QuestionNames.ApiOperation]: ["operation1"],
        templateState: {
          templateName: "custom-copilot-rag-custom-api",
          isPlugin: false,
          uri: "https://test.com",
          isYaml: false,
          type: ProjectType.TeamsAi,
        },
      };
      const context = createContext();
      const specError = new SpecParserError("test", ErrorType.Unknown);
      sandbox.stub(customEngineAgentGeneratorDeps, "generateFilesFromApiSpec").rejects(specError);
      const convertStub = sandbox
        .stub(customEngineAgentGeneratorDeps, "convertSpecParserErrorToFxError")
        .returns(new SystemError("ut", "spec-parser", "", ""));

      const generator = new CustomEngineAgentWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isErr());
      assert.isTrue(convertStub.calledOnceWithExactly(specError));
    });
  });
});
