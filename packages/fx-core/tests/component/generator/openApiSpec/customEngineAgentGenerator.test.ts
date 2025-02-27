import { ProjectType, SpecParser, ValidationStatus } from "@microsoft/m365-spec-parser";
import { ApiOperation, Inputs, ok, Platform } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import { RestoreFn } from "mocked-env";
import * as sinon from "sinon";
import { createContext, setTools } from "../../../../src/common/globalVars";
import { manifestUtils } from "../../../../src/component/driver/teamsApp/utils/ManifestUtils";
import { CustomEngineAgentWithExistingApiSpecGenerator } from "../../../../src/component/generator/openApiSpec/customEngineAgentGenerator";
import * as helper from "../../../../src/component/generator/openApiSpec/helper";
import { TemplateNames } from "../../../../src/component/generator/templates/templateNames";
import {
  CapabilityOptions,
  DeclarativeAgentStartOptions,
  ProgrammingLanguage,
  QuestionNames,
} from "../../../../src/question";
import { MockTools } from "../../../core/utils";
import { teamsManifest } from "./fakeData";

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
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.DeclarativeAgentType]: DeclarativeAgentStartOptions.apiSpec().id,
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
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(SpecParser.prototype, "getFilteredSpecs").resolves([
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
      ]);
      sandbox.stub(helper, "updateForCustomApi").resolves([]);
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generate")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(helper, "generateScaffoldingSummary").resolves("");

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
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(SpecParser.prototype, "getFilteredSpecs").resolves([
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
      ]);
      sandbox.stub(helper, "updateForCustomApi").resolves([]);
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generate")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(helper, "generateScaffoldingSummary").resolves("");

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
      sandbox
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(SpecParser.prototype, "getFilteredSpecs").resolves([
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
      ]);
      sandbox.stub(helper, "updateForCustomApi").resolves([]);
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      const generateBasedOnSpec = sandbox
        .stub(SpecParser.prototype, "generate")
        .resolves({ allSuccess: true, warnings: [] });
      sandbox.stub(helper, "generateScaffoldingSummary").resolves("warning message");

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
        .stub(SpecParser.prototype, "validate")
        .resolves({ status: ValidationStatus.Valid, errors: [], warnings: [] });
      sandbox.stub(SpecParser.prototype, "getFilteredSpecs").resolves([
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
        {
          openapi: "3.0.0",
          info: {
            title: "test",
            version: "1.0",
          },
          paths: {},
        },
      ]);
      sandbox.stub(helper, "updateForCustomApi").throws(new Error("test"));
      sandbox.stub(fs, "ensureDir").resolves();
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok(teamsManifest));
      sandbox.stub(SpecParser.prototype, "generate").resolves({ allSuccess: true, warnings: [] });

      const generator = new CustomEngineAgentWithExistingApiSpecGenerator();
      const result = await generator.post(context, inputs, "projectPath");

      assert.isTrue(result.isErr() && result.error.message === "test");
    });
  });
});
