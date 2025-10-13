// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yuqzho@microsoft.com
 */

import {
  err,
  Inputs,
  ok,
  Platform,
  PluginManifestSchema,
  UserError,
  signedIn,
  DeclarativeAgentManifest,
  signedOut,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import { RestoreFn } from "mocked-env";
import path from "path";
import sinon from "sinon";
import { createContext, setTools } from "../../../src/common/globalVars";
import { copilotGptManifestUtils } from "../../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import { pluginManifestUtils } from "../../../src/component/driver/teamsApp/utils/PluginManifestUtils";
import { DeclarativeAgentGenerator } from "../../../src/component/generator/declarativeAgent/generator";
import * as generatorHelper from "../../../src/component/generator/declarativeAgent/helper";
import { TemplateNames } from "../../../src/component/generator/templates/templateNames";
import * as commons from "../../../src/component/utils/common";
import { ActionStartOptions, ApiAuthOptions, QuestionNames } from "../../../src/question";
import { MockLogProvider, MockTools } from "../../core/utils";
import { GraphClient } from "../../../src/client/graphClient";
import { featureFlagManager } from "../../../src/common/featureFlags";
import * as utils from "../../../src/component/generator/utils";
import { DACapabilityOptions } from "../../../src/question/scaffold/vsc/CapabilityOptions";

describe("copilotExtension", async () => {
  setTools(new MockTools());
  let mockedEnvRestore: RestoreFn | undefined;
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });
  describe("activate and get template name", async () => {
    it("api plugin", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: ActionStartOptions.newApi().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromScratch,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "app",
      };
      let res = await generator.activate(context, inputs);
      let info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch");

      inputs[QuestionNames.ApiAuth] = ApiAuthOptions.apiKey().id;
      inputs[QuestionNames.TemplateName] =
        TemplateNames.DeclarativeAgentWithActionFromScratchBearer;
      res = await generator.activate(context, inputs);
      info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch-bearer");

      inputs[QuestionNames.ApiAuth] = ApiAuthOptions.oauth().id;
      inputs[QuestionNames.TemplateName] = TemplateNames.DeclarativeAgentWithActionFromScratchOAuth;
      res = await generator.activate(context, inputs);
      info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch-oauth");

      inputs[QuestionNames.ApiAuth] = ApiAuthOptions.microsoftEntra().id;
      inputs[QuestionNames.TemplateName] = TemplateNames.DeclarativeAgentWithActionFromScratchOAuth;
      res = await generator.activate(context, inputs);
      info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch-oauth");

      inputs[QuestionNames.TemplateName] = TemplateNames.DeclarativeAgentWithTypeSpec;
      res = await generator.activate(context, inputs);
      info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "declarative-agent-typespec");

      inputs[QuestionNames.TemplateName] = TemplateNames.DeclarativeAgentWithActionFromScratch;
      inputs.platform = Platform.VS;
      inputs[QuestionNames.ProgrammingLanguage] = "csharp";
      res = await generator.activate(context, inputs);
      info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch");
    });

    it("MCP server URL processing", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: ActionStartOptions.newApi().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "TestApp",
        [QuestionNames.MCPForDAServerUrl]: "https://example-mcp-server.com:8080/api",
      };

      const res = await generator.activate(context, inputs);
      const info = await generator.getTemplateInfos(context, inputs, ".");

      assert.isTrue(res);
      assert.isTrue(info.isOk());

      if (info.isOk() && info.value[0].replaceMap) {
        const replaceMap = info.value[0].replaceMap;

        // Verify MCPForDAServerUrl is included in replace map
        assert.equal(replaceMap.MCPForDAServerUrl, "https://example-mcp-server.com:8080/api");

        // Verify ServerName is correctly generated (host with alphanumeric only, max 10 chars)
        assert.equal(replaceMap.ServerName, "examplemcp");

        // Verify template name
        assert.equal(info.value[0].templateName, "declarative-agent-with-action-from-mcp");
      }
    });

    it("MCP server URL processing with special characters", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: ActionStartOptions.newApi().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "TestApp",
        [QuestionNames.MCPForDAServerUrl]: "https://my-server-123.example.com/mcp",
      };

      const res = await generator.activate(context, inputs);
      const info = await generator.getTemplateInfos(context, inputs, ".");

      assert.isTrue(res);
      assert.isTrue(info.isOk());

      if (info.isOk() && info.value[0].replaceMap) {
        const replaceMap = info.value[0].replaceMap;

        // Verify MCPForDAServerUrl is included in replace map
        assert.equal(replaceMap.MCPForDAServerUrl, "https://my-server-123.example.com/mcp");

        // Verify ServerName removes special characters and limits to 10 chars
        assert.equal(replaceMap.ServerName, "myserver12");

        // Verify template name
        assert.equal(info.value[0].templateName, "declarative-agent-with-action-from-mcp");
      }
    });

    it("MCP server URL processing - no URL provided", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: ActionStartOptions.newApi().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "TestApp",
        // No MCPForDAServerUrl provided
      };

      const res = await generator.activate(context, inputs);
      const info = await generator.getTemplateInfos(context, inputs, ".");

      assert.isTrue(res);
      assert.isTrue(info.isOk());

      if (info.isOk() && info.value[0].replaceMap) {
        const replaceMap = info.value[0].replaceMap;

        // Verify MCPForDAServerUrl and ServerName are not included when URL is not provided
        assert.isUndefined(replaceMap.MCPForDAServerUrl);
        assert.isUndefined(replaceMap.ServerName);

        // Verify template name
        assert.equal(info.value[0].templateName, "declarative-agent-with-action-from-mcp");
      }
    });

    it("declarative Copilot: Env func enabled", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: DACapabilityOptions.declarativeAgent().id,
        [QuestionNames.ActionType]: ActionStartOptions.newApi().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromScratch,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "app",
      };
      let res = await generator.activate(context, inputs);
      let info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch");

      inputs[QuestionNames.ApiAuth] = ApiAuthOptions.apiKey().id;
      inputs[QuestionNames.TemplateName] =
        TemplateNames.DeclarativeAgentWithActionFromScratchBearer;
      res = await generator.activate(context, inputs);
      info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch-bearer");

      inputs[QuestionNames.ApiAuth] = ApiAuthOptions.oauth().id;
      inputs[QuestionNames.TemplateName] = TemplateNames.DeclarativeAgentWithActionFromScratchOAuth;
      res = await generator.activate(context, inputs);
      info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch-oauth");
    });
  });

  describe("post", async () => {
    it("add plugin success", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithExistingAction,
        [QuestionNames.AppName]: "app",
      };

      sandbox.stub(utils, "setGeneralSensitivityLabel").resolves();
      sandbox
        .stub(copilotGptManifestUtils, "getManifestPath")
        .resolves(ok("declarativeAgent.json"));
      sandbox
        .stub(generatorHelper, "addExistingPlugin")
        .resolves(ok({ destinationPluginManifestPath: "test.json", warnings: [] }));

      let res = await generator.post(context, inputs, "");
      assert.isTrue(res.isOk());

      res = await generator.post(context, { ...inputs, platform: Platform.CLI }, "");
      assert.isTrue(res.isOk());

      res = await generator.post(context, { ...inputs, platform: Platform.VS }, "");
      assert.isTrue(res.isOk());
    });

    it("add plugin success with warnings", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();

      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithExistingAction,
        [QuestionNames.AppName]: "app",
      };

      const logStub = sandbox.stub(MockLogProvider.prototype, "info").resolves();
      // mock sensitivity label feature flag
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox.stub(utils, "setGeneralSensitivityLabel").resolves();
      sandbox
        .stub(copilotGptManifestUtils, "getManifestPath")
        .resolves(ok("declarativeAgent.json"));
      sandbox.stub(generatorHelper, "addExistingPlugin").resolves(
        ok({
          destinationPluginManifestPath: "test.json",
          warnings: [{ type: "test", content: "warningContent" }],
        })
      );

      let res = await generator.post(context, inputs, "");
      assert.isTrue(res.isOk());

      res = await generator.post(context, { ...inputs, platform: Platform.CLI }, "");
      assert.isTrue(res.isOk());
      assert.isTrue(logStub.called);

      res = await generator.post(context, { ...inputs, platform: Platform.VS }, "");
      assert.isTrue(logStub.called);
      assert.isTrue(res.isOk());
    });
    it("skip get manifest path error", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromExistingApiSpec,
        [QuestionNames.AppName]: "app",
      };

      sandbox
        .stub(copilotGptManifestUtils, "getManifestPath")
        .resolves(err(new UserError("fakeError", "fakeError", "fakeError", "fakeError")));

      const res = await generator.post(context, inputs, "");
      assert.isTrue(res.isOk());
    });
    it("get manifest path error - existing action", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithExistingAction,
        [QuestionNames.AppName]: "app",
      };

      sandbox
        .stub(copilotGptManifestUtils, "getManifestPath")
        .resolves(err(new UserError("fakeError", "fakeError", "fakeError", "fakeError")));

      const res = await generator.post(context, inputs, "");
      assert.isTrue(res.isErr() && res.error.name === "fakeError");
    });

    it("add plugin errror", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.AppName]: "app",
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithExistingAction,
      };

      sandbox
        .stub(copilotGptManifestUtils, "getManifestPath")
        .resolves(ok("declarativeAgent.json"));
      sandbox
        .stub(generatorHelper, "addExistingPlugin")
        .resolves(err(new UserError("fakeError", "fakeError", "fakeError", "fakeError")));

      const res = await generator.post(context, inputs, "");
      assert.isTrue(res.isErr() && res.error.name === "fakeError");
    });

    it("add EmbeddedKnowledge folder success - CLI", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: DACapabilityOptions.declarativeAgent().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentBasic,
        [QuestionNames.AppName]: "app",
      };

      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox.stub(fs, "ensureDir").resolves();

      const res = await generator.post(context, inputs, "");
      assert.isTrue(res.isOk());
    });

    it("add EmbeddedKnowledge folder success - VSC", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "./",
        [QuestionNames.Capabilities]: DACapabilityOptions.declarativeAgent().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentBasic,
        [QuestionNames.AppName]: "app",
      };

      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox.stub(fs, "ensureDir").resolves();

      const res = await generator.post(context, inputs, "");
      assert.isTrue(res.isOk());
    });

    it("add EmbeddedKnowledge folder skipped - VS", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.VS,
        projectPath: "./",
        [QuestionNames.Capabilities]: DACapabilityOptions.declarativeAgent().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentBasic,
        [QuestionNames.AppName]: "app",
      };

      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      sandbox.stub(fs, "ensureDir").throws("error");

      const res = await generator.post(context, inputs, "");
      assert.isTrue(res.isOk());
    });

    it("add EmbeddedKnowledge folder skipped - feature flag off", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "./",
        [QuestionNames.Capabilities]: DACapabilityOptions.declarativeAgent().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentBasic,
        [QuestionNames.AppName]: "app",
      };

      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox.stub(fs, "ensureDir").throws("error");

      const res = await generator.post(context, inputs, "");
      assert.isTrue(res.isOk());
    });
  });

  describe("setGeneralSensitivityLabel", async () => {
    const context = createContext();
    const manifestPath = "test/manifest.json";

    it("success", async () => {
      const infoStub = sandbox.stub(context.logProvider!, "info");
      const tokenStub = sandbox
        .stub(context.tokenProvider!.m365TokenProvider, "getStatus")
        .resolves(
          ok({
            status: signedIn,
            token: "fake-token",
          })
        );
      const getLabelStub = sandbox.stub(GraphClient.prototype, "getGeneralSentivityLabel").resolves(
        ok({
          id: "label-id",
        })
      );
      const DAManifest = {
        name: "test",
        description: "test description",
      } as DeclarativeAgentManifest;
      const readStub = sandbox
        .stub(copilotGptManifestUtils, "readDeclarativeAgentManifestFile")
        .resolves(ok(DAManifest as any));
      const writeStub = sandbox
        .stub(copilotGptManifestUtils, "writeDeclarativeAgentManifestFile")
        .resolves(ok(undefined));

      await utils.setGeneralSensitivityLabel(context, manifestPath);

      assert.isTrue(tokenStub.calledOnce);
      assert.isTrue(getLabelStub.calledOnceWith("fake-token"));
      assert.isTrue(readStub.calledOnceWith(manifestPath));
      assert.isTrue(writeStub.calledOnce);
      assert.deepEqual(writeStub.firstCall.args[0], {
        name: "test",
        description: "test description",
        sensitivity_label: {
          id: "label-id",
        },
      } as any);
      assert.equal(writeStub.firstCall.args[1], manifestPath);
      assert.isFalse(infoStub.called);
      assert.isTrue(DAManifest.sensitivity_label.id === "label-id");
    });

    it("token provider error", async () => {
      const infoStub = sandbox.stub(context.logProvider, "info");
      sandbox
        .stub(context.tokenProvider!.m365TokenProvider, "getStatus")
        .resolves(err(new UserError("source", "name", "message")));

      await utils.setGeneralSensitivityLabel(context, manifestPath);

      assert.isTrue(infoStub.calledOnce);

      const contextWithoutProvider = createContext() as any;
      contextWithoutProvider.tokenProvider = undefined;
      contextWithoutProvider.logProvider = undefined;
      await utils.setGeneralSensitivityLabel(contextWithoutProvider, manifestPath);
      assert.isTrue(infoStub.calledOnce);
    });

    it("not signed in", async () => {
      const infoStub = sandbox.stub(context.logProvider, "info");
      sandbox.stub(context.tokenProvider!.m365TokenProvider, "getStatus").resolves(
        ok({
          status: signedOut,
          token: undefined,
        })
      );

      await utils.setGeneralSensitivityLabel(context, manifestPath);
      assert.isTrue(infoStub.calledOnce);
    });
    it("not signed in  - no logger", async () => {
      const contextWithoutProvider = createContext() as any;
      const infoStub = sandbox.stub(contextWithoutProvider.logProvider, "info");
      sandbox.stub(contextWithoutProvider.tokenProvider!.m365TokenProvider, "getStatus").resolves(
        ok({
          status: signedOut,
          token: undefined,
        })
      );
      contextWithoutProvider.logProvider = undefined;
      await utils.setGeneralSensitivityLabel(contextWithoutProvider, manifestPath);
      assert.isFalse(infoStub.calledOnce);
    });
    it("token undefined", async () => {
      const infoStub = sandbox.stub(context.logProvider, "info");
      sandbox.stub(context.tokenProvider!.m365TokenProvider, "getStatus").resolves(
        ok({
          status: signedIn,
          token: undefined,
        })
      );

      await utils.setGeneralSensitivityLabel(context, manifestPath);

      assert.isTrue(infoStub.calledOnce);

      const contextWithoutProvider = createContext() as any;
      contextWithoutProvider.logProvider = undefined;
      await utils.setGeneralSensitivityLabel(contextWithoutProvider, manifestPath);
      assert.isTrue(infoStub.calledOnce);
    });

    it("get label id error", async () => {
      const infoStub = sandbox.stub(context.logProvider!, "info");
      sandbox.stub(context.tokenProvider!.m365TokenProvider, "getStatus").resolves(
        ok({
          status: signedIn,
          token: "fake-token",
        })
      );
      sandbox
        .stub(GraphClient.prototype, "getGeneralSentivityLabel")
        .resolves(err(new UserError("source", "name", "message")));

      await utils.setGeneralSensitivityLabel(context, manifestPath);

      assert.isTrue(infoStub.calledOnce);

      const contextWithoutProvider = createContext() as any;
      contextWithoutProvider.logProvider = undefined;
      await utils.setGeneralSensitivityLabel(contextWithoutProvider, manifestPath);
      assert.isTrue(infoStub.calledOnce);
    });

    it("read manifest error", async () => {
      const infoStub = sandbox.stub(context.logProvider!, "info");
      sandbox.stub(context.tokenProvider!.m365TokenProvider, "getStatus").resolves(
        ok({
          status: signedIn,
          token: "fake-token",
        })
      );
      sandbox
        .stub(GraphClient.prototype, "getGeneralSentivityLabel")
        .resolves(ok({ id: "label-id" }));
      sandbox
        .stub(copilotGptManifestUtils, "readDeclarativeAgentManifestFile")
        .resolves(err(new UserError("source", "name", "message")));

      await utils.setGeneralSensitivityLabel(context, manifestPath);

      assert.isTrue(infoStub.calledOnce);

      const contextWithoutProvider = createContext() as any;
      contextWithoutProvider.logProvider = undefined;
      await utils.setGeneralSensitivityLabel(contextWithoutProvider, manifestPath);
      assert.isTrue(infoStub.calledOnce);
    });

    it("write manifest error", async () => {
      const infoStub = sandbox.stub(context.logProvider!, "info");
      sandbox.stub(context.tokenProvider!.m365TokenProvider, "getStatus").resolves(
        ok({
          status: signedIn,
          token: "fake-token",
        })
      );
      sandbox
        .stub(GraphClient.prototype, "getGeneralSentivityLabel")
        .resolves(ok({ id: "label-id" }));
      sandbox.stub(copilotGptManifestUtils, "readDeclarativeAgentManifestFile").resolves(
        ok({
          name: "test",
          description: "test description",
        } as any)
      );
      sandbox
        .stub(copilotGptManifestUtils, "writeDeclarativeAgentManifestFile")
        .resolves(err(new UserError("source", "name", "message")));

      await utils.setGeneralSensitivityLabel(context, manifestPath);

      assert.isTrue(infoStub.calledOnce);

      const contextWithoutProvider = createContext() as any;
      contextWithoutProvider.logProvider = undefined;
      await utils.setGeneralSensitivityLabel(contextWithoutProvider, manifestPath);
      assert.isTrue(infoStub.calledOnce);
    });
  });
});

describe("helper", async () => {
  setTools(new MockTools());
  let mockedEnvRestore: RestoreFn | undefined;
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });
  const context = createContext();

  describe("addExistingPlugin", async () => {
    it("success: need to update plugin manifest", async () => {
      sandbox.stub(pluginManifestUtils, "readPluginManifestFile").resolves(
        ok({
          schema_version: "v1",
          name_for_human: "${{file}}",
          runtimes: [{ type: "OpenApi", spec: { url: "test.json" } }],
        } as any)
      );
      sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));
      const getApiSpecPath = sandbox
        .stub(pluginManifestUtils, "getDefaultNextAvailableApiSpecPath")
        .resolves("nextApiSpec.json");
      sandbox
        .stub(copilotGptManifestUtils, "getDefaultNextAvailablePluginManifestPath")
        .resolves("nextPluginManifest.json");
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "ensureFile").resolves();
      sandbox.stub(fs, "copyFile").resolves();
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(fs, "readFile").resolves();
      sandbox.stub(commons, "getEnvironmentVariables").returns([]);
      const res = await generatorHelper.addExistingPlugin(
        "test.json",
        "originalManifest.json",
        "originalManifest.yaml",
        "id",
        context,
        "source"
      );
      assert.isTrue(res.isOk());
      assert.isTrue(getApiSpecPath.calledOnce);
    });

    it("success: no need to update plugin manifest", async () => {
      sandbox.stub(pluginManifestUtils, "readPluginManifestFile").resolves(
        ok({
          schema_version: "v1",
          name_for_human: "test",
          runtimes: [{ type: "OpenApi", spec: { url: "test.json" } }],
        } as any)
      );
      sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));
      const getApiSpecPath = sandbox
        .stub(pluginManifestUtils, "getDefaultNextAvailableApiSpecPath")
        .resolves("nextApiSpec.json");
      sandbox.stub(commons, "getEnvironmentVariables").returns([]);
      sandbox
        .stub(copilotGptManifestUtils, "getDefaultNextAvailablePluginManifestPath")
        .resolves("nextPluginManifest.json");
      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(path, "relative").returns("test");
      sandbox.stub(fs, "ensureFile").resolves();
      sandbox.stub(fs, "copyFile").resolves();
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(fs, "readFile").resolves();
      const res = await generatorHelper.addExistingPlugin(
        "test.json",
        "originalManifest.json",
        "originalManifest.yaml",
        "id",
        context,
        "source"
      );
      assert.isTrue(res.isOk());
      assert.isTrue(getApiSpecPath.notCalled);
    });

    it("success: has warning", async () => {
      sandbox.stub(pluginManifestUtils, "readPluginManifestFile").resolves(
        ok({
          schema_version: "v1",
          name_for_human: "test",
          runtimes: [{ type: "OpenApi", spec: { url: "test.json" } }],
        } as any)
      );
      sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));
      const getApiSpecPath = sandbox
        .stub(pluginManifestUtils, "getDefaultNextAvailableApiSpecPath")
        .resolves("nextApiSpec.json");
      sandbox.stub(commons, "getEnvironmentVariables").returns(["TEST_ENV"]);
      sandbox
        .stub(copilotGptManifestUtils, "getDefaultNextAvailablePluginManifestPath")
        .resolves("nextPluginManifest.json");
      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(path, "relative").returns("test");
      sandbox.stub(fs, "ensureFile").resolves();
      sandbox.stub(fs, "copyFile").resolves();
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(fs, "readFile").resolves();
      const res = await generatorHelper.addExistingPlugin(
        "test.json",
        "originalManifest.json",
        "originalManifest.yaml",
        "id",
        context,
        "source"
      );
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.warnings.length, 2);
      }
      assert.isTrue(getApiSpecPath.notCalled);
    });

    it("success: only get partial warning", async () => {
      sandbox.stub(pluginManifestUtils, "readPluginManifestFile").resolves(
        ok({
          schema_version: "v1",
          name_for_human: "test",
          runtimes: [{ type: "OpenApi", spec: { url: "test.json" } }],
        } as any)
      );
      sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));
      const getApiSpecPath = sandbox
        .stub(pluginManifestUtils, "getDefaultNextAvailableApiSpecPath")
        .resolves("nextApiSpec.json");
      sandbox.stub(commons, "getEnvironmentVariables").returns(["TEST_ENV"]);
      sandbox
        .stub(copilotGptManifestUtils, "getDefaultNextAvailablePluginManifestPath")
        .resolves("nextPluginManifest.json");
      sandbox.stub(fs, "pathExists").resolves(false);
      sandbox.stub(path, "relative").returns("test");
      sandbox.stub(fs, "ensureFile").resolves();
      sandbox.stub(fs, "copyFile").resolves();
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(fs, "readFile").throws();
      const res = await generatorHelper.addExistingPlugin(
        "test.json",
        "originalManifest.json",
        "originalManifest.yaml",
        "id",
        context,
        "source"
      );
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.warnings.length, 1);
      }
      assert.isTrue(getApiSpecPath.notCalled);
    });

    it("error: readPluginManifestFile Error", async () => {
      sandbox
        .stub(pluginManifestUtils, "readPluginManifestFile")
        .resolves(err(new UserError("fakeError", "fakeError", "fakeError", "fakeError")));

      const res = await generatorHelper.addExistingPlugin(
        "test.json",
        "originalManifest.json",
        "originalManifest.yaml",
        "id",
        context,
        "source"
      );
      assert.isTrue(res.isErr() && res.error.name === "fakeError");
    });

    it("error: add action error", async () => {
      sandbox.stub(pluginManifestUtils, "readPluginManifestFile").resolves(
        ok({
          schema_version: "v1",
          name_for_human: "test",
          runtimes: [{ type: "OpenApi", spec: { url: "test.json" } }],
        } as any)
      );
      sandbox
        .stub(copilotGptManifestUtils, "addAction")
        .resolves(err(new UserError("fakeError", "fakeError", "fakeError", "fakeError")));
      const getApiSpecPath = sandbox
        .stub(pluginManifestUtils, "getDefaultNextAvailableApiSpecPath")
        .resolves("nextApiSpec.json");
      sandbox
        .stub(copilotGptManifestUtils, "getDefaultNextAvailablePluginManifestPath")
        .resolves("nextPluginManifest.json");
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "ensureFile").resolves();
      sandbox.stub(fs, "copyFile").resolves();
      sandbox.stub(fs, "writeFile").resolves();
      const res = await generatorHelper.addExistingPlugin(
        "test.json",
        "originalManifest.json",
        "originalManifest.yaml",
        "id",
        context,
        "source"
      );
      assert.isTrue(res.isErr() && res.error.name === "fakeError");
    });
  });

  describe("validateSourcePluginManifest", () => {
    it("Invalid manist", () => {
      const manifest: PluginManifestSchema = {
        schema_version: "",
        name_for_human: "test",
      } as any;
      manifest.runtimes = [{ type: "OpenApi", spec: { url: "test.json" } }];

      let res = generatorHelper.validateSourcePluginManifest(manifest as any, "source");
      assert.isTrue(res.isErr() && res.error.name === "MissingSchemaVersion");

      manifest.schema_version = "v1";
      delete manifest.runtimes;
      res = generatorHelper.validateSourcePluginManifest(manifest as any, "source");

      assert.isTrue(res.isErr() && res.error.name === "MissingRuntimes");

      manifest.runtimes = [
        { type: "OpenApi", spec: { url: "test.json" } },
        { type: "OpenApi", spec: { url: "test2.json" } },
      ];
      res = generatorHelper.validateSourcePluginManifest(manifest as any, "source");
      assert.isTrue(res.isErr() && res.error.name === "MultipleApiSpecInPluginManifest");

      manifest.runtimes = [{ type: "OpenApi" } as any];
      res = generatorHelper.validateSourcePluginManifest(manifest as any, "source");
      assert.isTrue(res.isErr() && res.error.name === "MissingApiSpec");
    });
  });

  describe("generateForMCPForDA", () => {
    const testDestinationPath = "/test/destination";
    const testAiPluginPath = path.join(testDestinationPath, "appPackage", "ai-plugin.json");

    it("error: ai-plugin.json file not found", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com",
        [QuestionNames.MCPForDAServerName]: "testServer",
        [QuestionNames.MCPForDATool]: "dynamic-fetch",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "PluginManifestNotFound");
      }
    });

    it("success: dynamic fetch tool configuration", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        name_for_human: "Test Plugin",
        functions: [{ name: "old-function" }],
        runtimes: [],
      };

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        [QuestionNames.MCPForDAServerName]: "testServer",
        [QuestionNames.MCPForDATool]: "dynamic-fetch",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      assert.isTrue(writeJSONStub.calledOnce);

      const writtenContent = writeJSONStub.firstCall.args[1];
      assert.deepEqual(writtenContent.functions, []);
      assert.equal(writtenContent.runtimes.length, 1);
      assert.equal(writtenContent.runtimes[0].type, "RemoteMCPServer");
      assert.equal(writtenContent.runtimes[0].spec.url, "https://example.com/mcp");
      assert.equal(writtenContent.runtimes[0].spec.enable_dynamic_discovery, true);
    });

    it("success: pre-fetch tool configuration without auth", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        name_for_human: "Test Plugin",
        functions: [],
        runtimes: [],
      };

      const mockToolsDetail = [
        {
          name: "testServer_tool1",
          description: "Tool 1 description",
          inputSchema: {
            type: "object",
            properties: { param1: { type: "string" } },
            required: ["param1"],
          },
          tags: ["tag1"],
        },
        {
          name: "testServer_tool2",
          description: "Tool 2 description",
          inputSchema: {
            type: "object",
            properties: { param2: { type: "number" } },
          },
          tags: ["tag2"],
        },
        {
          name: "otherServer_tool3",
          description: "Tool 3 from other server",
          inputSchema: { type: "object" },
          tags: [],
        },
      ];

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        [QuestionNames.MCPForDAServerName]: "testServer",
        [QuestionNames.MCPForDATool]: "pre-fetch",
        [QuestionNames.MCPForDAAvailableTools]: mockToolsDetail,
        [QuestionNames.MCPForDAPreFetchTools]: ["tool1", "tool2"],
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      assert.isTrue(writeJSONStub.calledOnce);

      const writtenContent = writeJSONStub.firstCall.args[1];
      assert.equal(writtenContent.functions.length, 2);

      // Check first function
      assert.equal(writtenContent.functions[0].name, "tool1");
      assert.equal(writtenContent.functions[0].description, "Tool 1 description");
      assert.deepEqual(writtenContent.functions[0].parameters, {
        type: "object",
        properties: { param1: { type: "string" } },
        required: ["param1"],
      });

      // Check second function
      assert.equal(writtenContent.functions[1].name, "tool2");
      assert.equal(writtenContent.functions[1].description, "Tool 2 description");
      assert.deepEqual(writtenContent.functions[1].parameters, {
        type: "object",
        properties: { param2: { type: "number" } },
        required: [],
      });

      // Check runtime configuration
      assert.equal(writtenContent.runtimes.length, 1);
      assert.equal(writtenContent.runtimes[0].type, "RemoteMCPServer");
      assert.equal(writtenContent.runtimes[0].spec.url, "https://example.com/mcp");
      assert.equal(writtenContent.runtimes[0].spec.enable_dynamic_discovery, false);
      assert.deepEqual(writtenContent.runtimes[0].run_for_functions, ["tool1", "tool2"]);
      assert.isUndefined(writtenContent.runtimes[0].auth);
    });

    it("success: pre-fetch tool configuration with OAuth auth", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        name_for_human: "Test Plugin",
        functions: [],
        runtimes: [],
      };

      const mockToolsDetail = [
        {
          name: "testServer_authenticatedTool",
          description: "Authenticated tool",
          inputSchema: {
            type: "object",
            properties: { data: { type: "string" } },
          },
          tags: [],
        },
      ];

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://secure.example.com/mcp",
        [QuestionNames.MCPForDAServerName]: "testServer",
        [QuestionNames.MCPForDATool]: "pre-fetch",
        [QuestionNames.MCPForDAAvailableTools]: mockToolsDetail,
        [QuestionNames.MCPForDAPreFetchTools]: ["authenticatedTool"],
        [QuestionNames.MCPForDAAuth]: "OAuthPluginVault",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      assert.isTrue(writeJSONStub.calledOnce);

      const writtenContent = writeJSONStub.firstCall.args[1];
      assert.equal(writtenContent.functions.length, 1);
      assert.equal(writtenContent.functions[0].name, "authenticatedTool");

      // Check runtime has auth configuration
      assert.equal(writtenContent.runtimes.length, 1);
      assert.deepEqual(writtenContent.runtimes[0].auth, {
        type: "OAuthPluginVault",
        reference_id: "${{MCP_DA_AUTH_ID}}",
      });
    });

    it("error: pre-fetch tools missing available tools", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves({});

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        [QuestionNames.MCPForDAServerName]: "testServer",
        [QuestionNames.MCPForDATool]: "pre-fetch",
        // Missing MCPForDAAvailableTools
        [QuestionNames.MCPForDAPreFetchTools]: ["tool1"],
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "PreFetchToolsNotFound");
      }
    });

    it("error: pre-fetch tools missing selected tools", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves({});

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        [QuestionNames.MCPForDAServerName]: "testServer",
        [QuestionNames.MCPForDATool]: "pre-fetch",
        [QuestionNames.MCPForDAAvailableTools]: [{ name: "tool1" }],
        // Missing MCPForDAPreFetchTools
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "PreFetchToolsNotFound");
      }
    });

    it("success: filters tools by server name correctly", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        name_for_human: "Test Plugin",
        functions: [],
        runtimes: [],
      };

      const mockToolsDetail = [
        {
          name: "serverA_toolX",
          description: "Tool X from server A",
          inputSchema: { type: "object" },
          tags: [],
        },
        {
          name: "serverB_toolY",
          description: "Tool Y from server B",
          inputSchema: { type: "object" },
          tags: [],
        },
        {
          name: "serverA_toolZ",
          description: "Tool Z from server A",
          inputSchema: { type: "object" },
          tags: [],
        },
      ];

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        [QuestionNames.MCPForDAServerName]: "serverA",
        [QuestionNames.MCPForDATool]: "pre-fetch",
        [QuestionNames.MCPForDAAvailableTools]: mockToolsDetail,
        [QuestionNames.MCPForDAPreFetchTools]: ["toolX", "toolZ"],
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());

      const writtenContent = writeJSONStub.firstCall.args[1];
      assert.equal(writtenContent.functions.length, 2);
      assert.equal(writtenContent.functions[0].name, "toolX");
      assert.equal(writtenContent.functions[1].name, "toolZ");

      // Should not include toolY from serverB
      const toolNames = writtenContent.functions.map((f: any) => f.name);
      assert.notInclude(toolNames, "toolY");
    });

    it("success: handles missing inputSchema properties gracefully", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        name_for_human: "Test Plugin",
        functions: [],
        runtimes: [],
      };

      const mockToolsDetail = [
        {
          name: "testServer_minimalTool",
          description: "Tool with minimal schema",
          inputSchema: {
            // Missing type, properties, required
          },
          tags: [],
        },
      ];

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        [QuestionNames.MCPForDAServerName]: "testServer",
        [QuestionNames.MCPForDATool]: "pre-fetch",
        [QuestionNames.MCPForDAAvailableTools]: mockToolsDetail,
        [QuestionNames.MCPForDAPreFetchTools]: ["minimalTool"],
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());

      const writtenContent = writeJSONStub.firstCall.args[1];
      assert.equal(writtenContent.functions.length, 1);
      assert.deepEqual(writtenContent.functions[0].parameters, {
        type: "object", // Default fallback
        properties: undefined,
        required: [], // Default fallback
      });
    });
  });
});
