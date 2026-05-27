// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yuqzho@microsoft.com
 */

import {
  DeclarativeAgentManifest,
  DynamicOptions,
  err,
  Inputs,
  MultiSelectQuestion,
  ok,
  OptionItem,
  Platform,
  PluginManifestSchema,
  signedIn,
  signedOut,
  SingleSelectQuestion,
  SystemError,
  UserError,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import { RestoreFn } from "mocked-env";
import path from "path";
import sinon from "sinon";
import { GraphClient } from "../../../src/client/graphClient";
import { featureFlagManager } from "../../../src/common/featureFlags";
import { createContext, setTools } from "../../../src/common/globalVars";
import { copilotGptManifestUtils } from "../../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import { pluginManifestUtils } from "../../../src/component/driver/teamsApp/utils/PluginManifestUtils";
import { developerPortalScaffoldUtils } from "../../../src/component/developerPortalScaffoldUtils";
import { DeclarativeAgentGenerator } from "../../../src/component/generator/declarativeAgent/generator";
import * as generatorHelper from "../../../src/component/generator/declarativeAgent/helper";
import * as oneDriveSharePointHandler from "../../../src/component/generator/declarativeAgent/oneDriveSharePointHandler";
import { TemplateNames } from "../../../src/component/generator/templates/templateNames";
import * as utils from "../../../src/component/generator/utils";
import * as commons from "../../../src/component/utils/common";
import { ODRProvider } from "../../../src/component/utils/odrProvider";
import { ActionStartOptions, ApiAuthOptions, QuestionNames } from "../../../src/question";
import {
  ActionStartOptions as CapabilityActionStartOptions,
  DACapabilityOptions,
} from "../../../src/question/scaffold/vsc/CapabilityOptions";
import {
  MCPLocalServerSelectionNode,
  MCPServerTypeNode,
} from "../../../src/question/scaffold/vsc/teamsProjectTypeNode";
import { MockLogProvider, MockTools } from "../../core/utils";

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

    it("post calls updateFilesForTdp when teamsAppFromTdp is set", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "./",
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentBasic,
        [QuestionNames.AppName]: "app",
        teamsAppFromTdp: { teamsAppId: "fake-id" },
      };

      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox
        .stub(copilotGptManifestUtils, "getManifestPath")
        .resolves(ok("declarativeAgent.json"));
      sandbox.stub(developerPortalScaffoldUtils, "updateFilesForTdp").resolves(ok(undefined));

      const res = await generator.post(context, inputs, "");
      assert.isTrue(res.isOk());
    });

    it("post returns error when updateFilesForTdp fails in TDP flow", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "./",
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentBasic,
        [QuestionNames.AppName]: "app",
        teamsAppFromTdp: { teamsAppId: "fake-id" },
      };

      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      sandbox
        .stub(copilotGptManifestUtils, "getManifestPath")
        .resolves(ok("declarativeAgent.json"));
      sandbox
        .stub(developerPortalScaffoldUtils, "updateFilesForTdp")
        .resolves(err(new UserError("fakeSource", "fakeError", "fakeError")));

      const res = await generator.post(context, inputs, "");
      assert.isTrue(res.isErr() && res.error.name === "fakeError");
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

    it("error: manifest fails validateSourcePluginManifest", async () => {
      sandbox.stub(pluginManifestUtils, "readPluginManifestFile").resolves(
        ok({
          // Missing schema_version → validation returns MissingSchemaVersion error
          name_for_human: "test",
          runtimes: [{ type: "OpenApi", spec: { url: "test.json" } }],
        } as any)
      );
      const res = await generatorHelper.addExistingPlugin(
        "test.json",
        "originalManifest.json",
        "originalManifest.yaml",
        "id",
        context,
        "source"
      );
      assert.isTrue(res.isErr() && res.error.name === "MissingSchemaVersion");
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

  describe("getODSPItemInfo", () => {
    it("error: missing itemUrl returns InvalidInput error", async () => {
      const res = await generatorHelper.getODSPItemInfo(context, undefined);
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "InvalidInput");
      }
    });

    it("error: createGraphClientWithToken fails propagates error", async () => {
      sandbox
        .stub(oneDriveSharePointHandler, "createGraphClientWithToken")
        .resolves(err(new UserError("source", "GetGraphTokenFailed", "msg", "msg")));

      const res = await generatorHelper.getODSPItemInfo(
        context,
        "https://example.sharepoint.com/sites/test"
      );
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "GetGraphTokenFailed");
      }
    });

    it("success: siteResult isOk returns site metadata", async () => {
      const fakeClient: any = {};
      sandbox
        .stub(oneDriveSharePointHandler, "createGraphClientWithToken")
        .resolves(ok(fakeClient));
      sandbox
        .stub(oneDriveSharePointHandler, "getSharePointSiteByRelativePath")
        .resolves(ok({ id: "site-id", name: "site-name", webId: "web-id", siteId: "s-id" }));

      const res = await generatorHelper.getODSPItemInfo(
        context,
        "https://example.sharepoint.com/sites/test"
      );
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.length, 1);
        assert.equal(res.value[0].id, "site-id");
        assert.equal(res.value[0].name, "site-name");
        assert.equal(res.value[0].webId, "web-id");
        assert.equal(res.value[0].siteId, "s-id");
      }
    });

    it("success: siteResult isErr falls through to getDriveItemInfo", async () => {
      const fakeClient: any = {};
      sandbox
        .stub(oneDriveSharePointHandler, "createGraphClientWithToken")
        .resolves(ok(fakeClient));
      sandbox
        .stub(oneDriveSharePointHandler, "getSharePointSiteByRelativePath")
        .resolves(err(new UserError("source", "GetSharePointSiteFailed", "msg", "msg")));
      sandbox.stub(oneDriveSharePointHandler, "getDriveItemInfo").resolves({
        id: "item-id",
        name: "item-name",
        uniqueId: "unique-id",
        listId: "list-id",
        webId: "web-id",
        siteId: "site-id",
      });

      const res = await generatorHelper.getODSPItemInfo(
        context,
        "https://example.sharepoint.com/personal/user_file.docx"
      );
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.equal(res.value.length, 1);
        assert.equal(res.value[0].id, "item-id");
        assert.equal(res.value[0].name, "item-name");
        assert.equal(res.value[0].uniqueId, "unique-id");
        assert.equal(res.value[0].listId, "list-id");
      }
    });

    it("error: axios error with 4xx status returns UserError", async () => {
      const fakeClient: any = {};
      sandbox
        .stub(oneDriveSharePointHandler, "createGraphClientWithToken")
        .resolves(ok(fakeClient));
      sandbox
        .stub(oneDriveSharePointHandler, "getSharePointSiteByRelativePath")
        .resolves(err(new UserError("source", "SiteFailed", "msg", "msg")));
      const axiosErr: any = Object.assign(new Error("Not Found"), {
        isAxiosError: true,
        response: { status: 404 },
      });
      sandbox.stub(oneDriveSharePointHandler, "getDriveItemInfo").rejects(axiosErr);
      sandbox.stub(context.logProvider!, "error");

      const res = await generatorHelper.getODSPItemInfo(
        context,
        "https://example.sharepoint.com/file.docx"
      );
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "GraphApiError");
        assert.isTrue(res.error instanceof UserError);
      }
    });

    it("error: non-axios error returns SystemError", async () => {
      const fakeClient: any = {};
      sandbox
        .stub(oneDriveSharePointHandler, "createGraphClientWithToken")
        .resolves(ok(fakeClient));
      sandbox
        .stub(oneDriveSharePointHandler, "getSharePointSiteByRelativePath")
        .resolves(err(new UserError("source", "SiteFailed", "msg", "msg")));
      sandbox
        .stub(oneDriveSharePointHandler, "getDriveItemInfo")
        .rejects(new Error("Unexpected network failure"));
      sandbox.stub(context.logProvider!, "error");

      const res = await generatorHelper.getODSPItemInfo(
        context,
        "https://example.sharepoint.com/file.docx"
      );
      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "GraphApiError");
        assert.isTrue(res.error instanceof SystemError);
      }
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
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "PluginManifestNotFound");
      }
    });

    it("success: no tools available returns warning and empty ai-plugin", async () => {
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
      };

      const mockFetchMCPTools = sandbox.stub().resolves({ requiresAuth: false, tools: [] });
      const importStub = sandbox
        .stub(generatorHelper, "generateForMCPForDA")
        .callsFake(async (destPath, inp) => {
          // Restore to call real implementation but with mocked fetch
          importStub.restore();
          // Pre-set empty tools to simulate no-tools scenario
          inp[QuestionNames.MCPForDAAvailableTools] = [];
          return generatorHelper.generateForMCPForDA(destPath, inp);
        });

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      if (res.isOk()) {
        // Should have no warnings since tools were pre-set as empty
        const writtenContent = writeJSONStub.firstCall.args[1];
        // Functions remain as template default (not modified since no tools)
        assert.deepEqual(writtenContent.functions, [{ name: "old-function" }]);
      }
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
        [QuestionNames.MCPForDAAvailableTools]: mockToolsDetail,
        [QuestionNames.MCPForDAPreFetchTools]: ["tool1", "tool2"],
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      // writeJSON called twice: mcp-tools-1.json and ai-plugin.json
      assert.isTrue(writeJSONStub.calledTwice);

      // First call: mcp-tools-1.json with full schemas
      const mcpToolsContent = writeJSONStub.firstCall.args[1];
      assert.equal(mcpToolsContent.tools.length, 2);
      assert.equal(mcpToolsContent.tools[0].name, "tool1");
      assert.isDefined(mcpToolsContent.tools[0].inputSchema);

      // Second call: ai-plugin.json with name+description only
      const writtenContent = writeJSONStub.secondCall.args[1];
      assert.equal(writtenContent.functions.length, 2);
      assert.equal(writtenContent.functions[0].name, "tool1");
      assert.equal(writtenContent.functions[0].description, "Tool 1 description");
      assert.isUndefined(writtenContent.functions[0].parameters);
      assert.equal(writtenContent.functions[1].name, "tool2");
      assert.equal(writtenContent.functions[1].description, "Tool 2 description");

      // Check runtime configuration
      assert.equal(writtenContent.runtimes.length, 1);
      assert.equal(writtenContent.runtimes[0].type, "RemoteMCPServer");
      assert.equal(writtenContent.runtimes[0].spec.url, "https://example.com/mcp");
      assert.equal(writtenContent.runtimes[0].spec.mcp_tool_description.file, "mcp-tools-1.json");
      assert.deepEqual(writtenContent.runtimes[0].run_for_functions, ["tool1", "tool2"]);
      assert.isUndefined(writtenContent.runtimes[0].auth);
    });

    it("error: pre-fetch tool configuration with OAuth auth but missing auth-type", async () => {
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
      sandbox.stub(fs, "writeJSON").resolves();

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://secure.example.com/mcp",
        [QuestionNames.MCPForDAServerName]: "testServer",
        [QuestionNames.MCPForDAAvailableTools]: mockToolsDetail,
        [QuestionNames.MCPForDAPreFetchTools]: ["authenticatedTool"],
        [QuestionNames.MCPForDAAuth]: "OAuthPluginVault",
        // No MCPForDAAuthType provided
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isErr());
      if (res.isErr()) {
        assert.equal(res.error.name, "MissingMCPAuthType");
      }
    });

    it("success: pre-fetch tool with OAuth auth and auth-type entraSSO", async () => {
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
      sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();
      // Stub readFile/writeFile for ActionInjector yml writing
      sandbox
        .stub(fs, "readFile")
        .resolves(
          "provision:\n  - uses: teamsApp/create\n    with:\n      name: test\n    writeToEnvironmentFile:\n      teamsAppId: TEAMS_APP_ID\n" as any
        );
      const writeFileStub = sandbox.stub(fs, "writeFile").resolves();

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://secure.example.com/mcp",
        [QuestionNames.MCPForDAServerName]: "testServer",
        [QuestionNames.MCPForDAAvailableTools]: mockToolsDetail,
        [QuestionNames.MCPForDAPreFetchTools]: ["authenticatedTool"],
        [QuestionNames.MCPForDAAuth]: "OAuthPluginVault",
        [QuestionNames.MCPForDAAuthType]: "entraSSO",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      // writeJSON called twice: mcp-tools-1.json and ai-plugin.json
      assert.isTrue(writeJSONStub.calledTwice);

      const writtenContent = writeJSONStub.secondCall.args[1];
      assert.equal(writtenContent.functions.length, 1);
      assert.equal(writtenContent.functions[0].name, "authenticatedTool");

      // Check runtime has auth configuration with correct registration ID
      assert.equal(writtenContent.runtimes.length, 1);
      assert.deepEqual(writtenContent.runtimes[0].auth, {
        type: "OAuthPluginVault",
        reference_id: "${{MCP_DA_AUTH_ID_ACTION_1}}",
      });

      // Verify ActionInjector wrote to yml — writeFile should be called with oauth/register
      assert.isTrue(writeFileStub.called);
      const ymlContent = writeFileStub.firstCall.args[1] as string;
      assert.include(ymlContent, "oauth/register");
      assert.include(ymlContent, "MCP_DA_AUTH_ID_ACTION_1");
    });

    it("success: no available tools leaves ai-plugin unchanged", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        functions: [],
        runtimes: [],
      };

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        [QuestionNames.MCPForDAServerName]: "testServer",
        // No MCPForDAAvailableTools, no MCPForDAPreFetchTools
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      // Mock fetchMCPTools to return empty
      const mcpToolFetcherModule = await import("../../../src/component/utils/mcpToolFetcher");
      sandbox
        .stub(mcpToolFetcherModule, "fetchMCPTools")
        .resolves({ requiresAuth: false, tools: [] });

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      if (res.isOk()) {
        // Should have a warning about no tools
        assert.isTrue(res.value.warnings!.length > 0);
      }
    });

    it("success: missing selected tools leaves ai-plugin unchanged", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        functions: [],
        runtimes: [],
      };

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        [QuestionNames.MCPForDAServerName]: "testServer",
        [QuestionNames.MCPForDAAvailableTools]: [{ name: "tool1" }],
        // Missing MCPForDAPreFetchTools
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      // writeJSON called once — just ai-plugin.json left unchanged
      const writtenContent = writeJSONStub.firstCall.args[1];
      assert.deepEqual(writtenContent.functions, []);
      assert.deepEqual(writtenContent.runtimes, []);
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
        [QuestionNames.MCPForDAAvailableTools]: mockToolsDetail,
        [QuestionNames.MCPForDAPreFetchTools]: ["toolX", "toolZ"],
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());

      // First call writes mcp-tools-1.json, second call writes ai-plugin.json
      const aiPluginContent = writeJSONStub.secondCall.args[1];
      assert.equal(aiPluginContent.functions.length, 2);
      assert.equal(aiPluginContent.functions[0].name, "toolX");
      assert.equal(aiPluginContent.functions[1].name, "toolZ");

      // Should not include toolY from serverB
      const toolNames = aiPluginContent.functions.map((f: any) => f.name);
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
        [QuestionNames.MCPForDAAvailableTools]: mockToolsDetail,
        [QuestionNames.MCPForDAPreFetchTools]: ["minimalTool"],
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());

      // First call writes mcp-tools-1.json with the raw inputSchema
      const mcpToolsContent = writeJSONStub.firstCall.args[1];
      assert.equal(mcpToolsContent.tools.length, 1);
      assert.equal(mcpToolsContent.tools[0].name, "minimalTool");

      // Second call writes ai-plugin.json with name+description only
      const aiPluginContent = writeJSONStub.secondCall.args[1];
      assert.equal(aiPluginContent.functions.length, 1);
      assert.equal(aiPluginContent.functions[0].name, "minimalTool");
      assert.isUndefined(aiPluginContent.functions[0].parameters);
    });

    it("success: OAuth auth-type resolves metadata and injects into yml", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        name_for_human: "Test Plugin",
        functions: [],
        runtimes: [],
      };

      const mockToolsDetail = [
        {
          name: "testServer_secureTool",
          description: "Secure tool",
          inputSchema: { type: "object" },
          tags: [],
        },
      ];

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();
      sandbox
        .stub(fs, "readFile")
        .resolves(
          "provision:\n  - uses: teamsApp/create\n    with:\n      name: test\n    writeToEnvironmentFile:\n      teamsAppId: TEAMS_APP_ID\n" as any
        );
      const writeFileStub = sandbox.stub(fs, "writeFile").resolves();

      const mcpToolFetcherModule = await import("../../../src/component/utils/mcpToolFetcher");
      sandbox.stub(mcpToolFetcherModule, "resolveMCPOAuthMetadata").resolves({
        authorizationUrl: "https://auth.example.com/authorize",
        tokenUrl: "https://auth.example.com/token",
        refreshUrl: "https://auth.example.com/token",
      });

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://secure.example.com/mcp",
        [QuestionNames.MCPForDAServerName]: "testServer",
        [QuestionNames.MCPForDAAvailableTools]: mockToolsDetail,
        [QuestionNames.MCPForDAPreFetchTools]: ["secureTool"],
        [QuestionNames.MCPForDAAuth]: "OAuthPluginVault",
        [QuestionNames.MCPForDAAuthType]: "oauth",
        [QuestionNames.MCPForDAAuthMetadataUrl]:
          "https://auth.example.com/.well-known/oauth-authorization-server",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      const aiPluginContent = writeJSONStub.secondCall.args[1];
      assert.deepEqual(aiPluginContent.runtimes[0].auth, {
        type: "OAuthPluginVault",
        reference_id: "${{MCP_DA_AUTH_ID_ACTION_1}}",
      });

      // Verify yml injection includes oauth URLs
      assert.isTrue(writeFileStub.called);
      const ymlContent = writeFileStub.firstCall.args[1] as string;
      assert.include(ymlContent, "oauth/register");
      assert.include(ymlContent, "https://auth.example.com/authorize");
      assert.include(ymlContent, "https://auth.example.com/token");
    });

    it("success: loads tools from file path (CLI flow)", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        functions: [],
        runtimes: [],
      };

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();

      const mcpToolFetcherModule = await import("../../../src/component/utils/mcpToolFetcher");
      sandbox.stub(mcpToolFetcherModule, "readMCPToolsFromFile").resolves([
        { name: "fileTool1", description: "File Tool 1", inputSchema: { type: "object" } },
        { name: "fileTool2", description: "File Tool 2", inputSchema: { type: "object" } },
      ]);
      sandbox.stub(mcpToolFetcherModule, "probeMCPServerAuth").resolves({
        requiresAuth: false,
      });

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        [QuestionNames.MCPToolsFilePath]: "/tmp/tools.json",
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      // Tools should have been loaded from file
      assert.equal(inputs[QuestionNames.MCPForDAAvailableTools].length, 2);
      assert.deepEqual(inputs[QuestionNames.MCPForDAPreFetchTools], ["fileTool1", "fileTool2"]);
      // writeJSON called twice: mcp-tools-1.json and ai-plugin.json
      assert.isTrue(writeJSONStub.calledTwice);
    });

    it("success: auto-fetch auth-required server sets auth metadata and warns", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        functions: [],
        runtimes: [],
      };

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();

      const mcpToolFetcherModule = await import("../../../src/component/utils/mcpToolFetcher");
      sandbox.stub(mcpToolFetcherModule, "fetchMCPTools").resolves({
        requiresAuth: true,
        tools: [],
        authMetadataUrl: "https://example.com/.well-known/oauth-authorization-server",
      });

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      if (res.isOk()) {
        // Should have auth-required warning
        assert.isTrue(res.value.warnings!.some((w) => w.type === "mcpAuthRequired"));
      }
      // Auth metadata should be set on inputs
      assert.equal(inputs[QuestionNames.MCPForDAAuth], "OAuthPluginVault");
      assert.equal(
        inputs[QuestionNames.MCPForDAAuthMetadataUrl],
        "https://example.com/.well-known/oauth-authorization-server"
      );
    });

    it("success: auto-fetch error adds warning and continues", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        functions: [],
        runtimes: [],
      };

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();

      const mcpToolFetcherModule = await import("../../../src/component/utils/mcpToolFetcher");
      sandbox.stub(mcpToolFetcherModule, "fetchMCPTools").rejects(new Error("Network error"));

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      if (res.isOk()) {
        // Should have fetch error warning
        assert.isTrue(res.value.warnings!.some((w) => w.type === "mcpFetchError"));
      }
    });

    it("success: file read error adds mcpToolsFileReadError warning and continues", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        functions: [],
        runtimes: [],
      };

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      sandbox.stub(fs, "writeJSON").resolves();

      const mcpToolFetcherModule = await import("../../../src/component/utils/mcpToolFetcher");
      sandbox.stub(mcpToolFetcherModule, "readMCPToolsFromFile").rejects(new Error("bad json"));
      sandbox
        .stub(mcpToolFetcherModule, "fetchMCPTools")
        .resolves({ requiresAuth: false, tools: [] });

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        [QuestionNames.MCPForDAServerName]: "testServer",
        [QuestionNames.MCPToolsFilePath]: "/tmp/bad-tools.json",
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.isTrue(res.value.warnings!.some((w) => w.type === "mcpToolsFileReadError"));
      }
    });

    it("success: auth probe after file load returns requiresAuth=true sets MCPForDAAuth", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        functions: [],
        runtimes: [],
      };

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      sandbox.stub(fs, "writeJSON").resolves();

      const mcpToolFetcherModule = await import("../../../src/component/utils/mcpToolFetcher");
      sandbox.stub(mcpToolFetcherModule, "probeMCPServerAuth").resolves({
        requiresAuth: true,
      });

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        // Pre-set MCPForDAAvailableTools to trigger the auth probe block
        // (file load skipped, no MCPToolsFilePath)
        [QuestionNames.MCPForDAAvailableTools]: [
          { name: "fileTool1", description: "File Tool 1", inputSchema: { type: "object" } },
        ],
        // No MCPForDAAuth pre-set → probe should set it
        // No MCPForDAPreFetchTools → function will skip tool writing and return ok with warning
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      // Function returns ok (no selected tools → skips tool writing, returns warning)
      assert.isTrue(res.isOk());
      // Auth should be set by the probe
      assert.equal(inputs[QuestionNames.MCPForDAAuth], "OAuthPluginVault");
    });

    it("success: auth probe after file load sets authMetadataUrl when present", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        functions: [],
        runtimes: [],
      };

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      sandbox.stub(fs, "writeJSON").resolves();

      const mcpToolFetcherModule = await import("../../../src/component/utils/mcpToolFetcher");
      sandbox.stub(mcpToolFetcherModule, "probeMCPServerAuth").resolves({
        requiresAuth: true,
        authMetadataUrl: "https://auth.example.com/.well-known/oauth",
      });

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        // Pre-set tools to trigger auth probe without going through file load
        [QuestionNames.MCPForDAAvailableTools]: [
          { name: "fileTool1", description: "File Tool 1", inputSchema: { type: "object" } },
        ],
        // No MCPForDAAuth, no MCPForDAPreFetchTools → probe sets auth, function returns ok
      };

      await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.equal(inputs[QuestionNames.MCPForDAAuth], "OAuthPluginVault");
      assert.equal(
        inputs[QuestionNames.MCPForDAAuthMetadataUrl],
        "https://auth.example.com/.well-known/oauth"
      );
    });

    it("success: auth probe throws after file load continues without auth errors", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        functions: [],
        runtimes: [],
      };

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      sandbox.stub(fs, "writeJSON").resolves();

      const mcpToolFetcherModule = await import("../../../src/component/utils/mcpToolFetcher");
      sandbox.stub(mcpToolFetcherModule, "probeMCPServerAuth").rejects(new Error("network error"));

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        // Pre-set tools to trigger the auth probe, no MCPForDAAuth
        [QuestionNames.MCPForDAAvailableTools]: [
          { name: "fileTool1", description: "File Tool 1", inputSchema: { type: "object" } },
        ],
        // No MCPForDAAuth → probe block runs; no MCPForDAPreFetchTools → skips tool writing
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      // Should succeed even when probe throws (silent catch)
      assert.isTrue(res.isOk());
      if (res.isOk()) {
        // No auth-probe-failure warning — the catch is silent
        assert.isFalse(res.value.warnings!.some((w) => w.type === "mcpAuthProbeError"));
      }
    });

    it("success: resolveMCPOAuthMetadata failure during OAuth injection adds mcpAuthMetadataError warning", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        name_for_human: "Test Plugin",
        functions: [],
        runtimes: [],
      };

      const mockToolsDetail = [
        {
          name: "testServer_secureTool",
          description: "Secure tool",
          inputSchema: { type: "object" },
          tags: [],
        },
      ];

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      sandbox.stub(fs, "writeJSON").resolves();

      const mcpToolFetcherModule = await import("../../../src/component/utils/mcpToolFetcher");
      sandbox
        .stub(mcpToolFetcherModule, "resolveMCPOAuthMetadata")
        .rejects(new Error("metadata unavailable"));

      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://secure.example.com/mcp",
        [QuestionNames.MCPForDAServerName]: "testServer",
        [QuestionNames.MCPForDAAvailableTools]: mockToolsDetail,
        [QuestionNames.MCPForDAPreFetchTools]: ["secureTool"],
        [QuestionNames.MCPForDAAuth]: "OAuthPluginVault",
        [QuestionNames.MCPForDAAuthType]: "oauth",
        [QuestionNames.MCPForDAAuthMetadataUrl]:
          "https://auth.example.com/.well-known/oauth-authorization-server",
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      if (res.isOk()) {
        assert.isTrue(res.value.warnings!.some((w) => w.type === "mcpAuthMetadataError"));
      }
    });

    it("success: auto-fetch returns non-empty tools and sets MCPForDAAvailableTools", async () => {
      const existingPluginContent = {
        schema_version: "v1",
        functions: [],
        runtimes: [],
      };

      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readJSON").resolves(existingPluginContent);
      const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();

      const mcpToolFetcherModule = await import("../../../src/component/utils/mcpToolFetcher");
      sandbox.stub(mcpToolFetcherModule, "fetchMCPTools").resolves({
        requiresAuth: false,
        tools: [
          { name: "autoTool1", description: "Auto Tool 1", inputSchema: { type: "object" } },
          { name: "autoTool2", description: "Auto Tool 2", inputSchema: { type: "object" } },
        ],
      });

      const testDestinationPath = "/test/destination";
      const inputs: Inputs = {
        platform: Platform.CLI,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        // No MCPForDAAvailableTools and no MCPForDAPreFetchTools pre-set
      };

      const res = await generatorHelper.generateForMCPForDA(testDestinationPath, inputs);

      assert.isTrue(res.isOk());
      // Lines 435-437: tools should be set from auto-fetch result
      assert.deepEqual(inputs[QuestionNames.MCPForDAAvailableTools], [
        { name: "autoTool1", description: "Auto Tool 1", inputSchema: { type: "object" } },
        { name: "autoTool2", description: "Auto Tool 2", inputSchema: { type: "object" } },
      ]);
      assert.deepEqual(inputs[QuestionNames.MCPForDAPreFetchTools], ["autoTool1", "autoTool2"]);
      // writeJSON called twice: mcp-tools-1.json + ai-plugin.json
      assert.isTrue(writeJSONStub.calledTwice);
    });
  });

  describe("generator.post MCPForDA branch", () => {
    it("post() calls generateForMCPForDA when template matches", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const destinationPath = "/test/destination";

      sandbox
        .stub(copilotGptManifestUtils, "getManifestPath")
        .resolves(ok("/test/destination/appPackage/da.json"));
      const generateStub = sandbox
        .stub(generatorHelper, "generateForMCPForDA")
        .resolves(ok({ warnings: [] }));

      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.AppName]: "TestApp",
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        [QuestionNames.MCPForDAServerUrl]: "https://example.com/mcp",
        [QuestionNames.MCPForDAAuth]: "NoneAuth",
      };

      const res = await generator.post(context, inputs, destinationPath);

      assert.isTrue(generateStub.calledOnce);
      assert.isTrue(res.isOk());
    });
  });

  describe("MCPServerTypeNode and Local MCP Server Support", async () => {
    it("MCPServerTypeNode should return correct structure", async () => {
      const mockServers = [
        {
          name: "test-server",
          display_name: "Test Server",
          description: "Test server",
          version: "1.0.0",
          identifier: "test.server",
          packageFamily: "packagefamily",
          command: "odr.exe",
          args: ["mcp", "--proxy", "test.server"],
          tools: [{ name: "tool1", description: "Tool 1", inputSchema: {} }],
        },
      ];

      const listServersStub = sandbox.stub(ODRProvider, "listServers").resolves(mockServers);

      const node = MCPServerTypeNode();

      // Test basic structure
      assert.isDefined(node);
      assert.isDefined(node.condition);
      assert.isDefined(node.data);
      assert.isDefined(node.children);

      // Test condition
      assert.deepEqual(node.condition, { equals: CapabilityActionStartOptions.mcp().id });

      const questionData = (await node.data) as SingleSelectQuestion;

      // Test data properties
      assert.equal(questionData.name, QuestionNames.MCPServerType);
      assert.equal(questionData.type, "singleSelect");
      assert.equal(questionData.default, "remote");
      assert.equal(questionData.staticOptions.length, 0);

      const inputs: Inputs = { platform: Platform.CLI };
      const options = (await (questionData.dynamicOptions as DynamicOptions)(
        inputs
      )) as OptionItem[];
      assert.equal(options.length, 2);
      assert.equal(options[0].id, "remote");
      assert.equal(options[1].id, "local");
      assert.isTrue(listServersStub.calledOnce);

      // Test children structure
      assert.equal(node.children?.length, 2);
    });

    it("MCPServerTypeNode should only show remote option when no servers available", async () => {
      const listServersStub = sandbox.stub(ODRProvider, "listServers").resolves([]);

      const node = MCPServerTypeNode();
      const questionData = (await node.data) as SingleSelectQuestion;

      const inputs: Inputs = { platform: Platform.CLI };
      const options = (await (questionData.dynamicOptions as DynamicOptions)(
        inputs
      )) as OptionItem[];

      assert.equal(options.length, 1);
      assert.equal(options[0].id, "remote");
      assert.isTrue(listServersStub.calledOnce);

      assert.deepEqual(inputs["_McpOdrOutput"], []);
    });

    it("MCPLocalServerSelectionNode should have correct structure for multiselect", async () => {
      const node = MCPLocalServerSelectionNode();
      const questionData = (await node.data) as MultiSelectQuestion;

      // Verify multiselect configuration
      assert.equal(questionData.type, "multiSelect");
      assert.equal(questionData.returnObject, true);
      assert.deepEqual(questionData.validation, { minItems: 1 });
      assert.equal(questionData.name, QuestionNames.MCPLocalServer);
    });

    it("MCPLocalServerSelectionNode dynamicOptions should work correctly", async () => {
      const mockServers = [
        {
          name: "test-server-1",
          display_name: "Test Server 1",
          description: "First test server",
          version: "1.0.0",
          identifier: "test.server.1",
          packageFamily: "packagefamily",
          command: "odr.exe",
          args: ["mcp", "--proxy", "test.server.1"],
          tools: [
            { name: "tool1", description: "Tool 1", inputSchema: {} },
            { name: "tool2", description: "Tool 2", inputSchema: {} },
          ],
        },
        {
          name: "test-server-2",
          display_name: "Test Server 2",
          description: "Second test server",
          version: "2.0.0",
          identifier: "test.server.2",
          packageFamily: "packagefamily",
          command: "odr.exe",
          args: ["mcp", "--proxy", "test.server.2"],
          tools: [{ name: "tool3", description: "Tool 3", inputSchema: {} }],
        },
      ];

      const node = MCPLocalServerSelectionNode();
      const questionData = (await node.data) as MultiSelectQuestion;

      // Set up the inputs with the mock servers as if MCPServerTypeNode had set them
      const inputs: Inputs = {
        platform: Platform.CLI,
        _McpOdrOutput: mockServers,
      };

      const options = (await (questionData.dynamicOptions as DynamicOptions)(
        inputs
      )) as OptionItem[];

      assert.isArray(options);
      assert.isDefined(options);
      assert.equal(options?.length, 2);

      // Test first option
      assert.equal(options[0].id, "test-server-1");
      assert.equal(options[0].label, "Test Server 1");
      assert.equal(options[0].detail, "First test server (2 tools available)");
      const data1 = JSON.parse(options[0].data as string);
      assert.equal(data1.identifier, "test.server.1");
      assert.equal(data1.command, "odr.exe");
      assert.deepEqual(data1.args, ["mcp", "--proxy", "test.server.1"]);

      // Test second option
      assert.equal(options[1].id, "test-server-2");
      assert.equal(options[1].label, "Test Server 2");
      assert.equal(options[1].detail, "Second test server (1 tools available)");
      const data2 = JSON.parse(options[1].data as string);
      assert.equal(data2.identifier, "test.server.2");
      assert.equal(data2.command, "odr.exe");
      assert.deepEqual(data2.args, ["mcp", "--proxy", "test.server.2"]);
    });

    it("processMCPLocalServers should handle single server selection", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: CapabilityActionStartOptions.mcp().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "TestApp",
        [QuestionNames.MCPServerType]: "local",
        [QuestionNames.MCPLocalServer]: [
          {
            id: "test-server",
            label: "Test Server",
            data: JSON.stringify({
              identifier: "test.server",
              command: "odr.exe",
              args: ["mcp", "--proxy", "test.server"],
            }),
          },
        ],
      };

      const res = await generator.activate(context, inputs);
      const info = await generator.getTemplateInfos(context, inputs, ".");

      assert.isTrue(res);
      assert.isTrue(info.isOk());

      if (info.isOk() && info.value[0].replaceMap) {
        const replaceMap = info.value[0].replaceMap;

        // Verify MCPLocalServers array is present
        assert.isDefined(replaceMap.MCPLocalServers);
        assert.isArray(replaceMap.MCPLocalServers);
        assert.equal(replaceMap.MCPLocalServers.length, 1);

        // Verify server details
        assert.equal(replaceMap.MCPLocalServers[0].name, "test-server");
        assert.equal(replaceMap.MCPLocalServers[0].identifier, "test.server");
        assert.equal(replaceMap.MCPLocalServers[0].command, "odr.exe");
        assert.equal(replaceMap.MCPLocalServers[0].args, '"mcp", "--proxy", "test.server"');
        assert.equal(replaceMap.MCPLocalServers[0].notLast, false);
      }
    });

    it("processMCPLocalServers should handle multiple server selection", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: CapabilityActionStartOptions.mcp().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "TestApp",
        [QuestionNames.MCPServerType]: "local",
        [QuestionNames.MCPLocalServer]: [
          {
            id: "server-1",
            label: "Server 1",
            data: JSON.stringify({
              identifier: "server.1",
              command: "odr.exe",
              args: ["mcp", "--proxy", "server.1"],
            }),
          },
          {
            id: "server-2",
            label: "Server 2",
            data: JSON.stringify({
              identifier: "server.2",
              command: "odr.exe",
              args: ["mcp", "--proxy", "server.2"],
            }),
          },
        ],
      };

      const res = await generator.activate(context, inputs);
      const info = await generator.getTemplateInfos(context, inputs, ".");

      assert.isTrue(res);
      assert.isTrue(info.isOk());

      if (info.isOk() && info.value[0].replaceMap) {
        const replaceMap = info.value[0].replaceMap;

        // Verify MCPLocalServers array has both servers
        assert.isDefined(replaceMap.MCPLocalServers);
        assert.isArray(replaceMap.MCPLocalServers);
        assert.equal(replaceMap.MCPLocalServers.length, 2);

        // Verify first server
        assert.equal(replaceMap.MCPLocalServers[0].name, "server-1");
        assert.equal(replaceMap.MCPLocalServers[0].notLast, true);

        // Verify second server
        assert.equal(replaceMap.MCPLocalServers[1].name, "server-2");
        assert.equal(replaceMap.MCPLocalServers[1].notLast, false);
      }
    });

    it("processMCPLocalServers should handle empty selection", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: CapabilityActionStartOptions.mcp().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "TestApp",
        [QuestionNames.MCPServerType]: "local",
        [QuestionNames.MCPLocalServer]: [],
      };

      const res = await generator.activate(context, inputs);
      const info = await generator.getTemplateInfos(context, inputs, ".");

      assert.isTrue(res);
      assert.isTrue(info.isOk());

      if (info.isOk() && info.value[0].replaceMap) {
        const replaceMap = info.value[0].replaceMap;

        // Verify empty array is returned
        assert.isDefined(replaceMap.MCPLocalServers);
        assert.isArray(replaceMap.MCPLocalServers);
        assert.equal(replaceMap.MCPLocalServers.length, 0);
      }
    });

    it("processMCPLocalServers throws when option data malformed", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: CapabilityActionStartOptions.mcp().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "TestApp",
        [QuestionNames.MCPServerType]: "local",
        [QuestionNames.MCPLocalServer]: [
          {
            id: "server-1",
            label: "Server 1",
            data: {
              identifier: "server.1",
              command: "odr.exe",
              args: ["mcp", "--proxy", "server.1"],
            },
          },
          {
            id: "server-2",
            label: "Server 2",
            data: JSON.stringify({
              identifier: "server.2",
              command: "odr.exe",
              args: ["mcp", "--proxy", "server.2"],
            }),
          },
        ],
      };

      const res = await generator.activate(context, inputs);
      const info = await generator.getTemplateInfos(context, inputs, ".");

      assert.isTrue(res);
      assert.isTrue(
        info.isErr() &&
          info.error.name === "processMCPLocalServers" &&
          info.error instanceof SystemError
      );
    });

    it("processMCPLocalServers throws UserError when option data missing command", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: CapabilityActionStartOptions.mcp().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "TestApp",
        [QuestionNames.MCPServerType]: "local",
        [QuestionNames.MCPLocalServer]: [
          {
            id: "server-1",
            label: "Server 1",
            data: JSON.stringify({
              identifier: "server.1",
              args: ["mcp", "--proxy", "server.1"],
            }),
          },
          {
            id: "server-2",
            label: "Server 2",
            data: JSON.stringify({
              identifier: "server.2",
              command: "odr.exe",
              args: ["mcp", "--proxy", "server.2"],
            }),
          },
        ],
      };

      const res = await generator.activate(context, inputs);
      const info = await generator.getTemplateInfos(context, inputs, ".");

      assert.isTrue(res);
      assert.isTrue(
        info.isErr() &&
          info.error.name === "processMCPLocalServers" &&
          info.error instanceof UserError
      );
    });

    it("processMCPLocalServers throws when MCPLocalServer malformed", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: CapabilityActionStartOptions.mcp().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "TestApp",
        [QuestionNames.MCPServerType]: "local",
        [QuestionNames.MCPLocalServer]: {
          id: "server-1",
          label: "Server 1",
          data: {
            identifier: "server.1",
            command: "odr.exe",
            args: ["mcp", "--proxy", "server.1"],
          },
        },
      };

      const res = await generator.activate(context, inputs);
      const info = await generator.getTemplateInfos(context, inputs, ".");

      assert.isTrue(res);
      assert.isTrue(
        info.isErr() &&
          info.error.name === "processMCPLocalServers" &&
          info.error instanceof SystemError
      );
    });

    it("processMCPLocalServers throws UserError when command malformed", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: CapabilityActionStartOptions.mcp().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "TestApp",
        [QuestionNames.MCPServerType]: "local",
        [QuestionNames.MCPLocalServer]: [
          {
            id: "server-1",
            label: "Server 1",
            data: JSON.stringify({
              identifier: "server.1",
              command: 12345, // Malformed command
              args: ["mcp", "--proxy", "server.1"],
            }),
          },
        ],
      };

      const res = await generator.activate(context, inputs);
      const info = await generator.getTemplateInfos(context, inputs, ".");

      assert.isTrue(res);
      assert.isTrue(
        info.isErr() &&
          info.error.name === "processMCPLocalServers" &&
          info.error instanceof UserError
      );
    });

    it("ODRProvider listServers should handle empty output", async () => {
      sandbox.stub(process, "platform").value("win32");
      const execStub = sandbox
        .stub(require("child_process"), "exec")
        .callsArgWith(1, null, JSON.stringify({ servers: [] }), "");

      const servers = await ODRProvider.listServers();

      assert.isArray(servers);
      assert.equal(servers.length, 0);
      assert.isTrue(execStub.calledOnce);
    });

    it("ODRProvider listServers should return empty array on non-Windows platform", async () => {
      sandbox.stub(process, "platform").value("darwin"); // macOS
      const execStub = sandbox.stub(require("child_process"), "exec");

      const servers = await ODRProvider.listServers();

      assert.isArray(servers);
      assert.equal(servers.length, 0);
      assert.isFalse(execStub.called); // Should not even attempt to call exec
    });

    it("ODRProvider listServers should return empty array when stdout is empty", async () => {
      sandbox.stub(process, "platform").value("win32");
      const execStub = sandbox.stub(require("child_process"), "exec").callsArgWith(1, null, "", "");

      const servers = await ODRProvider.listServers();

      assert.isArray(servers);
      assert.equal(servers.length, 0);
      assert.isTrue(execStub.calledOnce);
    });

    it("ODRProvider listServers should handle malformed JSON", async () => {
      sandbox.stub(process, "platform").value("win32");
      const execStub = sandbox
        .stub(require("child_process"), "exec")
        .callsArgWith(1, null, "invalid json", "");

      const servers = await ODRProvider.listServers();

      assert.isArray(servers);
      assert.equal(servers.length, 0);
      assert.isTrue(execStub.calledOnce);
    });

    it("ODRProvider listServers should handle exec errors", async () => {
      sandbox.stub(process, "platform").value("win32");
      const execStub = sandbox
        .stub(require("child_process"), "exec")
        .callsArgWith(1, new Error("Command failed"), "", "error output");

      const servers = await ODRProvider.listServers();

      assert.isArray(servers);
      assert.equal(servers.length, 0);
      assert.isTrue(execStub.calledOnce);
    });

    it("ODRProvider listServers should handle command not found (ODR not installed)", async () => {
      sandbox.stub(process, "platform").value("win32");
      const execStub = sandbox
        .stub(require("child_process"), "exec")
        .callsArgWith(
          1,
          new Error("'odr' is not recognized as an internal or external command"),
          "",
          ""
        );

      const servers = await ODRProvider.listServers();

      assert.isArray(servers);
      assert.equal(servers.length, 0);
      assert.isTrue(execStub.calledOnce);
    });

    it("ODRProvider parseODRListOutput should parse valid server data", async () => {
      const mockInput = {
        servers: [
          {
            name: "test-server",
            description: "Test server description",
            version: "1.0.0",
            packages: [{ identifier: "com.test.server" }],
            _meta: {
              "io.modelcontextprotocol.registry/publisher-provided": {
                "com.microsoft.windows": {
                  manifest: {
                    display_name: "Test MCP Server",
                    server: {
                      mcp_config: {
                        command: "odr.exe",
                        args: ["mcp", "--proxy", "com.test.server"],
                      },
                    },
                    _meta: {
                      "com.microsoft.windows": {
                        package_family_name: "TestPackageFamily",
                        static_responses: {
                          "tools/list": {
                            tools: [
                              {
                                name: "test-tool",
                                description: "Test tool description",
                                inputSchema: { type: "object", properties: {} },
                              },
                            ],
                          },
                        },
                      },
                    },
                  },
                },
              },
            },
          },
        ],
      };

      const servers = ODRProvider.parseODRListOutput(mockInput);

      assert.isArray(servers);
      assert.equal(servers.length, 1);

      const server = servers[0];
      assert.equal(server.name, "test-server");
      assert.equal(server.display_name, "Test MCP Server");
      assert.equal(server.description, "Test server description");
      assert.equal(server.version, "1.0.0");
      assert.equal(server.identifier, "com.test.server");
      assert.equal(server.command, "odr.exe");
      assert.deepEqual(server.args, ["mcp", "--proxy", "com.test.server"]);
      assert.equal(server.tools.length, 1);
      assert.equal(server.tools[0].name, "test-tool");
    });

    it("ODRProvider parseODRListOutput should filter servers without package family", async () => {
      const mockInput = {
        servers: [
          {
            name: "valid-server",
            _meta: {
              "io.modelcontextprotocol.registry/publisher-provided": {
                "com.microsoft.windows": {
                  manifest: {
                    server: {
                      mcp_config: {
                        command: "odr.exe",
                        args: ["mcp", "--proxy", "valid.server"],
                      },
                    },
                    _meta: {
                      "com.microsoft.windows": {
                        package_family_name: "ValidPackage",
                      },
                    },
                  },
                },
              },
            },
          },
          {
            name: "invalid-server",
            _meta: {
              "io.modelcontextprotocol.registry/publisher-provided": {
                "com.microsoft.windows": {
                  manifest: {
                    server: {
                      mcp_config: {
                        command: "odr.exe",
                        args: ["mcp", "--proxy", "invalid.server"],
                      },
                    },
                    _meta: {
                      "com.microsoft.windows": {
                        // Missing package_family_name
                      },
                    },
                  },
                },
              },
            },
          },
        ],
      };

      const servers = ODRProvider.parseODRListOutput(mockInput);

      assert.isArray(servers);
      assert.equal(servers.length, 1);
      assert.equal(servers[0].name, "valid-server");
    });

    it("ODRProvider parseODRListOutput should return empty array when servers is not an array", async () => {
      const mockInput = {
        servers: "not-an-array",
      };

      const servers = ODRProvider.parseODRListOutput(mockInput);

      assert.isArray(servers);
      assert.equal(servers.length, 0);
    });

    it("ODRProvider parseODRListOutput should return empty array when servers property is missing", async () => {
      const mockInput = {
        notServers: [],
      };

      const servers = ODRProvider.parseODRListOutput(mockInput);

      assert.isArray(servers);
      assert.equal(servers.length, 0);
    });

    it("ODRProvider parseODRListOutput should handle servers with empty tools list", async () => {
      const mockInput = {
        servers: [
          {
            name: "server-no-tools",
            packages: [{ identifier: "com.test.server" }],
            _meta: {
              "io.modelcontextprotocol.registry/publisher-provided": {
                "com.microsoft.windows": {
                  manifest: {
                    display_name: "Test Server",
                    server: {
                      mcp_config: {
                        command: "test-command",
                        args: ["test-arg"],
                      },
                    },
                    _meta: {
                      "com.microsoft.windows": {
                        package_family_name: "TestPackage",
                        static_responses: {
                          "tools/list": {
                            tools: [],
                          },
                        },
                      },
                    },
                  },
                },
              },
            },
          },
        ],
      };

      const servers = ODRProvider.parseODRListOutput(mockInput);

      assert.isArray(servers);
      assert.equal(servers.length, 1);
      assert.equal(servers[0].tools.length, 0);
    });

    it("processMCPLocalServers should throw when option data malformed JSON", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: CapabilityActionStartOptions.mcp().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "TestApp",
        [QuestionNames.MCPServerType]: "local",
        [QuestionNames.MCPLocalServer]: [
          {
            id: "valid-server",
            label: "Valid Server",
            data: JSON.stringify({
              identifier: "valid.server",
              command: "odr.exe",
              args: ["mcp", "--proxy", "valid.server"],
            }),
          },
          {
            id: "invalid-server",
            label: "Invalid Server",
            data: "not-valid-json",
          },
        ],
      };

      const res = await generator.activate(context, inputs);
      const info = await generator.getTemplateInfos(context, inputs, ".");

      assert.isTrue(res);
      assert.isTrue(info.isErr());
    });

    it("declarative agent generator should handle both remote and local MCP server configurations", async () => {
      const generator = new DeclarativeAgentGenerator();
      const context = createContext();

      // Test remote configuration
      const remoteInputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: CapabilityActionStartOptions.mcp().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "TestApp",
        [QuestionNames.MCPServerType]: "remote",
        [QuestionNames.MCPForDAServerUrl]: "https://remote-mcp.example.com",
      };

      let res = await generator.activate(context, remoteInputs);
      let info = await generator.getTemplateInfos(context, remoteInputs, ".");

      assert.isTrue(res);
      assert.isTrue(info.isOk());

      if (info.isOk() && info.value[0].replaceMap) {
        const replaceMap = info.value[0].replaceMap;
        assert.equal(replaceMap.MCPForDAServerUrl, "https://remote-mcp.example.com");
        assert.equal(replaceMap.ServerName, "remotemcpe");
      }

      // Test local configuration with multiselect
      const localInputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: "api-plugin",
        [QuestionNames.ActionType]: CapabilityActionStartOptions.mcp().id,
        [QuestionNames.TemplateName]: TemplateNames.DeclarativeAgentWithActionFromMCP,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "TestApp",
        [QuestionNames.MCPServerType]: "local",
        [QuestionNames.MCPLocalServer]: [
          {
            id: "local-server",
            label: "Local Server",
            data: JSON.stringify({
              identifier: "local.server.id",
              command: "odr.exe",
              args: ["mcp", "--proxy", "local.server.id"],
            }),
          },
        ],
      };

      res = await generator.activate(context, localInputs);
      info = await generator.getTemplateInfos(context, localInputs, ".");

      assert.isTrue(res);
      assert.isTrue(info.isOk());

      if (info.isOk() && info.value[0].replaceMap) {
        const replaceMap = info.value[0].replaceMap;
        // Verify MCPLocalServers array for local configuration
        assert.isDefined(replaceMap.MCPLocalServers);
        assert.isArray(replaceMap.MCPLocalServers);
        assert.equal(replaceMap.MCPLocalServers.length, 1);
        assert.equal(replaceMap.MCPLocalServers[0].name, "local-server");
      }
    });
  });
});

describe("deriveMCPServerNameFromUrl", () => {
  it("returns fallback when url is undefined", () => {
    assert.equal(generatorHelper.deriveMCPServerNameFromUrl(undefined), "mcpServer");
  });
  it("returns fallback when url is empty", () => {
    assert.equal(generatorHelper.deriveMCPServerNameFromUrl(""), "mcpServer");
  });
  it("returns fallback for invalid url", () => {
    assert.equal(generatorHelper.deriveMCPServerNameFromUrl("not a url"), "mcpServer");
  });
  it("strips non-alphanumeric characters from host", () => {
    const name = generatorHelper.deriveMCPServerNameFromUrl("https://my-host.example.com/mcp");
    assert.equal(name, "myhostexam");
  });
  it("truncates host to 10 characters", () => {
    const name = generatorHelper.deriveMCPServerNameFromUrl(
      "https://averyverylonghostname.example.com/path"
    );
    assert.equal(name.length, 10);
  });
  it("returns fallback when host is empty after stripping", () => {
    // Construct a URL whose host is only non-alphanumeric chars - in practice
    // URL will reject most such hosts, so "not a url" path covers it.
    const name = generatorHelper.deriveMCPServerNameFromUrl("file:///local/path");
    assert.equal(name, "mcpServer");
  });
});

describe("createNewActionPluginManifest", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });

  it("creates a new manifest, registers the action, and returns the path", async () => {
    sandbox
      .stub(copilotGptManifestUtils, "getDefaultNextAvailablePluginManifestPath")
      .resolves("/proj/appPackage/ai-plugin.json");
    const ensureFileStub = sandbox.stub(fs, "ensureFile").resolves();
    const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();
    const addActionStub = sandbox
      .stub(copilotGptManifestUtils, "addAction")
      .resolves(ok({} as any));

    const res = await generatorHelper.createNewActionPluginManifest(
      "/proj",
      "ai-plugin.json",
      "/proj/appPackage/declarativeAgent.json"
    );

    assert.isTrue(res.isOk());
    if (res.isOk()) {
      assert.equal(res.value.pluginManifestPath, "/proj/appPackage/ai-plugin.json");
      assert.equal(res.value.actionId, "ai-plugin");
    }
    assert.isTrue(ensureFileStub.calledOnce);
    assert.isTrue(writeJSONStub.calledOnce);
    assert.isTrue(addActionStub.calledOnce);

    const writtenSkeleton = writeJSONStub.firstCall.args[1] as any;
    assert.equal(writtenSkeleton.schema_version, "v2.4");
    assert.deepEqual(writtenSkeleton.functions, []);
    assert.deepEqual(writtenSkeleton.runtimes, []);
    assert.isString(writtenSkeleton.namespace);
  });

  it("falls back to DefaultPluginManifestFileName when desired name is empty", async () => {
    const getPathStub = sandbox
      .stub(copilotGptManifestUtils, "getDefaultNextAvailablePluginManifestPath")
      .resolves("/proj/appPackage/ai-plugin.json");
    sandbox.stub(fs, "ensureFile").resolves();
    sandbox.stub(fs, "writeJSON").resolves();
    sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));

    const res = await generatorHelper.createNewActionPluginManifest(
      "/proj",
      "   ",
      "/proj/appPackage/declarativeAgent.json"
    );
    assert.isTrue(res.isOk());
    // second argument should fall back to a non-empty default file name
    const fileNameArg = getPathStub.firstCall.args[1];
    assert.isString(fileNameArg);
    assert.notEqual((fileNameArg as string).trim(), "");
  });

  it("returns err when addAction fails", async () => {
    sandbox
      .stub(copilotGptManifestUtils, "getDefaultNextAvailablePluginManifestPath")
      .resolves("/proj/appPackage/ai-plugin.json");
    sandbox.stub(fs, "ensureFile").resolves();
    sandbox.stub(fs, "writeJSON").resolves();
    sandbox
      .stub(copilotGptManifestUtils, "addAction")
      .resolves(err(new SystemError("test", "addActionFailed", "msg", "msg")));

    const res = await generatorHelper.createNewActionPluginManifest(
      "/proj",
      "ai-plugin.json",
      "/proj/appPackage/declarativeAgent.json"
    );
    assert.isTrue(res.isErr());
  });

  it("derives a sanitized namespace from project folder name", async () => {
    sandbox
      .stub(copilotGptManifestUtils, "getDefaultNextAvailablePluginManifestPath")
      .resolves("/proj/appPackage/ai-plugin.json");
    sandbox.stub(fs, "ensureFile").resolves();
    const writeJSONStub = sandbox.stub(fs, "writeJSON").resolves();
    sandbox.stub(copilotGptManifestUtils, "addAction").resolves(ok({} as any));

    await generatorHelper.createNewActionPluginManifest(
      "/some/path/My-Cool App!",
      "ai-plugin.json",
      "/some/path/My-Cool App!/appPackage/declarativeAgent.json"
    );
    const written = writeJSONStub.firstCall.args[1] as any;
    assert.equal(written.namespace, "mycoolapp");
  });
});
