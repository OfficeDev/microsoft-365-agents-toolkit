// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yuqzho@microsoft.com
 */

import { err, Inputs, ok, Platform, PluginManifestSchema, UserError } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import { createContext, setTools, TOOLS } from "../../../src/common/globalVars";
import {
  ApiAuthOptions,
  ApiPluginStartOptions,
  CapabilityOptions,
  QuestionNames,
} from "../../../src/question";
import { CopilotExtensionGenerator } from "../../../src/component/generator/copilotExtension/generator";
import { TemplateNames } from "../../../src/component/generator/templates/templateNames";
import { RestoreFn } from "mocked-env";
import sinon from "sinon";
import { copilotGptManifestUtils } from "../../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import * as generatorHelper from "../../../src/component/generator/copilotExtension/helper";
import { pluginManifestUtils } from "../../../src/component/driver/teamsApp/utils/PluginManifestUtils";
import fs from "fs-extra";
import path from "path";
import { MockLogProvider, MockTools } from "../../core/utils";
import * as commons from "../../../src/component/utils/common";
import * as oneDriveHandler from "../../../src/component/generator/copilotExtension/oneDriveSharePointHandler";
import { ItemMetadata } from "../../../src/component/generator/copilotExtension/oneDriveSharePointHandler";
import { OneDriveSharePointItemType } from "../../../src/component/generator/constant";
import { AxiosInstance } from "axios";
import { SystemError } from "@microsoft/teamsfx-api";

describe("copilotExtension", async () => {
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
      const tools = new MockTools();
      setTools(tools);
      const generator = new CopilotExtensionGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.newApi().id,
        [QuestionNames.TemplateName]: TemplateNames.ApiPluginFromScratch,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "app",
      };
      let res = await generator.activate(context, inputs);
      let info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch");

      inputs[QuestionNames.ApiAuth] = ApiAuthOptions.apiKey().id;
      inputs[QuestionNames.TemplateName] = TemplateNames.ApiPluginFromScratchBearer;
      res = await generator.activate(context, inputs);
      info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch-bearer");

      inputs[QuestionNames.ApiAuth] = ApiAuthOptions.oauth().id;
      inputs[QuestionNames.TemplateName] = TemplateNames.ApiPluginFromScratchOAuth;
      res = await generator.activate(context, inputs);
      info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch-oauth");

      inputs[QuestionNames.ApiAuth] = ApiAuthOptions.microsoftEntra().id;
      inputs[QuestionNames.TemplateName] = TemplateNames.ApiPluginFromScratchOAuth;
      res = await generator.activate(context, inputs);
      info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch-oauth");
    });

    it("declarative Copilot: Env func enabled", async () => {
      const generator = new CopilotExtensionGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.declarativeAgent().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.newApi().id,
        [QuestionNames.TemplateName]: TemplateNames.ApiPluginFromScratch,
        [QuestionNames.ApiAuth]: ApiAuthOptions.none().id,
        [QuestionNames.AppName]: "app",
      };
      let res = await generator.activate(context, inputs);
      let info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch");

      inputs[QuestionNames.ApiAuth] = ApiAuthOptions.apiKey().id;
      inputs[QuestionNames.TemplateName] = TemplateNames.ApiPluginFromScratchBearer;
      res = await generator.activate(context, inputs);
      info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch-bearer");

      inputs[QuestionNames.ApiAuth] = ApiAuthOptions.oauth().id;
      inputs[QuestionNames.TemplateName] = TemplateNames.ApiPluginFromScratchOAuth;
      res = await generator.activate(context, inputs);
      info = await generator.getTemplateInfos(context, inputs, ".");
      assert.isTrue(res);
      assert.equal(info.isOk() && info.value[0].templateName, "api-plugin-from-scratch-oauth");
    });
  });

  describe("post", async () => {
    it("add plugin success", async () => {
      const generator = new CopilotExtensionGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.existingPlugin().id,
        [QuestionNames.AppName]: "app",
      };

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
      const generator = new CopilotExtensionGenerator();
      const context = createContext();

      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.existingPlugin().id,
        [QuestionNames.AppName]: "app",
      };

      const logStub = sandbox.stub(MockLogProvider.prototype, "info").resolves();
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
      assert.isFalse(logStub.called);
      assert.isTrue(res.isOk());

      res = await generator.post(context, { ...inputs, platform: Platform.CLI }, "");
      assert.isTrue(res.isOk());
      assert.isTrue(logStub.called);

      res = await generator.post(context, { ...inputs, platform: Platform.VS }, "");
      assert.isTrue(logStub.called);
      assert.isTrue(res.isOk());
    });
    it("get manifest path error", async () => {
      const generator = new CopilotExtensionGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.existingPlugin().id,
        [QuestionNames.AppName]: "app",
      };

      sandbox
        .stub(copilotGptManifestUtils, "getManifestPath")
        .resolves(err(new UserError("fakeError", "fakeError", "fakeError", "fakeError")));

      const res = await generator.post(context, inputs, "");
      assert.isTrue(res.isErr() && res.error.name === "fakeError");
    });

    it("add plugin errror", async () => {
      const generator = new CopilotExtensionGenerator();
      const context = createContext();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: CapabilityOptions.apiPlugin().id,
        [QuestionNames.ApiPluginType]: ApiPluginStartOptions.existingPlugin().id,
        [QuestionNames.AppName]: "app",
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
  });
});

describe("helper", async () => {
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
});

describe("getODSPItemInfo", () => {
  let mockedEnvRestore: RestoreFn | undefined;
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });

  beforeEach(() => {
    sandbox.restore();
  });

  it("should return error for undefined URL", async () => {
    const context = createContext();
    const result = await generatorHelper.getODSPItemInfo(context, undefined);
    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "InvalidInput");
      assert.equal(result.error.message, "Item URL is required");
    }
  });

  it("should handle SharePoint site path successfully", async () => {
    const context = createContext();
    const mockGraphClient = {} as AxiosInstance;
    const mockSite = {
      id: "site-id",
      name: "site-name",
      webId: "web-id",
      siteId: "site-id-value",
    };

    sandbox
      .stub(oneDriveHandler, "createGraphClientWithToken")
      .callsFake(async () => ok(mockGraphClient));
    sandbox
      .stub(oneDriveHandler, "getSharePointSiteByRelativePath")
      .callsFake(async () => ok(mockSite));

    const result = await generatorHelper.getODSPItemInfo(context, "https://sharepoint.site/path");

    assert.isTrue(result.isOk());
    if (result.isOk()) {
      assert.deepEqual(result.value[0], {
        id: mockSite.id,
        name: mockSite.name,
        webId: mockSite.webId,
        siteId: mockSite.siteId,
      });
    }
  });

  it("should handle drive item path successfully", async () => {
    const mockGraphClient = {} as AxiosInstance;
    const mockDriveItem: ItemMetadata = {
      id: "drive-id",
      name: "drive-name",
      uniqueId: "unique-id",
      listId: "list-id",
      webId: "web-id",
      siteId: "site-id",
      itemType: OneDriveSharePointItemType.File,
    };

    sandbox
      .stub(oneDriveHandler, "createGraphClientWithToken")
      .callsFake(async () => ok(mockGraphClient));
    sandbox
      .stub(oneDriveHandler, "getSharePointSiteByRelativePath")
      .callsFake(async () => err(new SystemError("test", "test", "test")));
    sandbox.stub(oneDriveHandler, "getDriveItemInfo").callsFake(async () => mockDriveItem);
    const context = createContext();
    const result = await generatorHelper.getODSPItemInfo(
      context,
      "https://sharepoint.site/drive/path"
    );

    assert.isTrue(result.isOk());
    if (result.isOk()) {
      assert.deepEqual(result.value[0], mockDriveItem);
    }
  });

  it("should handle Graph API error", async () => {
    const mockGraphClient = {} as AxiosInstance;
    const axiosError = {
      isAxiosError: true,
      response: {
        status: 404,
        data: { message: "Not Found" },
      },
      message: "Not Found",
    };

    sandbox
      .stub(oneDriveHandler, "createGraphClientWithToken")
      .callsFake(async () => ok(mockGraphClient));
    sandbox.stub(oneDriveHandler, "getSharePointSiteByRelativePath").throws(axiosError);
    sandbox.stub(oneDriveHandler, "getDriveItemInfo").throws(axiosError);
    const context = createContext();
    sandbox.stub(context.logProvider!, "error").resolves();

    const result = await generatorHelper.getODSPItemInfo(context, "https://sharepoint.site/path");

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "GraphApiError");
      assert.equal(result.error.source, "ValidateOneDriveSharePointURL");
    }
  });

  it("should handle system error", async () => {
    const mockGraphClient = {} as AxiosInstance;
    const systemError = new Error("System error occurred");
    const context = createContext();
    sandbox
      .stub(oneDriveHandler, "createGraphClientWithToken")
      .callsFake(async () => ok(mockGraphClient));
    sandbox.stub(oneDriveHandler, "getSharePointSiteByRelativePath").throws(systemError);
    sandbox.stub(oneDriveHandler, "getDriveItemInfo").throws(systemError);
    sandbox.stub(context.logProvider!, "error").resolves();

    const result = await generatorHelper.getODSPItemInfo(context, "https://sharepoint.site/path");

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "GraphApiError");
      assert.equal(result.error.source, "ValidateOneDriveSharePointURL");
    }
  });

  it("should handle graph client creation failure", async () => {
    const context = createContext();
    sandbox
      .stub(oneDriveHandler, "createGraphClientWithToken")
      .callsFake(async () => err(new UserError("source", "name", "message")));

    const result = await generatorHelper.getODSPItemInfo(context, "https://sharepoint.site/path");

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.name, "name");
      assert.equal(result.error.message, "message");
    }
  });
});
