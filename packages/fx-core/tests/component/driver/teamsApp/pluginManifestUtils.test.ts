// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as sinon from "sinon";
import chai from "chai";
import fs from "fs-extra";
import { pluginManifestUtils } from "../../../../src/component/driver/teamsApp/utils/PluginManifestUtils";
import {
  Colors,
  ManifestUtil,
  Platform,
  PluginManifestSchema,
  SystemError,
  TeamsAppManifest,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import {
  FileNotFoundError,
  JSONSyntaxError,
  MissingEnvironmentVariablesError,
} from "../../../../src";
import path from "path";
import { AppStudioError } from "../../../../src/component/driver/teamsApp/errors";
import { PluginManifestValidationResult } from "../../../../src/component/driver/teamsApp/interfaces/ValidationResult";
import mockedEnv, { RestoreFn } from "mocked-env";
import { MockedLogProvider, MockedTelemetryReporter } from "../../../plugins/solution/util";
import { createContext, setTools } from "../../../../src/common/globalVars";
import * as commonUtils from "../../../../src/common/utils";
import { WrapDriverContext } from "../../../../src/component/driver/util/wrapUtil";
import { MockTools } from "../../../core/utils";

describe("pluginManifestUtils", () => {
  const sandbox = sinon.createSandbox();
  let mockedEnvRestore: RestoreFn;

  afterEach(async () => {
    sandbox.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });

  const pluginManifest: PluginManifestSchema = {
    schema_version: "2.0",
    name_for_human: "test",
    description_for_human: "test",
    runtimes: [
      {
        type: "OpenApi",
        auth: { type: "None" },
        spec: {
          url: "openapi.yaml",
        },
      },
      {
        type: "LocalPlugin",
        spec: {
          local_endpoint: "localEndpoint",
        },
        runs_for_functions: ["add_todo"],
      },
    ],
  };

  const teamsManifest: TeamsAppManifest = {
    $schema:
      "https://developer.microsoft.com/en-us/json-schemas/teams/v1.9/MicrosoftTeams.schema.json",
    manifestVersion: "1.9",
    version: "1.0.0",
    id: "test",
    packageName: "test",
    developer: {
      name: "test",
      websiteUrl: "https://test.com",
      privacyUrl: "https://test.com/privacy",
      termsOfUseUrl: "https://test.com/termsofuse",
    },
    icons: {
      color: "icon-color.png",
      outline: "icon-outline.png",
    },
    name: {
      short: "test",
      full: "test",
    },
    description: {
      short: "test",
      full: "test",
    },
    accentColor: "#FFFFFF",
    bots: [],
    composeExtensions: [],
    configurableTabs: [],
    staticTabs: [],
    permissions: [],
    validDomains: [],
    copilotExtensions: {
      plugins: [
        {
          file: "resources/plugin.json",
          id: "plugin1",
        },
      ],
    },
  };

  it("readPluginManifestFile success", async () => {
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").resolves(JSON.stringify(pluginManifest) as any);

    const result = await pluginManifestUtils.readPluginManifestFile("path");
    chai.assert.isTrue(result.isOk());
    if (result.isOk()) {
      chai.assert.deepEqual(result.value, pluginManifest);
    }
  });

  it("readPluginManifestFile error: JsonSyntaxError", async () => {
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").resolves("invalid json" as any);

    const result = await pluginManifestUtils.readPluginManifestFile("path");
    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof JSONSyntaxError);
    }
  });

  it("readPluginManifestFile error: file does not exist", async () => {
    sandbox.stub(fs, "pathExists").resolves(false);

    const result = await pluginManifestUtils.readPluginManifestFile("path");
    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof FileNotFoundError);
    }
  });

  it("getApiSpecFilePathFromTeamsManifest sucess", async () => {
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").resolves(JSON.stringify(pluginManifest) as any);
    const res = await pluginManifestUtils.getApiSpecFilePathFromTeamsManifest(
      teamsManifest,
      "/test/path"
    );
    chai.assert.isTrue(res.isOk());

    if (res.isOk()) {
      chai.assert.isTrue(res.value.length === 1);
      chai.assert.equal(res.value[0], path.resolve("/test/resources/openapi.yaml"));
    }
  });

  it("getApiSpecFilePathFromTeamsManifest error: plugin file not exist", async () => {
    sandbox.stub(fs, "pathExists").resolves(false);
    const readPlugin = sandbox.stub(fs, "readFile").resolves(JSON.stringify(pluginManifest) as any);
    const res = await pluginManifestUtils.getApiSpecFilePathFromTeamsManifest(
      teamsManifest,
      "path"
    );
    chai.assert.isTrue(res.isErr());

    if (res.isErr()) {
      chai.assert.isTrue(res.error instanceof FileNotFoundError);
      chai.assert.isTrue(readPlugin.notCalled);
    }
  });

  it("getApiSpecFilePathFromTeamsManifest error: invalid plugin node case 1", async () => {
    const testManifest = {
      ...teamsManifest,
      copilotExtensions: { plugins: [] },
    };
    sandbox.stub(fs, "readFile").resolves(JSON.stringify(pluginManifest) as any);
    const res = await pluginManifestUtils.getApiSpecFilePathFromTeamsManifest(
      testManifest,
      "/test/path"
    );
    chai.assert.isTrue(res.isErr());

    if (res.isErr()) {
      chai.assert.equal(res.error.name, AppStudioError.TeamsAppRequiredPropertyMissingError.name);
    }
  });

  it("getApiSpecFilePathFromTeamsManifest error: invalid plugin node case 2", async () => {
    const testManifest = {
      $schema:
        "https://developer.microsoft.com/en-us/json-schemas/teams/v1.9/MicrosoftTeams.schema.json",
      manifestVersion: "1.9",
      version: "1.0.0",
      id: "test",
      packageName: "test",
      developer: {
        name: "test",
        websiteUrl: "https://test.com",
        privacyUrl: "https://test.com/privacy",
        termsOfUseUrl: "https://test.com/termsofuse",
      },
      icons: {
        color: "icon-color.png",
        outline: "icon-outline.png",
      },
      name: {
        short: "test",
        full: "test",
      },
      description: {
        short: "test",
        full: "test",
      },
    };
    sandbox.stub(fs, "readFile").resolves(JSON.stringify(pluginManifest) as any);
    const res = await pluginManifestUtils.getApiSpecFilePathFromTeamsManifest(
      testManifest as unknown as TeamsAppManifest,
      "/test/path"
    );
    chai.assert.isTrue(res.isErr());

    if (res.isErr()) {
      chai.assert.equal(res.error.name, AppStudioError.TeamsAppRequiredPropertyMissingError.name);
    }
  });

  it("getApiSpecFilePathFromTeamsManifest error: spec file not exist", async () => {
    sandbox.stub(fs, "pathExists").callsFake(async (testPath) => {
      if (testPath === path.resolve("/test/resources/openapi.yaml")) {
        return false;
      } else {
        return true;
      }
    });
    sandbox.stub(fs, "readFile").resolves(JSON.stringify(pluginManifest) as any);
    const res = await pluginManifestUtils.getApiSpecFilePathFromTeamsManifest(
      teamsManifest,
      "/test/path"
    );
    chai.assert.isTrue(res.isOk());

    if (res.isOk()) {
      chai.assert.equal(res.value.length, 0);
    }
  });

  it("getApiSpecFilePathFromTeamsManifest error: runtime without url", async () => {
    const testPluginManifest = {
      ...pluginManifest,
      runtimes: [
        {
          type: "OpenApi",
          auth: { type: "None" },
          spec: {
            url: "",
          },
        },
      ],
    };
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").resolves(JSON.stringify(testPluginManifest) as any);
    const res = await pluginManifestUtils.getApiSpecFilePathFromTeamsManifest(
      teamsManifest,
      "/test/path"
    );
    chai.assert.isTrue(res.isOk());

    if (res.isOk()) {
      chai.assert.equal(res.value.length, 0);
    }
  });

  it("getApiSpecFilePathFromTeamsManifest error: teams manifest without plugin", async () => {
    sandbox.stub(fs, "pathExists").resolves(true);

    const res = await pluginManifestUtils.getApiSpecFilePathFromTeamsManifest(
      { ...teamsManifest, copilotExtensions: {} },
      "/test/path"
    );
    chai.assert.isTrue(res.isErr());
  });

  describe("logValidationErrors", () => {
    it("skip if no errors", () => {
      const validationRes: PluginManifestValidationResult = {
        id: "1",
        filePath: "testPath",
        validationResult: [],
      };

      const res = pluginManifestUtils.logValidationErrors(validationRes, Platform.VSCode);
      chai.assert.isEmpty(res);
    });
    it("log if VSC", () => {
      const validationRes: PluginManifestValidationResult = {
        id: "1",
        filePath: "testPath",
        validationResult: ["error1", "error2"],
      };

      const res = pluginManifestUtils.logValidationErrors(validationRes, Platform.VSCode) as string;

      chai.assert.isTrue(res.includes("error1"));
      chai.assert.isTrue(res.includes("error2"));
    });

    it("log if CLI", () => {
      const validationRes: PluginManifestValidationResult = {
        id: "1",
        filePath: "testPath",
        validationResult: ["error1", "error2"],
      };

      const res = pluginManifestUtils.logValidationErrors(validationRes, Platform.CLI) as Array<{
        content: string;
        color: Colors;
      }>;

      chai.assert.isTrue(res.find((item) => item.content.includes("error1")) !== undefined);
      chai.assert.isTrue(res.find((item) => item.content.includes("error2")) !== undefined);
    });
  });

  describe("getManifest", async () => {
    setTools(new MockTools());
    const context = commonUtils.generateDriverContext(createContext(), {
      platform: Platform.VSCode,
      projectPath: "",
    });
    const mockedContex = new WrapDriverContext(context, "test", "test");
    const testPluginManifest = {
      ...pluginManifest,
      name_for_human: "name${{APP_NAME_SUFFIX}}",
      runtimes: [
        {
          type: "OpenApi",
          auth: { type: "None" },
          spec: {
            url: "",
          },
        },
      ],
    };
    it("get manifest success", async () => {
      mockedEnvRestore = mockedEnv({
        ["APP_NAME_SUFFIX"]: "test",
      });
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(JSON.stringify(testPluginManifest) as any);

      const res = await pluginManifestUtils.getManifest("testPath", mockedContex);

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.equal("nametest", res.value.name_for_human);
      }
    });

    it("get manifest error: file not found", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      const res = await pluginManifestUtils.getManifest("testPath", mockedContex);
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof FileNotFoundError);
      }
    });

    it("get manifest error: unresolved env error", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(JSON.stringify(testPluginManifest) as any);

      const res = await pluginManifestUtils.getManifest("testPath", mockedContex);

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof MissingEnvironmentVariablesError);
      }
    });
  });

  describe("validateAgainstSchema", async () => {
    const driverContext = {
      logProvider: new MockedLogProvider(),
      telemetryReporter: new MockedTelemetryReporter(),
      projectPath: "test",
      addTelemetryProperties: () => {},
    };
    it("validate success", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(JSON.stringify(pluginManifest) as any);
      sandbox.stub(ManifestUtil, "validateManifest").resolves([]);

      const res = await pluginManifestUtils.validateAgainstSchema(
        { id: "1", file: "file" },
        "testPath",
        driverContext as any
      );
      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.deepEqual(res.value, {
          id: "1",
          filePath: "testPath",
          validationResult: [],
        });
      }
    });

    it("validate action error", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(JSON.stringify(pluginManifest) as any);
      sandbox.stub(ManifestUtil, "validateManifest").resolves([]);
      sandbox
        .stub(pluginManifestUtils, "validateAgainstSchema")
        .resolves(err(new SystemError("error", "error", "error", "error")));

      const res = await pluginManifestUtils.validateAgainstSchema(
        { id: "1", file: "file" },
        "testPath",
        context as any
      );
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal("error", res.error.name);
      }
    });

    it("validate schema error", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(JSON.stringify(pluginManifest) as any);
      sandbox.stub(ManifestUtil, "validateManifest").throws("error");

      const res = await pluginManifestUtils.validateAgainstSchema(
        { id: "1", file: "file" },
        "testPath",
        driverContext as any
      );
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(AppStudioError.ValidationFailedError.name, res.error.name);
      }
    });

    it("error: cannot get manifest", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);

      const res = await pluginManifestUtils.validateAgainstSchema(
        { id: "1", file: "file" },
        "testPath",
        driverContext as any
      );
      chai.assert.isTrue(res.isErr());
    });
  });

  describe("getDefaultNextAvailableApiSpecPath", async () => {
    it("Json file: success on second try", async () => {
      mockedEnvRestore = mockedEnv({
        TEAMSFX_KIOTA_NPM_INTEGRATION: "false",
      });
      sandbox
        .stub(fs, "pathExists")
        .onFirstCall()
        .resolves(true)
        .onSecondCall()
        .resolves(true)
        .onThirdCall()
        .resolves(false);

      const res = await pluginManifestUtils.getDefaultNextAvailableApiSpecPath(
        "testPath.json",
        "test"
      );

      chai.assert.equal(res, path.join("test", "openapi_2.json"));
    });

    it("Json file: success on kiota flag open", async () => {
      mockedEnvRestore = mockedEnv({
        TEAMSFX_KIOTA_NPM_INTEGRATION: "true",
      });
      sandbox
        .stub(fs, "pathExists")
        .onFirstCall()
        .resolves(true)
        .onSecondCall()
        .resolves(true)
        .onThirdCall()
        .resolves(false);

      const res = await pluginManifestUtils.getDefaultNextAvailableApiSpecPath(
        "testPath.json",
        "test"
      );

      chai.assert.equal(res, path.join("test", "openapi_2.yaml"));
    });

    it("Yaml file: success on first try", async () => {
      sandbox.stub(fs, "pathExists").onFirstCall().resolves(true).onSecondCall().resolves(false);

      const res = await pluginManifestUtils.getDefaultNextAvailableApiSpecPath(
        "testPath.yaml",
        "test"
      );

      chai.assert.equal(res, path.join("test", "openapi_1.yaml"));
    });

    it("success on third try with ", async () => {
      mockedEnvRestore = mockedEnv({
        TEAMSFX_KIOTA_NPM_INTEGRATION: "false",
      });
      sandbox.stub(commonUtils, "isJsonSpecFile").throws("fail");
      sandbox
        .stub(fs, "pathExists")
        .onFirstCall()
        .resolves(true)
        .onSecondCall()
        .resolves(true)
        .onThirdCall()
        .resolves(true)
        .onCall(4)
        .resolves(false);

      const res = await pluginManifestUtils.getDefaultNextAvailableApiSpecPath("testPath", "test");

      chai.assert.equal(res, path.join("test", "openapi_3.json"));
    });
  });
});
