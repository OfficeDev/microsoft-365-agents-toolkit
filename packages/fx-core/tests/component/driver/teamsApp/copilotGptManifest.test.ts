// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SpecParser } from "@microsoft/m365-spec-parser";
import {
  CapabilityName,
  Colors,
  DeclarativeAgentManifestConverter,
  DeclarativeAgentManifestWrapper,
  PluginManifestWrapper,
  err,
  Err,
  FxError,
  ManifestUtil,
  ok,
  Ok,
  Platform,
  Result,
  SystemError,
  UserError,
} from "@microsoft/teamsfx-api";
import chai from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import path from "path";
import * as sinon from "sinon";
import { featureFlagManager, FeatureFlags } from "../../../../src";
import { createContext, setTools } from "../../../../src/common/globalVars";
import { generateDriverContext } from "../../../../src/common/utils";
import { EmbeddedKnowledgeLocalDirectoryName } from "../../../../src/component/driver/teamsApp/constants";
import { AppStudioError } from "../../../../src/component/driver/teamsApp/errors";
import { DeclarativeCopilotManifestValidationResult } from "../../../../src/component/driver/teamsApp/interfaces/ValidationResult";
import { copilotGptManifestUtils } from "../../../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import { manifestUtils } from "../../../../src/component/driver/teamsApp/utils/ManifestUtils";
import { pluginManifestUtils } from "../../../../src/component/driver/teamsApp/utils/PluginManifestUtils";
import * as utils from "../../../../src/component/driver/teamsApp/utils/utils";
import { WrapDriverContext } from "../../../../src/component/driver/util/wrapUtil";
import {
  FileNotFoundError,
  MissingEnvironmentVariablesError,
  WriteFileError,
} from "../../../../src/error";
import { MockTools } from "../../../core/utils";
import { MockedLogProvider, MockedTelemetryReporter } from "../../../plugins/solution/util";

describe("copilotGptManifestUtils", () => {
  const sandbox = sinon.createSandbox();
  let mockedEnvRestore: RestoreFn;

  afterEach(async () => {
    sandbox.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });

  const gptManifest = {
    name: "name${{APP_NAME_SUFFIX}}",
    description: "description",
  } as any;

  describe("add plugin", async () => {
    it("add plugin without appending conversation starters success", async () => {
      const manifestData = {
        name: "name${{APP_NAME_SUFFIX}}",
        description: "description",
        conversation_starters: [
          {
            text: "List all repairs",
          },
        ],
      };
      const mockDAWrapper = {
        data: manifestData,
        save: sinon.stub().resolves(),
      };
      const mockPluginWrapper = {
        data: {},
        runtimes: [],
        functions: [],
        save: sinon.stub().resolves(),
      };
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(DeclarativeAgentManifestWrapper, "read").resolves(mockDAWrapper as any);
      sandbox.stub(DeclarativeAgentManifestWrapper, "fromJSON").returns(mockDAWrapper as any);
      sandbox.stub(PluginManifestWrapper, "read").resolves(mockPluginWrapper as any);
      sandbox.stub(path, "dirname").returns("testDir");
      sandbox.stub(path, "join").returns("testDir/testFile");

      const res = await copilotGptManifestUtils.addAction("testPath", "testId", "testFile");

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        const updatedManifest = res.value;
        chai.assert.deepEqual(updatedManifest.actions![0], {
          id: "testId",
          file: "testFile",
        });

        chai.assert.deepEqual(updatedManifest.conversation_starters, [
          {
            text: "List all repairs",
          },
        ]);
      }
    });

    it("add plugin success - use conversation_starters in action file", async () => {
      const manifestData = {
        name: "name${{APP_NAME_SUFFIX}}",
        description: "description",
      };
      const mockDAWrapper = {
        data: manifestData,
        save: sinon.stub().resolves(),
      };
      const mockPluginWrapper = {
        data: {
          capabilities: {
            conversation_starters: [
              {
                text: "List all repairs",
              },
            ],
          },
        },
        runtimes: [],
        functions: [],
      };
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(DeclarativeAgentManifestWrapper, "read").resolves(mockDAWrapper as any);
      sandbox.stub(DeclarativeAgentManifestWrapper, "fromJSON").returns(mockDAWrapper as any);
      sandbox.stub(PluginManifestWrapper, "read").resolves(mockPluginWrapper as any);
      sandbox.stub(path, "dirname").returns("testDir");
      sandbox.stub(path, "join").returns("testDir/testFile");

      const res = await copilotGptManifestUtils.addAction("testPath", "testId", "testFile");

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        const updatedManifest = res.value;
        chai.assert.deepEqual(updatedManifest.actions![0], {
          id: "testId",
          file: "testFile",
        });

        chai.assert.deepEqual(updatedManifest.conversation_starters, [
          {
            text: "List all repairs",
          },
        ]);
      }
    });

    it("add plugin success - parse conversation_starters in open api spec file", async () => {
      const manifestData = {
        name: "name${{APP_NAME_SUFFIX}}",
        description: "description",
      };
      const mockDAWrapper = {
        data: manifestData,
        save: sinon.stub().resolves(),
      };
      const mockPluginWrapper = {
        data: {
          capabilities: {
            conversation_starters: [],
          },
          runtimes: [
            {
              type: "OpenApi",
              auth: {
                type: "None",
              },
              spec: {
                url: "apiSpecificationFile/openapi.json",
              },
              run_for_functions: ["deleteRepairs", "listRepairs", "patchRepairs", "createRepair"],
            },
          ],
          functions: [
            {
              name: "listRepairs",
              description: "List all repairs",
            },
            {
              name: "deleteRepairs",
              description: "Delete repairs",
            },
            {
              name: "patchRepairs",
              description: "Patch repairs",
            },
            {
              name: "createRepair",
              description: "Create repairs",
            },
          ],
        },
        runtimes: [
          {
            type: "OpenApi",
            auth: {
              type: "None",
            },
            spec: {
              url: "apiSpecificationFile/openapi.json",
            },
            run_for_functions: ["deleteRepairs", "listRepairs", "patchRepairs", "createRepair"],
          },
        ],
        functions: [
          {
            name: "listRepairs",
            description: "List all repairs",
          },
          {
            name: "deleteRepairs",
            description: "Delete repairs",
          },
          {
            name: "patchRepairs",
            description: "Patch repairs",
          },
          {
            name: "createRepair",
            description: "Create repairs",
          },
        ],
      };
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox
        .stub(featureFlagManager, "getBooleanValue")
        .withArgs(FeatureFlags.KiotaNPMIntegration)
        .returns(false);
      sandbox.stub(DeclarativeAgentManifestWrapper, "read").resolves(mockDAWrapper as any);
      sandbox.stub(DeclarativeAgentManifestWrapper, "fromJSON").returns(mockDAWrapper as any);
      sandbox.stub(PluginManifestWrapper, "read").resolves(mockPluginWrapper as any);
      sandbox.stub(path, "dirname").returns("testDir");
      sandbox.stub(path, "join").returns("testDir/testFile");
      sandbox.stub(path, "resolve").returns("testDir/testFile");

      sandbox.stub(SpecParser.prototype, "list").resolves({
        APIs: [
          {
            api: "GET /repairs",
            server: "https://example.com",
            operationId: "listRepairs",
            isValid: true,
            reason: [],
            summary: "List all repairs",
            description: "It is used to list all repairs",
          },
          {
            api: "POST /repairs",
            server: "https://example.com",
            operationId: "createRepair",
            isValid: true,
            reason: [],
          },
          {
            api: "DELETE /repairs",
            server: "https://example.com",
            operationId: "deleteRepairs",
            isValid: true,
            reason: [],
            description: "It is used to delete a repair",
          },
          {
            api: "PATCH /repairs",
            server: "https://example.com",
            operationId: "patchRepairs",
            isValid: false,
            reason: [],
            summary: "Patch a repair",
            description: "It is used to patch a repair",
          },
          {
            api: "Put /repairs",
            server: "https://example.com",
            operationId: "putRepairs",
            isValid: true,
            reason: [],
            summary: "Put a repair",
            description: "It is used to put a repair",
          },
        ],
        allAPICount: 1,
        validAPICount: 1,
      });

      const res = await copilotGptManifestUtils.addAction("testPath", "testId", "testFile");

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        const updatedManifest = res.value;
        chai.assert.deepEqual(updatedManifest.actions![0], {
          id: "testId",
          file: "testFile",
        });

        chai.assert.deepEqual(updatedManifest.conversation_starters, [
          {
            text: "List all repairs",
          },
          {
            text: "It is used to delete a repair",
          },
        ]);
      }
    });

    it("add plugin success - parse conversation_starters in open api spec file with undefined existing conversation starter", async () => {
      const manifestData = {
        name: "name${{APP_NAME_SUFFIX}}",
        description: "description",
      };
      const mockDAWrapper = {
        data: manifestData,
        save: sinon.stub().resolves(),
      };
      const mockPluginWrapper = {
        data: {
          capabilities: {},
          runtimes: [
            {
              type: "OpenApi",
              auth: {
                type: "None",
              },
              spec: {
                url: "apiSpecificationFile/openapi.json",
              },
              run_for_functions: ["listRepairs"],
            },
          ],
          functions: [
            {
              name: "listRepairs",
              description: "List all repairs",
            },
          ],
        },
        runtimes: [
          {
            type: "OpenApi",
            auth: {
              type: "None",
            },
            spec: {
              url: "apiSpecificationFile/openapi.json",
            },
            run_for_functions: ["listRepairs"],
          },
        ],
        functions: [
          {
            name: "listRepairs",
            description: "List all repairs",
          },
        ],
      };
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(DeclarativeAgentManifestWrapper, "read").resolves(mockDAWrapper as any);
      sandbox.stub(DeclarativeAgentManifestWrapper, "fromJSON").returns(mockDAWrapper as any);
      sandbox.stub(PluginManifestWrapper, "read").resolves(mockPluginWrapper as any);
      sandbox
        .stub(featureFlagManager, "getBooleanValue")
        .withArgs(FeatureFlags.KiotaNPMIntegration)
        .returns(false);
      sandbox.stub(path, "dirname").returns("testDir");
      sandbox.stub(path, "join").returns("testDir/testFile");
      sandbox.stub(path, "resolve").returns("testDir/testFile");

      sandbox.stub(SpecParser.prototype, "list").resolves({
        APIs: [
          {
            api: "GET /repairs",
            server: "https://example.com",
            operationId: "listRepairs",
            isValid: true,
            reason: [],
            summary: "List all repairs",
          },
        ],
        allAPICount: 1,
        validAPICount: 1,
      });

      const res = await copilotGptManifestUtils.addAction("testPath", "testId", "testFile");

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        const updatedManifest = res.value;
        chai.assert.deepEqual(updatedManifest.actions![0], {
          id: "testId",
          file: "testFile",
        });

        chai.assert.deepEqual(updatedManifest.conversation_starters, [
          {
            text: "List all repairs",
          },
        ]);
      }
    });

    it("add plugin and append conversation starters success - use conversation_starters in action file", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      const mockDAWrapper = {
        data: {
          name: "name${{APP_NAME_SUFFIX}}",
          description: "description",
          conversation_starters: [
            {
              text: "List all repairs1",
            },
          ],
          actions: [
            {
              id: "action_1",
              file: "plugin1.json",
            },
          ],
        },
        save: sinon.stub().resolves(),
      };
      const mockPluginWrapper = {
        data: {
          capabilities: {
            conversation_starters: [
              {
                text: "List all repairs2",
              },
            ],
          },
        },
        runtimes: [],
        functions: [],
      };
      sandbox.stub(DeclarativeAgentManifestWrapper, "read").resolves(mockDAWrapper as any);
      sandbox.stub(DeclarativeAgentManifestWrapper, "fromJSON").returns(mockDAWrapper as any);
      sandbox.stub(PluginManifestWrapper, "read").resolves(mockPluginWrapper as any);
      sandbox.stub(path, "dirname").returns("testDir");
      sandbox.stub(path, "join").returns("testDir/testFile");

      const res = await copilotGptManifestUtils.addAction("testPath", "testId", "testFile");

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        const updatedManifest = res.value;
        chai.assert.deepEqual(updatedManifest.actions![0], {
          id: "action_1",
          file: "plugin1.json",
        });
        chai.assert.deepEqual(updatedManifest.actions![1], {
          id: "testId",
          file: "testFile",
        });

        chai.assert.deepEqual(updatedManifest.conversation_starters, [
          {
            text: "List all repairs1",
          },
          {
            text: "List all repairs2",
          },
        ]);
      }
    });

    it("conversation starters count should less than 6", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      const mockDAWrapper = {
        data: {
          name: "name${{APP_NAME_SUFFIX}}",
          description: "description",
          conversation_starters: [
            {
              text: "List all repairs1",
            },
          ],
          actions: [
            {
              id: "action_1",
              file: "plugin1.json",
            },
          ],
        },
        save: sinon.stub().resolves(),
      };
      const mockPluginWrapper = {
        data: {
          capabilities: {
            conversation_starters: [
              {
                text: "List all repairs2",
              },
              {
                text: "List all repairs3",
              },
              {
                text: "List all repairs4",
              },
              {
                text: "List all repairs5",
              },
              {
                text: "List all repairs6",
              },
              {
                text: "List all repairs7",
              },
            ],
          },
        },
        runtimes: [],
        functions: [],
      };
      sandbox.stub(DeclarativeAgentManifestWrapper, "read").resolves(mockDAWrapper as any);
      sandbox.stub(DeclarativeAgentManifestWrapper, "fromJSON").returns(mockDAWrapper as any);
      sandbox.stub(PluginManifestWrapper, "read").resolves(mockPluginWrapper as any);
      sandbox.stub(path, "dirname").returns("testDir");
      sandbox.stub(path, "join").returns("testDir/testFile");

      const res = await copilotGptManifestUtils.addAction("testPath", "testId", "testFile");

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        const updatedManifest = res.value;
        chai.assert.deepEqual(updatedManifest.actions![0], {
          id: "action_1",
          file: "plugin1.json",
        });
        chai.assert.deepEqual(updatedManifest.actions![1], {
          id: "testId",
          file: "testFile",
        });

        chai.assert.deepEqual(updatedManifest.conversation_starters, [
          {
            text: "List all repairs1",
          },
          {
            text: "List all repairs2",
          },
          {
            text: "List all repairs3",
          },
          {
            text: "List all repairs4",
          },
          {
            text: "List all repairs5",
          },
          {
            text: "List all repairs6",
          },
        ]);
      }
    });

    it("conversation starters should unique", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      const mockDAWrapper = {
        data: {
          name: "name${{APP_NAME_SUFFIX}}",
          description: "description",
          conversation_starters: [
            {
              text: "List all repairs1",
            },
          ],
          actions: [
            {
              id: "action_1",
              file: "plugin1.json",
            },
          ],
        },
        save: sinon.stub().resolves(),
      };
      const mockPluginWrapper = {
        data: {
          capabilities: {
            conversation_starters: [
              {
                text: "List all repairs1",
              },
              {
                text: "List all repairs2",
              },
            ],
          },
        },
        runtimes: [],
        functions: [],
      };
      sandbox.stub(DeclarativeAgentManifestWrapper, "read").resolves(mockDAWrapper as any);
      sandbox.stub(DeclarativeAgentManifestWrapper, "fromJSON").returns(mockDAWrapper as any);
      sandbox.stub(PluginManifestWrapper, "read").resolves(mockPluginWrapper as any);
      sandbox.stub(path, "dirname").returns("testDir");
      sandbox.stub(path, "join").returns("testDir/testFile");

      const res = await copilotGptManifestUtils.addAction("testPath", "testId", "testFile");

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        const updatedManifest = res.value;
        chai.assert.deepEqual(updatedManifest.actions![0], {
          id: "action_1",
          file: "plugin1.json",
        });
        chai.assert.deepEqual(updatedManifest.actions![1], {
          id: "testId",
          file: "testFile",
        });

        chai.assert.deepEqual(updatedManifest.conversation_starters, [
          {
            text: "List all repairs1",
          },
          {
            text: "List all repairs2",
          },
        ]);
      }
    });

    it("add plugin error: read manifest error", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      const res = await copilotGptManifestUtils.addAction("testPath", "testId", "testFile");
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof FileNotFoundError);
      }
    });

    it("add plugin error: write file error", async () => {
      const mockDAWrapper = {
        data: gptManifest,
      };
      const mockPluginWrapper = {
        data: {},
        runtimes: [],
        functions: [],
      };
      const mockSaveWrapper = {
        data: gptManifest,
        save: sinon.stub().throws("some error"),
      };
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(DeclarativeAgentManifestWrapper, "read").resolves(mockDAWrapper as any);
      sandbox.stub(DeclarativeAgentManifestWrapper, "fromJSON").returns(mockSaveWrapper as any);
      sandbox.stub(PluginManifestWrapper, "read").resolves(mockPluginWrapper as any);
      sandbox.stub(path, "dirname").returns("testDir");
      sandbox.stub(path, "join").returns("testDir/testFile");
      const res = await copilotGptManifestUtils.addAction("testPath", "testId", "testFile");
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof WriteFileError);
      }
    });
  });

  describe("getManifest", async () => {
    setTools(new MockTools());
    const context = generateDriverContext(createContext(), {
      platform: Platform.VSCode,
      projectPath: "",
    });
    const mockedContex = new WrapDriverContext(context, "test", "test");
    it("get manifest success", async () => {
      mockedEnvRestore = mockedEnv({
        ["APP_NAME_SUFFIX"]: "test",
      });
      const manifestData = {
        name: "name${{APP_NAME_SUFFIX}}",
        description: "description",
      };
      const mockDAWrapper = {
        data: manifestData,
      };
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(DeclarativeAgentManifestWrapper, "read").resolves(mockDAWrapper as any);
      sandbox.stub(utils, "getResolvedManifest").resolves(
        ok(
          JSON.stringify({
            name: "nametest",
            description: "description",
          })
        )
      );

      const res = await copilotGptManifestUtils.getManifest("testPath", mockedContex);

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.equal("nametest", res.value.name);
      }
    });

    it("get manifest error: file not found", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);
      const res = await copilotGptManifestUtils.getManifest("testPath", mockedContex);
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof FileNotFoundError);
      }
    });

    it("get manifest error: unresolved env error", async () => {
      const manifestData = {
        name: "name${{APP_NAME_SUFFIX}}",
        description: "description",
      };
      const mockDAWrapper = {
        data: manifestData,
      };
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(DeclarativeAgentManifestWrapper, "read").resolves(mockDAWrapper as any);
      sandbox
        .stub(utils, "getResolvedManifest")
        .resolves(
          err(new MissingEnvironmentVariablesError("test", "APP_NAME_SUFFIX", "test", "test"))
        );

      const res = await copilotGptManifestUtils.getManifest("testPath", mockedContex);

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
      const manifest = {
        ...gptManifest,
        actions: [
          {
            id: "1",
            file: "testFile",
          },
        ],
      };
      mockedEnvRestore = mockedEnv({
        ["APP_NAME_SUFFIX"]: "test",
      });
      const mockDAWrapper = {
        data: manifest,
      };
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(DeclarativeAgentManifestWrapper, "read").resolves(mockDAWrapper as any);
      sandbox.stub(utils, "getResolvedManifest").resolves(ok(JSON.stringify(manifest)));
      sandbox.stub(ManifestUtil, "validateManifest").resolves([]);
      sandbox.stub(pluginManifestUtils, "validateAgainstSchema").resolves(
        ok({
          id: "1",
          filePath: "testFile",
          validationResult: ["error1"],
        })
      );
      sandbox.stub(path, "dirname").returns("testDir");
      sandbox.stub(path, "join").returns("testFile");

      const res = await copilotGptManifestUtils.validateAgainstSchema(
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
          actionValidationResult: [
            {
              id: "1",
              filePath: "testFile",
              validationResult: ["error1"],
            },
          ],
        });
      }
    });

    it("validate action error", async () => {
      const manifest = {
        ...gptManifest,
        actions: [
          {
            id: "1",
            file: "testFile",
          },
        ],
      };
      mockedEnvRestore = mockedEnv({
        ["APP_NAME_SUFFIX"]: "test",
      });
      const mockDAWrapper = {
        data: manifest,
      };
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(DeclarativeAgentManifestWrapper, "read").resolves(mockDAWrapper as any);
      sandbox.stub(utils, "getResolvedManifest").resolves(ok(JSON.stringify(manifest)));
      sandbox.stub(ManifestUtil, "validateManifest").resolves([]);
      sandbox
        .stub(pluginManifestUtils, "validateAgainstSchema")
        .resolves(err(new SystemError("error", "error", "error", "error")));
      sandbox.stub(path, "dirname").returns("testDir");
      sandbox.stub(path, "join").returns("testFile");

      const res = await copilotGptManifestUtils.validateAgainstSchema(
        { id: "1", file: "file" },
        "testPath",
        driverContext as any
      );
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal("error", res.error.name);
      }
    });

    it("validate schema error", async () => {
      const manifest = {
        ...gptManifest,
        actions: [
          {
            id: "1",
            file: "testFile",
          },
        ],
      };
      mockedEnvRestore = mockedEnv({
        ["APP_NAME_SUFFIX"]: "test",
      });
      const mockDAWrapper = {
        data: manifest,
      };
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(DeclarativeAgentManifestWrapper, "read").resolves(mockDAWrapper as any);
      sandbox.stub(utils, "getResolvedManifest").resolves(ok(JSON.stringify(manifest)));
      sandbox.stub(ManifestUtil, "validateManifest").throws("error");

      const res = await copilotGptManifestUtils.validateAgainstSchema(
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

      const res = await copilotGptManifestUtils.validateAgainstSchema(
        { id: "1", file: "file" },
        "testPath",
        driverContext as any
      );
      chai.assert.isTrue(res.isErr());
    });
  });

  describe("logValidationErrors", () => {
    it("skip if no errors", () => {
      const validationRes: DeclarativeCopilotManifestValidationResult = {
        id: "1",
        filePath: "testPath",
        validationResult: [],
        actionValidationResult: [
          {
            id: "1",
            filePath: "testPath",
            validationResult: [],
          },
          {
            id: "2",
            filePath: "pluginPath",
            validationResult: [],
          },
        ],
      };

      const res = copilotGptManifestUtils.logValidationErrors(validationRes, Platform.VSCode);
      chai.assert.isEmpty(res);
    });

    it("log if VSC", () => {
      const validationRes: DeclarativeCopilotManifestValidationResult = {
        id: "1",
        filePath: "testPath",
        validationResult: ["error1"],
        actionValidationResult: [
          {
            id: "1",
            filePath: "testPath",
            validationResult: ["errorAction1"],
          },
          {
            id: "2",
            filePath: "pluginPath",
            validationResult: ["errorAction2"],
          },
        ],
      };

      const res = copilotGptManifestUtils.logValidationErrors(
        validationRes,
        Platform.VSCode
      ) as string;

      chai.assert.isFalse(res.includes("errorActions2"));
      chai.assert.isTrue(res.includes("error1"));
      chai.assert.isTrue(res.includes("errorAction1"));
    });

    it("log if VSC and action error only", () => {
      const validationRes: DeclarativeCopilotManifestValidationResult = {
        id: "1",
        filePath: "testPath",
        validationResult: [],
        actionValidationResult: [
          {
            id: "1",
            filePath: "testPath",
            validationResult: ["errorAction1"],
          },
          {
            id: "2",
            filePath: "pluginPath",
            validationResult: ["errorAction2"],
          },
        ],
      };

      const res = copilotGptManifestUtils.logValidationErrors(
        validationRes,
        Platform.VSCode
      ) as string;

      chai.assert.isFalse(res.includes("errorActions2"));
      chai.assert.isTrue(res.includes("errorAction1"));
    });

    it("log if CLI", () => {
      const validationRes: DeclarativeCopilotManifestValidationResult = {
        id: "1",
        filePath: "testPath",
        validationResult: ["error1"],
        actionValidationResult: [
          {
            id: "1",
            filePath: "testPath",
            validationResult: ["errorAction1"],
          },
        ],
      };

      const res = copilotGptManifestUtils.logValidationErrors(
        validationRes,
        Platform.CLI
      ) as Array<{ content: string; color: Colors }>;
      chai.assert.isTrue(res.find((item) => item.content.includes("error1")) !== undefined);
      chai.assert.isTrue(res.find((item) => item.content.includes("errorAction1")) !== undefined);
      chai.assert.isUndefined(res.find((item) => item.content.includes("errorAction2")));
    });

    it("log if CLI and action error only", () => {
      const validationRes: DeclarativeCopilotManifestValidationResult = {
        id: "1",
        filePath: "testPath",
        validationResult: [],
        actionValidationResult: [
          {
            id: "1",
            filePath: "testPath",
            validationResult: ["errorAction1"],
          },
          {
            id: "2",
            filePath: "pluginPath",
            validationResult: ["errorAction2"],
          },
        ],
      };

      const res = copilotGptManifestUtils.logValidationErrors(
        validationRes,
        Platform.CLI
      ) as Array<{ content: string; color: Colors }>;
      chai.assert.isTrue(res.find((item) => item.content.includes("errorAction2")) !== undefined);
      chai.assert.isTrue(res.find((item) => item.content.includes("errorAction1")) !== undefined);
    });
  });

  describe("getManifestPath", async () => {
    setTools(new MockTools());
    const context = generateDriverContext(createContext(), {
      platform: Platform.VSCode,
      projectPath: "",
    });

    it("get manifest success", async () => {
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(
        ok({
          copilotExtensions: {
            declarativeCopilots: [
              {
                file: "test",
                id: "1",
              },
            ],
          },
        } as any)
      );
      sandbox.stub(path, "dirname").returns("testFolder");
      sandbox.stub(path, "resolve").returns("testFolder/test");

      const res = await copilotGptManifestUtils.getManifestPath("testPath");

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.equal(res.value, "testFolder/test");
      }
    });

    it("get manifest success - copilot agent", async () => {
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(
        ok({
          copilotAgents: {
            declarativeAgents: [
              {
                file: "test",
                id: "1",
              },
            ],
          },
        } as any)
      );
      sandbox.stub(path, "dirname").returns("testFolder");
      sandbox.stub(path, "resolve").returns("testFolder/test");

      const res = await copilotGptManifestUtils.getManifestPath("testPath");

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.equal(res.value, "testFolder/test");
      }
    });

    it("declarativeAgents error 1", async () => {
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(
        ok({
          copilotAgents: {},
        } as any)
      );
      const res = await copilotGptManifestUtils.getManifestPath("testPath");
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof UserError);
      }
    });

    it("declarativeAgents error 2", async () => {
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok({} as any));
      const res = await copilotGptManifestUtils.getManifestPath("testPath");
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof UserError);
      }
    });

    it("declarativeCopilots error 1", async () => {
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(
        ok({
          copilotExtensions: {},
        } as any)
      );
      const res = await copilotGptManifestUtils.getManifestPath("testPath");
      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof UserError);
      }
    });

    it("read Teams manifest error", async () => {
      sandbox
        .stub(manifestUtils, "_readAppManifest")
        .resolves(err(new UserError("readError", "readError", "", "")));

      const res = await copilotGptManifestUtils.getManifestPath("testPath");

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.name, "readError");
      }
    });

    it("missing file property", async () => {
      sandbox.stub(manifestUtils, "_readAppManifest").resolves(
        ok({
          copilotExtensions: {
            declarativeCopilots: [
              {
                id: "1",
              },
            ],
          },
        } as any)
      );

      const res = await copilotGptManifestUtils.getManifestPath("testPath");

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error.name, AppStudioError.TeamsAppRequiredPropertyMissingError.name);
      }
    });
  });

  describe("getDefaultNextAvailablePluginManifestPath", async () => {
    setTools(new MockTools());
    const context = generateDriverContext(createContext(), {
      platform: Platform.VSCode,
      projectPath: "",
    });

    it("Success on second try", async () => {
      sandbox
        .stub(fs, "pathExists")
        .onFirstCall()
        .resolves(true)
        .onSecondCall()
        .resolves(true)
        .onThirdCall()
        .resolves(false);
      const res = await copilotGptManifestUtils.getDefaultNextAvailablePluginManifestPath("test");
      chai.assert.equal(res, path.join("test", "ai-plugin_2.json"));
    });

    it("Success on first try", async () => {
      sandbox.stub(fs, "pathExists").onFirstCall().resolves(true).onSecondCall().resolves(false);
      const res = await copilotGptManifestUtils.getDefaultNextAvailablePluginManifestPath("test");
      chai.assert.equal(res, path.join("test", "ai-plugin_1.json"));
    });
  });

  describe("add knowledge for Copilot connector", async () => {
    setTools(new MockTools());
    const context = generateDriverContext(createContext(), {
      platform: Platform.VSCode,
      projectPath: "",
    });

    const agentManifestPath = "test/agentManifestPath";
    let manifestRes: Result<any, FxError>;

    it("happy path", async () => {
      sandbox
        .stub(copilotGptManifestUtils, "writeCopilotGptManifestFile")
        .resolves(new Ok(undefined));
      const connectionIds = ["connectionId1", "connectionId2"];
      const manifest = {
        version: "v1.6" as const,
        name: "name${{APP_NAME_SUFFIX}}",
        description: "description",
        capabilities: [
          {
            name: CapabilityName.GraphConnectors,
            connections: [
              {
                connection_id: "123",
              },
            ],
          },
        ],
      };
      const res = await copilotGptManifestUtils.addGCCapability(
        agentManifestPath,
        connectionIds,
        new Ok(manifest)
      );
      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.deepEqual(manifest, {
          version: "v1.6",
          name: "name${{APP_NAME_SUFFIX}}",
          description: "description",
          capabilities: [
            {
              name: CapabilityName.GraphConnectors,
              connections: [
                {
                  connection_id: "123",
                },
                {
                  connection_id: "connectionId1",
                },
                {
                  connection_id: "connectionId2",
                },
              ],
            },
          ],
        });
      }
    });

    it("duplicated id", async () => {
      sandbox
        .stub(copilotGptManifestUtils, "writeCopilotGptManifestFile")
        .resolves(new Ok(undefined));
      const connectionIds = ["123"];
      const manifest = {
        version: "v1.6" as const,
        name: "name${{APP_NAME_SUFFIX}}",
        description: "description",
        capabilities: [
          {
            name: CapabilityName.GraphConnectors,
            connections: [
              {
                connection_id: "123",
              },
            ],
          },
        ],
      } as any;
      const res = await copilotGptManifestUtils.addGCCapability(
        agentManifestPath,
        connectionIds,
        new Ok(manifest)
      );
      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.deepEqual(manifest, {
          version: "v1.6",
          name: "name${{APP_NAME_SUFFIX}}",
          description: "description",
          capabilities: [
            {
              name: CapabilityName.GraphConnectors,
              connections: [
                {
                  connection_id: "123",
                },
              ],
            },
          ],
        });
      }
    });

    it("empty capability", async () => {
      sandbox
        .stub(copilotGptManifestUtils, "writeCopilotGptManifestFile")
        .resolves(new Ok(undefined));
      const connectionIds = ["123"];
      const manifest = {
        version: "v1.6" as const,
        name: "name${{APP_NAME_SUFFIX}}",
        description: "description",
      } as any;
      const res = await copilotGptManifestUtils.addGCCapability(
        agentManifestPath,
        connectionIds,
        new Ok(manifest)
      );
      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.deepEqual(manifest, {
          version: "v1.6",
          name: "name${{APP_NAME_SUFFIX}}",
          description: "description",
          capabilities: [
            {
              name: CapabilityName.GraphConnectors,
              connections: [
                {
                  connection_id: "123",
                },
              ],
            },
          ],
        });
      }
    });

    it("error manifest", async () => {
      const res = await copilotGptManifestUtils.addGCCapability(
        agentManifestPath,
        ["123"],
        new Err("manifest") as any
      );
      chai.assert.isFalse(res.isOk());
    });
  });

  describe("addEmbeddedKnowledgeFiles", () => {
    setTools(new MockTools());
    const context = generateDriverContext(createContext(), {
      platform: Platform.VSCode,
      projectPath: "",
    });
    it("should add embedded knowledge files successfully - empty declarative agent manifest", async () => {
      const manifestFilePath = "test/manifest.json";
      const resolvedManifestPath = "test/resolvedManifest.json";
      const manifest: any = {};

      sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok(resolvedManifestPath));
      sandbox.stub(copilotGptManifestUtils, "readCopilotGptManifestFile").resolves(ok(manifest));
      const writeStub = sandbox
        .stub(copilotGptManifestUtils, "writeCopilotGptManifestFile")
        .resolves(ok(undefined));
      const ensureDirStub = sandbox.stub(fs, "ensureDir").resolves();
      const copyFileStub = sandbox.stub(fs, "copyFile").resolves();

      const filePathList = ["dummy.txt"];
      const result = await copilotGptManifestUtils.addEmbeddedKnowledgeFiles(
        manifestFilePath,
        filePathList
      );
      chai.assert.isTrue(result.isOk());

      const expectedDir = path.resolve(
        path.dirname(manifestFilePath),
        EmbeddedKnowledgeLocalDirectoryName
      );
      sinon.assert.calledWith(ensureDirStub, expectedDir);

      const expectedSavedPath = path.resolve(
        path.dirname(manifestFilePath),
        EmbeddedKnowledgeLocalDirectoryName,
        path.basename("dummy.txt")
      );

      chai.assert.isArray(manifest.capabilities);
      const capability = manifest.capabilities.find(
        (cap: any) => cap.name === CapabilityName.EmbeddedKnowledge
      );
      chai.assert.exists(capability);
      chai.assert.isArray(capability.files);
      const expectedRelativePath = path
        .relative(path.dirname(manifestFilePath), expectedSavedPath)
        .replace(/\\/g, "/");
      chai.assert.equal(capability.files[0].file, expectedRelativePath);

      sinon.assert.calledWith(writeStub, manifest, resolvedManifestPath);
    });

    it("should add embedded knowledge files successfully - declarative agent manifest with knowledge", async () => {
      const manifestFilePath = "test/manifest.json";
      const resolvedManifestPath = "test/resolvedManifest.json";
      const manifest: any = {
        capabilities: [
          {
            name: CapabilityName.EmbeddedKnowledge,
            files: [{ file: "existing.txt" }],
          },
        ],
      };

      sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok(resolvedManifestPath));
      sandbox.stub(copilotGptManifestUtils, "readCopilotGptManifestFile").resolves(ok(manifest));
      const writeStub = sandbox
        .stub(copilotGptManifestUtils, "writeCopilotGptManifestFile")
        .resolves(ok(undefined));
      const ensureDirStub = sandbox.stub(fs, "ensureDir").resolves();
      const copyFileStub = sandbox.stub(fs, "copyFile").resolves();

      const filePathList = ["dummy.txt"];
      const result = await copilotGptManifestUtils.addEmbeddedKnowledgeFiles(
        manifestFilePath,
        filePathList
      );
      chai.assert.isTrue(result.isOk());

      const expectedDir = path.resolve(
        path.dirname(manifestFilePath),
        EmbeddedKnowledgeLocalDirectoryName
      );
      sinon.assert.calledWith(ensureDirStub, expectedDir);

      const expectedSavedPath = path.resolve(
        path.dirname(manifestFilePath),
        EmbeddedKnowledgeLocalDirectoryName,
        path.basename("dummy.txt")
      );

      chai.assert.isArray(manifest.capabilities);
      const capability = manifest.capabilities.find(
        (cap: any) => cap.name === CapabilityName.EmbeddedKnowledge
      );
      chai.assert.exists(capability);
      chai.assert.isArray(capability.files);
      chai.assert.equal(capability.files.length, 2);
      chai.assert.isTrue(capability.files[0].file == "existing.txt");
      const expectedRelativePath = path
        .relative(path.dirname(manifestFilePath), expectedSavedPath)
        .replace(/\\/g, "/");
      chai.assert.equal(capability.files[1].file, expectedRelativePath);
      sinon.assert.calledWith(writeStub, manifest, resolvedManifestPath);
    });

    it("should return error if getManifestPath fails", async () => {
      const fackeErr = new SystemError("FakeError", "getManifestPath failed", "test", "");
      sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(err(fackeErr));
      const result = await copilotGptManifestUtils.addEmbeddedKnowledgeFiles("test/manifest.json", [
        "dummy.txt",
      ]);
      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert.deepEqual(result.error, fackeErr);
      }
    });
    it("should return error if readCopilotGptManifestFile fails", async () => {
      const resolvedManifestPath = "test/resolvedManifest.json";
      sandbox.stub(copilotGptManifestUtils, "getManifestPath").resolves(ok(resolvedManifestPath));
      const fackeErr = new SystemError(
        "FakeError",
        "readCopilotGptManifestFile failed",
        "test",
        ""
      );
      sandbox.stub(copilotGptManifestUtils, "readCopilotGptManifestFile").resolves(err(fackeErr));
      const result = await copilotGptManifestUtils.addEmbeddedKnowledgeFiles("test/manifest.json", [
        "dummy.txt",
      ]);
      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert.deepEqual(result.error, fackeErr);
      }
    });
  });

  describe("add knowledge for Web Content and OneDrive and Sharepoint", () => {
    setTools(new MockTools());
    afterEach(async () => {
      if (await fs.pathExists("fake agent manifest path")) {
        await fs.unlink("fake agent manifest path");
      }
      sandbox.restore();
    });

    it("happy path: manifestRes has no capabilities for addOrUpdateCapability ", async () => {
      const agentManifest = {
        name: "name${{APP_NAME_SUFFIX}}",
        description: "description",
      } as any;
      const res = await copilotGptManifestUtils.addOrUpdateCapability(
        "fake agent manifest path",
        CapabilityName.WebSearch,
        ok(agentManifest),
        {}
      );
      chai.assert.isTrue(res.isOk());
    });

    it("error path: manifestRes is error", async () => {
      let res = await copilotGptManifestUtils.addOneDriveSharePointCapability(
        "fake agent manifest path",
        null,
        null,
        err(new UserError("fake error", "fake error", "fake error", "fake error"))
      );
      chai.assert.isTrue(res.isErr());

      res = await copilotGptManifestUtils.addWebSearchCapability(
        context,
        "fake agent manifest path",
        null,
        err(new UserError("fake error", "fake error", "fake error", "fake error"))
      );
      chai.assert.isTrue(res.isErr());
    });

    it("error path: manifestRes error for addOrUpdateCapability ", async () => {
      const res = await copilotGptManifestUtils.addOrUpdateCapability(
        "fake agent manifest path",
        CapabilityName.WebSearch,
        err(new UserError("fake error", "fake error", "fake error", "fake error")),
        {}
      );
      chai.assert.isTrue(res.isErr());
    });

    it("error path: updateGptManifestRes error for addOrUpdateCapability ", async () => {
      const agentManifest = {
        name: "name${{APP_NAME_SUFFIX}}",
        description: "description",
        capabilities: [
          {
            name: CapabilityName.WebSearch,
          },
        ],
      } as any;

      sandbox
        .stub(copilotGptManifestUtils, "writeCopilotGptManifestFile")
        .resolves(err(new UserError("fake error", "fake error", "fake error", "fake error")));
      const res = await copilotGptManifestUtils.addOrUpdateCapability(
        "fake agent manifest path",
        CapabilityName.WebSearch,
        ok(agentManifest),
        {}
      );
      chai.assert.isTrue(res.isErr());
    });
  });

  describe("readCopilotGptManifestFileSync", () => {
    it("should read manifest file successfully", () => {
      const manifest = {
        name: "name${{APP_NAME_SUFFIX}}",
        description: "description",
      };
      const mockDAWrapper = {
        data: manifest,
      };
      sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(DeclarativeAgentManifestWrapper, "readSync").returns(mockDAWrapper as any);

      const res = copilotGptManifestUtils.readCopilotGptManifestFileSync("testPath");

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.deepEqual(res.value, manifest);
      }
    });

    it("should return FileNotFoundError if file does not exist", () => {
      sandbox.stub(DeclarativeAgentManifestWrapper, "readSync").throws(new Error("File not found"));

      const res = copilotGptManifestUtils.readCopilotGptManifestFileSync("testPath");

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof FileNotFoundError);
      }
    });

    it("should return FileNotFoundError if JSON parse fails", () => {
      sandbox
        .stub(DeclarativeAgentManifestWrapper, "readSync")
        .throws(new SyntaxError("Invalid JSON"));

      const res = copilotGptManifestUtils.readCopilotGptManifestFileSync("testPath");

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof FileNotFoundError);
      }
    });
  });

  describe("readDeclarativeAgentManifestFile", async () => {
    it("should read manifest file successfully", async () => {
      const manifest = {
        schema: "schema",
        description: "description",
      } as any;
      const mockDAWrapper = {
        data: manifest,
      };
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(DeclarativeAgentManifestWrapper, "read").resolves(mockDAWrapper as any);

      const res = await copilotGptManifestUtils.readDeclarativeAgentManifestFile("testPath");

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.deepEqual(res.value, manifest);
      }
    });

    it("should return FileNotFoundError if file does not exist", async () => {
      sandbox.stub(fs, "pathExists").resolves(false);

      const res = await copilotGptManifestUtils.readDeclarativeAgentManifestFile("testPath");

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof FileNotFoundError);
      }
    });

    it("should return JSONSyntaxError if manifest conversion fails", async () => {
      sandbox
        .stub(DeclarativeAgentManifestWrapper, "read")
        .rejects(new SyntaxError("Invalid JSON"));

      const res = await copilotGptManifestUtils.readDeclarativeAgentManifestFile("testPath");

      chai.assert.isTrue(res.isErr());
    });
  });

  describe("readDeclarativeAgentManifestFileSync", () => {
    it("should read manifest file successfully", () => {
      const manifest = {
        schema: "schema",
        description: "description",
      } as any;
      const mockDAWrapper = {
        data: manifest,
      };
      sandbox.stub(fs, "pathExistsSync").returns(true);
      sandbox.stub(DeclarativeAgentManifestWrapper, "readSync").returns(mockDAWrapper as any);

      const res = copilotGptManifestUtils.readDeclarativeAgentManifestFileSync("testPath");

      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.deepEqual(res.value, manifest);
      }
    });

    it("should return FileNotFoundError if file does not exist", () => {
      sandbox.stub(fs, "pathExistsSync").returns(false);

      const res = copilotGptManifestUtils.readDeclarativeAgentManifestFileSync("testPath");

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof FileNotFoundError);
      }
    });

    it("should return JSONSyntaxError if manifest conversion fails", () => {
      sandbox
        .stub(DeclarativeAgentManifestWrapper, "readSync")
        .throws(new SyntaxError("Invalid JSON"));

      const res = copilotGptManifestUtils.readDeclarativeAgentManifestFileSync("testPath");

      chai.assert.isTrue(res.isErr());
    });
  });

  describe("writeDeclarativeAgentManifestFile", async () => {
    it("write manifest successfully", async () => {
      const manifest = {
        schema: "schema",
        description: "description",
      };
      const mockWrapper = {
        data: manifest,
        save: sinon.stub().resolves(),
      };
      sandbox.stub(DeclarativeAgentManifestWrapper, "fromJSON").returns(mockWrapper as any);

      const res = await copilotGptManifestUtils.writeDeclarativeAgentManifestFile(
        manifest as any,
        "testPath"
      );

      chai.assert.isTrue(res.isOk());
    });

    it("write manifest error: write file error", async () => {
      const manifest = {
        schema: "schema",
        description: "description",
      };
      const mockWrapper = {
        data: manifest,
        save: sinon.stub().throws("some error"),
      };
      sandbox.stub(DeclarativeAgentManifestWrapper, "fromJSON").returns(mockWrapper as any);

      const res = await copilotGptManifestUtils.writeDeclarativeAgentManifestFile(
        manifest as any,
        "testPath"
      );

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.isTrue(res.error instanceof WriteFileError);
      }
    });
  });
});
