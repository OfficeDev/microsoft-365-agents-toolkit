// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import * as sinon from "sinon";
import chai from "chai";
import fs from "fs-extra";
import {
  DeclarativeCopilotManifestSchema,
  ManifestUtil,
  Platform,
  SystemError,
  ok,
  err,
  Colors,
  UserError,
} from "@microsoft/teamsfx-api";
import { copilotGptManifestUtils } from "../../../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import {
  FileNotFoundError,
  MissingEnvironmentVariablesError,
  WriteFileError,
} from "../../../../src/error";
import mockedEnv, { RestoreFn } from "mocked-env";
import { pluginManifestUtils } from "../../../../src/component/driver/teamsApp/utils/PluginManifestUtils";
import { AppStudioError } from "../../../../src/component/driver/teamsApp/errors";
import { DeclarativeCopilotManifestValidationResult } from "../../../../src/component/driver/teamsApp/interfaces/ValidationResult";
import { MockedLogProvider, MockedTelemetryReporter } from "../../../plugins/solution/util";
import { WrapDriverContext } from "../../../../src/component/driver/util/wrapUtil";
import { createContext, setTools } from "../../../../src/common/globalVars";
import { generateDriverContext } from "../../../../src/common/utils";
import { MockTools } from "../../../core/utils";
import { manifestUtils } from "../../../../src/component/driver/teamsApp/utils/ManifestUtils";
import path from "path";
import { SpecParser } from "@microsoft/m365-spec-parser";

describe("copilotGptManifestUtils", () => {
  const sandbox = sinon.createSandbox();
  let mockedEnvRestore: RestoreFn;

  afterEach(async () => {
    sandbox.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });

  const gptManifest: DeclarativeCopilotManifestSchema = {
    name: "name${{APP_NAME_SUFFIX}}",
    description: "description",
  };

  describe("add plugin", async () => {
    it("add plugin without appending conversation starters success", async () => {
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(
        JSON.stringify({
          name: "name${{APP_NAME_SUFFIX}}",
          description: "description",
          conversation_starters: [
            {
              text: "List all repairs",
            },
          ],
        }) as any
      );
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(fs, "readJson").resolves({} as any);

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
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(JSON.stringify(gptManifest) as any);
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(fs, "readJson").resolves({
        capabilities: {
          conversation_starters: [
            {
              text: "List all repairs",
            },
          ],
        },
      } as any);

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
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(JSON.stringify(gptManifest) as any);
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(fs, "readJson").resolves({
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
      } as any);

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
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(JSON.stringify(gptManifest) as any);
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(fs, "readJson").resolves({
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
      } as any);

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
      sandbox.stub(fs, "readFile").resolves(
        JSON.stringify({
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
        }) as any
      );
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(fs, "readJson").resolves({
        capabilities: {
          conversation_starters: [
            {
              text: "List all repairs2",
            },
          ],
        },
      } as any);

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
      sandbox.stub(fs, "readFile").resolves(
        JSON.stringify({
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
        }) as any
      );
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(fs, "readJson").resolves({
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
      } as any);

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
      sandbox.stub(fs, "readFile").resolves(
        JSON.stringify({
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
        }) as any
      );
      sandbox.stub(fs, "writeFile").resolves();
      sandbox.stub(fs, "readJson").resolves({
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
      } as any);

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
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(JSON.stringify(gptManifest) as any);
      sandbox.stub(fs, "writeFile").throws("some error");
      sandbox.stub(fs, "readJson").resolves({} as any);
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
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(JSON.stringify(gptManifest) as any);

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
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(JSON.stringify(gptManifest) as any);

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
      const manifest: DeclarativeCopilotManifestSchema = {
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
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(JSON.stringify(manifest) as any);
      sandbox.stub(ManifestUtil, "validateManifest").resolves([]);
      sandbox.stub(pluginManifestUtils, "validateAgainstSchema").resolves(
        ok({
          id: "1",
          filePath: "testFile",
          validationResult: ["error1"],
        })
      );

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
      const manifest: DeclarativeCopilotManifestSchema = {
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
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(JSON.stringify(manifest) as any);
      sandbox.stub(ManifestUtil, "validateManifest").resolves([]);
      sandbox
        .stub(pluginManifestUtils, "validateAgainstSchema")
        .resolves(err(new SystemError("error", "error", "error", "error")));

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
      const manifest: DeclarativeCopilotManifestSchema = {
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
      sandbox.stub(fs, "pathExists").resolves(true);
      sandbox.stub(fs, "readFile").resolves(JSON.stringify(gptManifest) as any);
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

      const res = copilotGptManifestUtils.logValidationErrors(validationRes, Platform.VSCode, "");
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
        Platform.VSCode,
        "pluginPath"
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
        Platform.VSCode,
        "pluginPath"
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
          {
            id: "2",
            filePath: "pluginPath",
            validationResult: ["errorAction2"],
          },
        ],
      };

      const res = copilotGptManifestUtils.logValidationErrors(
        validationRes,
        Platform.CLI,
        "pluginPath"
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
        Platform.CLI,
        ""
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
});
