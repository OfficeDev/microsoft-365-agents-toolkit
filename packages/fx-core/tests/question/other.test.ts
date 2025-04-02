// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  ConditionFunc,
  FuncValidation,
  Inputs,
  Platform,
  TextInputQuestion,
  ok,
  err,
  SystemError,
  SingleSelectQuestion,
  SingleFileQuestion,
  AppPackageFolderName,
  TeamsAppManifest,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import fs from "fs-extra";
import "mocha";
import * as sinon from "sinon";
import { environmentNameManager } from "../../src/core/environmentName";
import { QuestionNames } from "../../src/question/constants";
import {
  addAuthActionQuestion,
  apiFromPluginManifestQuestion,
  apiSpecFromPluginManifestQuestion,
  authNameQuestion,
  kiotaRegenerateQuestion,
  oauthAuthorizationUrlQuestion,
  oauthRefreshUrlQuestion,
  oauthScopeQuestion,
  oauthTokenUrlQuestion,
  selectTargetEnvQuestion,
  setSensitivityLabelNode,
  selectDeclarativeAgentManifestQuestion,
} from "../../src/question/other";
import { graphAPIClient } from "../../src/client/graphAPIClient";
import { setTools, TOOLS } from "../../src/common/globalVars";
import path from "path";
import { manifestUtils, TeamsAppAdmin } from "../../src";

describe("env question", () => {
  it("should not show testtool env", async () => {
    const dynamicOptions = selectTargetEnvQuestion(
      QuestionNames.TargetEnvName,
      false
    ).dynamicOptions;
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    if (dynamicOptions) {
      const envs = (await dynamicOptions(inputs)) as string[];
      assert.notInclude(envs, environmentNameManager.getTestToolEnvName());
    }
  });

  it("should not show testtool env for non-remote", async () => {
    const dynamicOptions = selectTargetEnvQuestion(
      QuestionNames.TargetEnvName,
      true
    ).dynamicOptions;
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    if (dynamicOptions) {
      const envs = (await dynamicOptions(inputs)) as string[];
      assert.notInclude(envs, environmentNameManager.getTestToolEnvName());
    }
  });
});

describe("kiotaRegenerate question", () => {
  it("should ask for manifest", async () => {
    const question = kiotaRegenerateQuestion();
    assert.equal(question.data.name, QuestionNames.TeamsAppManifestFilePath);
  });
});

describe("addAuthActionQuestion", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("apiSpecFromPluginManifestQuestion", async () => {
    const inputs = {
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: "test",
    };
    sandbox.stub(fs, "readJson").resolves({
      schema_version: "1.0",
      name_for_human: "test",
      description_for_human: "test",
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec1.yaml",
          },
          run_for_functions: ["function1"],
        },
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec2.yaml",
          },
          run_for_functions: ["function2"],
        },
        {
          type: "LocalPlugin",
          spec: {
            local_endpoint: "spec3.yaml",
          },
        },
      ],
    });
    const apiSpecOptions = apiSpecFromPluginManifestQuestion().dynamicOptions;
    if (apiSpecOptions) {
      const options = await apiSpecOptions(inputs);
      assert.equal(options.length, 2);
    }
  });

  it("apiSpecFromPluginManifestQuestion condition: should skip", async () => {
    const inputs = {
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: "test",
    };
    sandbox.stub(fs, "readJson").resolves({
      schema_version: "1.0",
      name_for_human: "test",
      description_for_human: "test",
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec1.yaml",
          },
          run_for_functions: ["function1"],
        },
      ],
    });
    const condition = addAuthActionQuestion().children![0].condition;
    if (condition) {
      const res = await (condition as ConditionFunc)(inputs);
      assert.isFalse(res);
    }
  });

  it("apiSpecFromPluginManifestQuestion condition: should skip when no plugin manifest file path", async () => {
    const inputs = {
      platform: Platform.VSCode,
    };
    sandbox.stub(fs, "readJson").resolves({
      schema_version: "1.0",
      name_for_human: "test",
      description_for_human: "test",
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec1.yaml",
          },
          run_for_functions: ["function1"],
        },
      ],
    });
    const condition = addAuthActionQuestion().children![0].condition;
    if (condition) {
      const res = await (condition as ConditionFunc)(inputs);
      assert.isFalse(res);
    }
  });

  it("apiSpecFromPluginManifestQuestion condition: should ask question", async () => {
    const inputs = {
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: "test",
    };
    sandbox.stub(fs, "readJson").resolves({
      schema_version: "1.0",
      name_for_human: "test",
      description_for_human: "test",
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec1.yaml",
          },
          run_for_functions: ["function1"],
        },
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec2.yaml",
          },
          run_for_functions: ["function2"],
        },
        {
          type: "LocalPlugin",
          spec: {
            local_endpoint: "spec3.yaml",
          },
        },
      ],
    });
    const condition = addAuthActionQuestion().children![0].condition;
    if (condition) {
      const res = await (condition as ConditionFunc)(inputs);
      assert.isTrue(res);
    }
  });

  it("apiFromPluginManifestQuestion", async () => {
    const inputs = {
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: "test",
      [QuestionNames.ApiSpecLocation]: "spec.yaml",
    };
    sandbox.stub(fs, "readJson").resolves({
      schema_version: "1.0",
      name_for_human: "test",
      description_for_human: "test",
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec.yaml",
          },
          run_for_functions: ["function1"],
        },
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec.yaml",
          },
          run_for_functions: ["function2"],
        },
        {
          type: "LocalPlugin",
          spec: {
            local_endpoint: "spec.yaml",
          },
        },
      ],
    });
    const apiOptions = apiFromPluginManifestQuestion().dynamicOptions;
    if (apiOptions) {
      const options = await apiOptions(inputs);
      assert.equal(options.length, 2);
    }
  });

  it("apiFromPluginManifestQuestion condition: should ask question", async () => {
    const inputs = {
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: "test",
      [QuestionNames.ApiSpecLocation]: "spec.yaml",
    };
    sandbox.stub(fs, "readJson").resolves({
      schema_version: "1.0",
      name_for_human: "test",
      description_for_human: "test",
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec.yaml",
          },
          run_for_functions: ["function1"],
        },
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec.yaml",
          },
          run_for_functions: ["function2"],
        },
        {
          type: "LocalPlugin",
          spec: {
            local_endpoint: "spec.yaml",
          },
        },
      ],
    });
    const condition = addAuthActionQuestion().children![1].condition;
    if (condition) {
      const res = await (condition as ConditionFunc)(inputs);
      assert.isTrue(res);
    }
  });

  it("apiFromPluginManifestQuestion condition: should skip", async () => {
    const inputs = {
      platform: Platform.VSCode,
      [QuestionNames.PluginManifestFilePath]: "test",
      [QuestionNames.ApiSpecLocation]: "spec.yaml",
    };
    sandbox.stub(fs, "readJson").resolves({
      schema_version: "1.0",
      name_for_human: "test",
      description_for_human: "test",
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec.yaml",
          },
          run_for_functions: ["function1"],
        },
      ],
    });
    const condition = addAuthActionQuestion().children![1].condition;
    if (condition) {
      const res = await (condition as ConditionFunc)(inputs);
      assert.isFalse(res);
    }
  });

  it("apiFromPluginManifestQuestion condition: should skip when no plugin manifest file path", async () => {
    const inputs = {
      platform: Platform.VSCode,
    };
    sandbox.stub(fs, "readJson").resolves({
      schema_version: "1.0",
      name_for_human: "test",
      description_for_human: "test",
      runtimes: [
        {
          type: "OpenApi",
          auth: {
            type: "None",
          },
          spec: {
            url: "spec.yaml",
          },
          run_for_functions: ["function1"],
        },
      ],
    });
    const condition = addAuthActionQuestion().children![1].condition;
    if (condition) {
      const res = await (condition as ConditionFunc)(inputs);
      assert.isFalse(res);
    }
  });

  it("authname: validate auth name", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    const validation = (
      (addAuthActionQuestion().children![2].data as TextInputQuestion)
        .additionalValidationOnAccept as FuncValidation<string>
    ).validFunc;
    const res = await validation("input", inputs);
    assert.equal(inputs[QuestionNames.ActionType], "new-api");
  });

  it("authname: should fail if no inputs when validate auth name", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    const validation = (
      (addAuthActionQuestion().children![2].data as TextInputQuestion)
        .additionalValidationOnAccept as FuncValidation<string>
    ).validFunc;
    try {
      const res = await validation("input", undefined);
    } catch (error) {
      assert.equal(error.message, "inputs is undefined");
    }
  });

  it("oauthAuthorizationUrlQuestion: should throw error if no input", async () => {
    const validation = (oauthAuthorizationUrlQuestion().validation as FuncValidation<string>)
      .validFunc;
    try {
      const res = await validation("", undefined);
    } catch (error) {
      assert.equal(error.message, "Invalid URL format. Please enter a valid URL.");
    }
  });

  it("oauthAuthorizationUrlQuestion: happy path", async () => {
    const validation = (oauthAuthorizationUrlQuestion().validation as FuncValidation<string>)
      .validFunc;
    const res = await validation("https://mock-auth-url", undefined);
    assert.isUndefined(res);
  });

  it("oauthAuthorizationUrlQuestion: should throw error if not valid url", async () => {
    const validation = (oauthAuthorizationUrlQuestion().validation as FuncValidation<string>)
      .validFunc;
    try {
      const res = await validation("testUrl", undefined);
    } catch (error) {
      assert.equal(error.message, "Invalid URL format. Please enter a valid URL.");
    }
  });

  it("oauthTokenUrlQuestion: should throw error if no input", async () => {
    const validation = (oauthTokenUrlQuestion().validation as FuncValidation<string>).validFunc;
    try {
      const res = await validation("", undefined);
    } catch (error) {
      assert.equal(error.message, "Invalid URL format. Please enter a valid URL.");
    }
  });

  it("oauthTokenUrlQuestion: happy path", async () => {
    const validation = (oauthTokenUrlQuestion().validation as FuncValidation<string>).validFunc;
    const res = await validation("https://mock-token-url", undefined);
    assert.isUndefined(res);
  });

  it("oauthTokenUrlQuestion: should throw error if not valid url", async () => {
    const validation = (oauthTokenUrlQuestion().validation as FuncValidation<string>).validFunc;
    try {
      const res = await validation("testUrl", undefined);
    } catch (error) {
      assert.equal(error.message, "Invalid URL format. Please enter a valid URL.");
    }
  });

  it("oauthRefreshUrlQuestion: should not throw error if no input", async () => {
    const validation = (oauthRefreshUrlQuestion().validation as FuncValidation<string>).validFunc;
    const res = await validation("", undefined);
    assert.isUndefined(res);
  });

  it("oauthRefreshUrlQuestion: happy path", async () => {
    const validation = (oauthRefreshUrlQuestion().validation as FuncValidation<string>).validFunc;
    const res = await validation("https://mock-refresh-url", undefined);
    assert.isUndefined(res);
  });

  it("oauthRefreshUrlQuestion: should throw error if not valid url", async () => {
    const validation = (oauthRefreshUrlQuestion().validation as FuncValidation<string>).validFunc;
    try {
      const res = await validation("testUrl", undefined);
    } catch (error) {
      assert.equal(error.message, "Invalid URL format. Please enter a valid URL.");
    }
  });

  it("oauthScopeQuestion: should throw error if invalid input", async () => {
    const validation = (oauthScopeQuestion().validation as FuncValidation<string>).validFunc;
    try {
      const res = await validation("scope", undefined);
    } catch (error) {
      assert.equal(
        error.message,
        "Invalid scope format. Please enter a valid scope. Samle: scope1: description for scope1; scope2: description for scope2"
      );
    }
  });

  it("oauthScopeQuestion: happy path", async () => {
    const validation = (oauthScopeQuestion().validation as FuncValidation<string>).validFunc;
    const res = await validation(
      "api://tenant_id: description; api://clientId: description",
      undefined
    );
    assert.isUndefined(res);
  });

  it("authNameQuestion: should throw error if no input", async () => {
    const validation = (authNameQuestion().validation as FuncValidation<string>).validFunc;
    try {
      const res = await validation("", undefined);
    } catch (error) {
      assert.equal(error.message, "Auth name cannot be empty.");
    }
  });

  it("authNameQuestion: happy path", async () => {
    const validation = (authNameQuestion().validation as FuncValidation<string>).validFunc;
    const res = await validation("test", undefined);
    assert.isUndefined(res);
  });
});

describe("setSensitivityLabelNode", () => {
  const sandbox = sinon.createSandbox();
  setTools({
    tokenProvider: {
      m365TokenProvider: {
        getAccessToken: async () => {
          return ok("mockToken");
        },
      },
    },
  } as any);
  afterEach(() => {
    sandbox.restore();
  });

  it("should have correct structure", () => {
    const node = setSensitivityLabelNode();
    assert.isTrue(node.data.type === "group");
    assert.isTrue(node.children?.length === 2);
  });

  it("validate sensitivity label question", async () => {
    const node = setSensitivityLabelNode();
    const sensitivityLabelQuestion = node.children?.[1].data as SingleSelectQuestion;
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    const mockLabels = [
      { id: "1", displayName: "Label1" },
      { id: "2", displayName: "Label2" },
    ];
    sandbox.stub(graphAPIClient, "listSensitivityLabels").resolves(ok(mockLabels));
    // mock token provider
    sandbox.stub(TOOLS.tokenProvider.m365TokenProvider, "getAccessToken").resolves(ok("mockToken"));
    const options = await sensitivityLabelQuestion?.dynamicOptions?.(inputs);
    assert.equal(options?.length, 2);
    assert.equal((options?.[0] as any).id, "1");
    assert.equal((options?.[0] as any).label, "Label1");
  });

  it("should handle graphAPI exception", async () => {
    const node = setSensitivityLabelNode();
    const sensitivityLabelQuestion = node.children?.[1].data as SingleSelectQuestion;
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    sandbox.stub(graphAPIClient, "listSensitivityLabels").throws(new Error("Graph API error"));
    // mock token provider
    sandbox.stub(TOOLS.tokenProvider.m365TokenProvider, "getAccessToken").resolves(ok("mockToken"));
    const options = await sensitivityLabelQuestion?.dynamicOptions?.(inputs);
    assert.equal(options?.length, 0);
  });

  it("should handle token error in sensitivity label question", async () => {
    const node = setSensitivityLabelNode();
    const sensitivityLabelQuestion = node.children?.[1].data as SingleSelectQuestion;
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    sandbox
      .stub(TOOLS.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(err(new SystemError("TestError", "Test error message", "TestSource")));
    const mockLabels = [
      { id: "1", displayName: "Label1" },
      { id: "2", displayName: "Label2" },
    ];
    sandbox.stub(graphAPIClient, "listSensitivityLabels").resolves(ok(mockLabels));
    const options = await sensitivityLabelQuestion?.dynamicOptions?.(inputs);
    assert.equal(options?.length, 0);
  });

  it("should handle error in sensitivity label question", async () => {
    const node = setSensitivityLabelNode();
    const sensitivityLabelQuestion = node.children?.[1].data as SingleSelectQuestion;
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    sandbox.stub(graphAPIClient, "listSensitivityLabels").resolves(
      err(
        new SystemError({
          name: "TestError",
          message: "Test error message",
          source: "TestSource",
        })
      )
    );
    const options = await sensitivityLabelQuestion?.dynamicOptions?.(inputs);
    assert.equal(options?.length, 0);
  });

  it("should return the correct default path for selectDeclarativeAgentManifestQuestion - CLI_HELP", async () => {
    const inputs: Inputs = {
      platform: Platform.CLI_HELP,
      projectPath: "./testProject",
    };
    const question = selectDeclarativeAgentManifestQuestion() as SingleFileQuestion;
    const defaultPath = await ((question?.default as any)(inputs) as Promise<string | undefined>);
    assert.equal(defaultPath, "./appPackage/declarativeAgent.json");
  });

  it("should return the correct default path for selectDeclarativeAgentManifestQuestion", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: "./testProject",
    };
    sandbox.stub(fs, "pathExistsSync").returns(true);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(
      ok({
        copilotAgents: {
          declarativeAgents: [
            {
              file: "agent.json",
            },
          ],
        },
      } as TeamsAppManifest)
    );
    const question = selectDeclarativeAgentManifestQuestion() as SingleFileQuestion;
    const defaultPath = await ((question?.default as any)(inputs) as Promise<string | undefined>);
    assert.equal(defaultPath, path.join(inputs.projectPath!, AppPackageFolderName, "agent.json"));
  });

  it("should return undefined if projectPath is not defined for selectDeclarativeAgentManifestQuestion", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    const question = selectDeclarativeAgentManifestQuestion() as SingleFileQuestion;
    const defaultPath = await ((question?.default as any)(inputs) as Promise<string | undefined>);
    assert.isUndefined(defaultPath);
  });

  it("should return undefined if manifest path does not exist for selectDeclarativeAgentManifestQuestion", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: "./nonExistentProject",
    };
    const question = selectDeclarativeAgentManifestQuestion() as SingleFileQuestion;
    const defaultPath = await ((question?.default as any)(inputs) as Promise<string | undefined>);
    assert.isUndefined(defaultPath);
  });

  it("should return undefined if manifest does not contain DA for selectDeclarativeAgentManifestQuestion", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: "./testProject",
    };
    sandbox.stub(fs, "pathExistsSync").returns(true);
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(ok({} as any));
    const question = selectDeclarativeAgentManifestQuestion() as SingleFileQuestion;
    const defaultPath = await ((question?.default as any)(inputs) as Promise<string | undefined>);
    assert.isUndefined(defaultPath);
  });

  it("should return error if failed to read manifest for selectDeclarativeAgentManifestQuestion", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: "./testProject",
    };
    sandbox.stub(fs, "pathExistsSync").returns(true);
    sandbox
      .stub(manifestUtils, "_readAppManifest")
      .resolves(err(new SystemError("TestError", "Test error message", "TestSource")));
    const question = selectDeclarativeAgentManifestQuestion() as SingleFileQuestion;
    const defaultPath = await ((question?.default as any)(inputs) as Promise<string | undefined>);
    assert.isUndefined(defaultPath);
  });

  it("should return error if declarativeAgentManifest path does not exist", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: "./testProject",
    };
    sandbox.stub(fs, "pathExistsSync").callsFake((path: string) => {
      if (path.includes("manifest")) {
        return true;
      }
      return false;
    });
    sandbox.stub(manifestUtils, "_readAppManifest").resolves(
      ok({
        copilotAgents: {
          declarativeAgents: [
            {
              file: "agent.json",
            },
          ],
        },
      } as TeamsAppManifest)
    );
    const question = selectDeclarativeAgentManifestQuestion() as SingleFileQuestion;
    const defaultPath = await ((question?.default as any)(inputs) as Promise<string | undefined>);
    assert.isUndefined(defaultPath);
  });
});
