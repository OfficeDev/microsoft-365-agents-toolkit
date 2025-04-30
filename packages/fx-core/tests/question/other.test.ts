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
  OptionItem,
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
  shareNode,
  removeSharedAccessNode,
  selectUsersToRemoveSharedAccess,
} from "../../src/question/other";
import { GraphClient } from "../../src/client/graphClient";
import { setTools, TOOLS } from "../../src/common/globalVars";
import path from "path";
import { manifestUtils, teamsDevPortalClient } from "../../src";
import { CollaborationUtil } from "../../src/core/collaborator";
import * as shareUtils from "../../src/component/driver/share/utils";
import { MockTools } from "../core/utils";

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

  it("list sensitivity label happy path", async () => {
    const node = setSensitivityLabelNode();
    const sensitivityLabelQuestion = node.children?.[1].data as SingleSelectQuestion;
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    const mockLabels = [
      { id: "1", displayName: "Label1" },
      { id: "2", displayName: "Label2" },
    ];
    sandbox.stub(GraphClient.prototype, "listSensitivityLabels").resolves(ok(mockLabels));
    // mock token provider
    sandbox.stub(TOOLS.tokenProvider.m365TokenProvider, "getAccessToken").resolves(ok("mockToken"));
    const options = await sensitivityLabelQuestion?.dynamicOptions?.(inputs);
    assert.equal(options?.length, 2);
    assert.equal((options?.[0] as any).id, "1");
    assert.equal((options?.[0] as any).label, "Label1");
  });

  it("list sensitivity label - empty label", async () => {
    const node = setSensitivityLabelNode();
    const sensitivityLabelQuestion = node.children?.[1].data as SingleSelectQuestion;
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    const mockLabels = [{}, {}];
    sandbox.stub(GraphClient.prototype, "listSensitivityLabels").resolves(ok(mockLabels));
    // mock token provider
    sandbox.stub(TOOLS.tokenProvider.m365TokenProvider, "getAccessToken").resolves(ok("mockToken"));
    const options = await sensitivityLabelQuestion?.dynamicOptions?.(inputs);
    assert.equal(options?.length, 2);
    assert.equal((options?.[0] as any).id, "");
    assert.equal((options?.[0] as any).label, "");
  });

  it("throw graphAPI exception", async () => {
    const node = setSensitivityLabelNode();
    const sensitivityLabelQuestion = node.children?.[1].data as SingleSelectQuestion;
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    sandbox
      .stub(GraphClient.prototype, "listSensitivityLabels")
      .throws(new Error("Graph API error"));
    // mock token provider
    sandbox.stub(TOOLS.tokenProvider.m365TokenProvider, "getAccessToken").resolves(ok("mockToken"));
    let exception = undefined;
    try {
      await sensitivityLabelQuestion?.dynamicOptions?.(inputs);
    } catch (e) {
      exception = e;
    }
    assert.isTrue((exception as any).message.includes("Graph API error"));
  });

  it("throw token error", async () => {
    const node = setSensitivityLabelNode();
    const sensitivityLabelQuestion = node.children?.[1].data as SingleSelectQuestion;
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    sandbox
      .stub(TOOLS.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(err(new SystemError("TestSource", "TestError", "Test error message")));
    const mockLabels = [
      { id: "1", displayName: "Label1" },
      { id: "2", displayName: "Label2" },
    ];
    sandbox.stub(GraphClient.prototype, "listSensitivityLabels").resolves(ok(mockLabels));

    let exception = undefined;
    try {
      await sensitivityLabelQuestion?.dynamicOptions?.(inputs);
    } catch (e) {
      exception = e;
    }
    assert.isTrue((exception as any).message.includes("Test error message"));
  });

  it("throw list sensitivity label error", async () => {
    const node = setSensitivityLabelNode();
    const sensitivityLabelQuestion = node.children?.[1].data as SingleSelectQuestion;
    const inputs: Inputs = {
      platform: Platform.VSCode,
    };
    sandbox.stub(GraphClient.prototype, "listSensitivityLabels").resolves(
      err(
        new SystemError({
          name: "TestError",
          message: "Test error message",
          source: "TestSource",
        })
      )
    );
    let exception = undefined;
    try {
      await sensitivityLabelQuestion?.dynamicOptions?.(inputs);
    } catch (e) {
      exception = e;
    }
    assert.isTrue((exception as any).message.includes("Test error message"));
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

describe("shareNode and removeSharedAccessNode", () => {
  const sandbox = sinon.createSandbox();
  setTools(new MockTools());
  afterEach(() => {
    sandbox.restore();
  });

  it("shareNode should return IQTreeNode with correct children", () => {
    const result = shareNode();

    // Verify the main node structure
    assert.isObject(result);
    assert.property(result, "data");
    assert.property(result, "children");
    assert.isArray(result.children);
    assert.lengthOf(result.children!, 1);

    // Verify share option question node
    const shareOptionNode = result.children![0];
    assert.property(shareOptionNode.data, "name");
    assert.equal(shareOptionNode.data.name, "option");
    assert.property(shareOptionNode.data, "type");
    assert.equal(shareOptionNode.data.type, "singleSelect");

    // Verify children of share option question
    assert.property(shareOptionNode, "children");
    assert.isArray(shareOptionNode.children);
    assert.lengthOf(shareOptionNode.children!, 1);
  });

  it("shareOption question should have correct options", () => {
    const result = shareNode();
    const shareOptionQuestion = result.children![0].data as SingleSelectQuestion;

    assert.isArray(shareOptionQuestion.staticOptions);
    assert.lengthOf(shareOptionQuestion.staticOptions, 2);

    const [shareApp, shareToUser] = shareOptionQuestion.staticOptions as OptionItem[];
    assert.equal(shareApp.id, "share-app");
    assert.equal(shareToUser.id, "share-with-users");
  });

  it("shareToUser question should be shown when shareOption is shareToUser", () => {
    const result = shareNode();
    const shareToUserNode = result.children![0].children![0] as any;

    // Test condition function
    const inputs = { [QuestionNames.ShareOption]: QuestionNames.ShareOptionShareToUser };
    assert.isTrue(shareToUserNode.condition(inputs));

    // Test when not shareToUser
    const inputs2 = { shareOption: "share-app" };
    assert.isFalse(shareToUserNode.condition(inputs2));
  });

  it("shareToUser validation should work correctly", () => {
    const result = shareNode();
    const shareToUserQuestion = result.children![0].children![0].data as any;

    // Test validation function
    const validation = shareToUserQuestion.validation!.validFunc;

    // Empty input should return error message
    assert.isString(validation(""));
    assert.isString(validation("   "));

    // Valid input should return undefined
    assert.isUndefined(validation("user@example.com"));
  });

  it("removeSharedAccessNode should return IQTreeNode with correct structure", () => {
    const result = removeSharedAccessNode();

    assert.isObject(result);
    assert.property(result, "data");
    assert.property(result, "children");
    assert.isArray(result.children);
    assert.lengthOf(result.children!, 1);

    // Check selectUsersToRemoveSharedAccess question
    const selectUsersQuestion = result.children![0].data as any;
    assert.equal(selectUsersQuestion.name, "users");
    assert.equal(selectUsersQuestion.type, "multiSelect");
    assert.property(selectUsersQuestion, "dynamicOptions");
    assert.isTrue(selectUsersQuestion.skipValidation);
  });

  it("selectUsersToRemoveSharedAccess should throw error when project path is undefined", async () => {
    const result = selectUsersToRemoveSharedAccess();

    try {
      await result.dynamicOptions!({} as any);
      assert.fail("Should have thrown error");
    } catch (error: any) {
      assert.equal(error.message, "Project path is not defined");
    }
  });

  it("selectUsersToRemoveSharedAccess should return correct options", async () => {
    sandbox
      .stub(TOOLS.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(ok("mock-token"));

    // Mock CollaborationUtil.getCurrentUserInfo
    sandbox
      .stub(CollaborationUtil, "getCurrentUserInfo")
      .resolves(ok({ aadId: "current-user-id" } as any));

    sandbox
      .stub(shareUtils, "parseShareAppActionYamlConfig")
      .resolves(ok(["mock-teams-app-id"] as any));

    // Mock teamsDevPortalClient.getApp
    sandbox.stub(teamsDevPortalClient, "getApp").resolves({
      userList: [
        {
          aadId: "current-user-id",
          displayName: "Current User",
          userPrincipalName: "current@example.com",
        },
        {
          aadId: "other-user-id",
          displayName: "Other User",
          userPrincipalName: "other@example.com",
        },
      ],
    } as any);

    const result = selectUsersToRemoveSharedAccess();
    const options = await result.dynamicOptions!({ projectPath: "mockPath" } as any);

    assert.isArray(options);
    assert.lengthOf(options, 1); // Should only include the other user, not current user
    assert.equal((options[0] as any).id, "other@example.com");
    assert.equal((options[0] as any).label, "Other User");
  });

  it("selectUsersToRemoveSharedAccess should handle empty user list", async () => {
    sandbox
      .stub(TOOLS.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(ok("mock-token"));

    // Mock parseShareAppActionYamlConfig
    sandbox.stub(shareUtils, "parseShareAppActionYamlConfig").resolves(ok(["mock-teams-app-id"]));

    // Mock teamsDevPortalClient.getApp with empty user list
    sandbox.stub(teamsDevPortalClient, "getApp").resolves({
      userList: [],
    });

    const result = selectUsersToRemoveSharedAccess();

    try {
      await result.dynamicOptions!({ projectPath: "mockPath" } as any);
      assert.fail("Should have thrown error");
    } catch (error: any) {
      assert.equal(error.message, "No owner found in the app");
    }
  });

  it("selectUsersToRemoveSharedAccess - token error", async () => {
    // Mock TOOLS and required responses
    sandbox
      .stub(TOOLS.tokenProvider.m365TokenProvider, "getAccessToken")
      .resolves(err(new SystemError("TestSource", "TestError", "Test error message")));

    // Mock CollaborationUtil.getCurrentUserInfo
    sandbox
      .stub(CollaborationUtil, "getCurrentUserInfo")
      .resolves(ok({ aadId: "current-user-id" } as any));

    sandbox
      .stub(shareUtils, "parseShareAppActionYamlConfig")
      .resolves(ok(["mock-teams-app-id"] as any));

    // Mock teamsDevPortalClient.getApp
    sandbox.stub(teamsDevPortalClient, "getApp").resolves({
      userList: [
        {
          aadId: "current-user-id",
          displayName: "Current User",
          userPrincipalName: "current@example.com",
        },
        {
          aadId: "other-user-id",
          displayName: "Other User",
          userPrincipalName: "other@example.com",
        },
      ],
    } as any);

    const result = selectUsersToRemoveSharedAccess();
    let exception = undefined;
    try {
      const options = await result.dynamicOptions!({ projectPath: "mockPath" } as any);
    } catch (error: any) {
      exception = error;
    }
    assert.isTrue((exception as any).message.includes("Test error message"));
  });
});
