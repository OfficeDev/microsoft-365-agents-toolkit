// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  ConditionFunc,
  FuncValidation,
  Inputs,
  Platform,
  TextInputQuestion,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import { environmentNameManager } from "../../src/core/environmentName";
import { QuestionNames } from "../../src/question/constants";
import {
  addAuthActionQuestion,
  apiFromPluginManifestQuestion,
  apiSpecFromPluginManifestQuestion,
  kiotaRegenerateQuestion,
  selectTargetEnvQuestion,
} from "../../src/question/other";
import * as sinon from "sinon";
import fs from "fs-extra";

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
    assert.equal(inputs[QuestionNames.ApiPluginType], "new-api");
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
});
