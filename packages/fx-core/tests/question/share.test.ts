// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { ConditionFunc, Inputs, OptionItem, SingleSelectQuestion } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import * as sinon from "sinon";
import { setTools } from "../../src/common/globalVars";
import { QuestionNames } from "../../src/question/constants";
import {
  shareNode,
  ShareOperationOption,
  ShareOperationOptions,
  ShareScopeOption,
} from "../../src/question/share";
import { MockTools } from "../core/utils";

describe("shareNode", () => {
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
    assert.lengthOf(result.children!, 2);

    // Verify the data is share operation question
    const shareOperationQuestion = result.data as SingleSelectQuestion;
    assert.property(shareOperationQuestion, "name");
    assert.equal(shareOperationQuestion.name, QuestionNames.ShareOperation);
    assert.property(shareOperationQuestion, "type");
    assert.equal(shareOperationQuestion.type, "singleSelect");

    // Verify share scope options is the first child
    const shareScopeNode = result.children![0];
    assert.isObject(shareScopeNode);
    // Check the condition - it should be an object with the equals property
    assert.property(shareScopeNode, "condition");
    const condition = shareScopeNode.condition as any;
    assert.equal(condition.equals, ShareOperationOptions.shareWithUsers().id);
  });

  it("shareOperation question should have correct options", () => {
    const result = shareNode();
    const shareOperationQuestion = result.data as SingleSelectQuestion;

    assert.isArray(shareOperationQuestion.staticOptions);
    assert.lengthOf(shareOperationQuestion.staticOptions, 2);

    const [shareWithUsers, removeShareAccess] =
      shareOperationQuestion.staticOptions as OptionItem[];
    assert.equal(shareWithUsers.id, ShareOperationOption.ShareWithUsers);
    assert.equal(removeShareAccess.id, ShareOperationOption.RemoveShareAccessFromUsers);
  });

  it("shareScope should have correct options", () => {
    const result = shareNode();
    const shareScopeNode = result.children![0];
    const shareScopeQuestion = shareScopeNode.data as SingleSelectQuestion;

    assert.property(shareScopeQuestion, "name");
    assert.equal(shareScopeQuestion.name, QuestionNames.ShareScope);
    assert.property(shareScopeQuestion, "type");
    assert.equal(shareScopeQuestion.type, "singleSelect");
    assert.isArray(shareScopeQuestion.staticOptions);
    assert.lengthOf(shareScopeQuestion.staticOptions, 3);

    // Check scope options
    const [shareTenant, shareUsers, shareOwners] = shareScopeQuestion.staticOptions as OptionItem[];
    assert.equal(shareTenant.id, ShareScopeOption.ShareAppWithTenantUsers);
    assert.equal(shareUsers.id, ShareScopeOption.ShareAppWithSpecificUsers);
    assert.equal(shareOwners.id, ShareScopeOption.ShareAppWithOwners);
  });

  it("email input should be shown when ShareScope is specific users or owners", () => {
    const result = shareNode();
    const shareScopeNode = result.children![0];
    assert.isArray(shareScopeNode.children);
    const emailInputNode = shareScopeNode.children![0];

    // Test condition function
    assert.property(emailInputNode, "condition");
    assert.isFunction(emailInputNode.condition);

    // Create test inputs
    const inputsSpecificUsers = {
      [QuestionNames.ShareScope]: ShareScopeOption.ShareAppWithSpecificUsers,
    };
    const inputsOwners = { [QuestionNames.ShareScope]: ShareScopeOption.ShareAppWithOwners };
    const inputsTenant = { [QuestionNames.ShareScope]: ShareScopeOption.ShareAppWithTenantUsers };

    // Call the condition function with different inputs
    const conditionFunc = emailInputNode.condition as ConditionFunc;
    assert.isTrue(conditionFunc(inputsSpecificUsers as unknown as Inputs));
    assert.isTrue(conditionFunc(inputsOwners as unknown as Inputs));
    assert.isFalse(conditionFunc(inputsTenant as unknown as Inputs));
  });

  it("email input should be shown when ShareOperation is RemoveShareAccessFromUsers", () => {
    const result = shareNode();
    assert.isArray(result.children);
    const emailInputNode = result.children![1];

    // Test condition function
    assert.property(emailInputNode, "condition");
    assert.isFunction(emailInputNode.condition);

    // Create test inputs
    const inputsRemoveAccess = {
      [QuestionNames.ShareOperation]: ShareOperationOption.RemoveShareAccessFromUsers,
    };
    const inputsShare = { [QuestionNames.ShareOperation]: ShareOperationOption.ShareWithUsers };

    // Call the condition function with different inputs
    const conditionFunc = emailInputNode.condition as ConditionFunc;
    assert.isTrue(conditionFunc(inputsRemoveAccess as unknown as Inputs));
    assert.isFalse(conditionFunc(inputsShare as unknown as Inputs));
  });
});
