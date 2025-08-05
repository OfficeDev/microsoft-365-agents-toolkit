// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Inputs,
  MultiSelectQuestion,
  Platform,
  StringArrayValidation,
  TextInputQuestion,
} from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import sinon from "sinon";
import { featureFlagManager } from "../../src/common/featureFlags";
import { CollaborationConstants } from "../../src/core/collaborator";
import {
  grantPermissionQuestionNode,
  listCollaboratorQuestionNode,
} from "../../src/question/collaborator";

describe("Collaboration Question Node Tests", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  describe("grantPermissionQuestionNode", () => {
    it("should return question node with correct structure", () => {
      const node = grantPermissionQuestionNode();

      // Check root node structure
      assert.equal(node.data.type, "group");
      assert.isDefined(node.children);
      assert.lengthOf(node.children!, 1);

      // Check root child node
      const rootChild = node.children![0];
      assert.isDefined(rootChild.condition);
      assert.isDefined(rootChild.data);
      assert.equal(rootChild.cliOptionDisabled, "self");
      assert.equal(rootChild.inputsDisabled, "self");
      assert.isDefined(rootChild.children);
      assert.lengthOf(rootChild.children!, 3); // Teams app, AAD app, and email input

      // Check that the condition function works correctly
      const conditionFn = rootChild.condition as (inputs: Inputs) => boolean;
      assert.isTrue(conditionFn({ platform: Platform.VSCode }));
      assert.isTrue(conditionFn({ platform: Platform.CLI }));

      // Check Teams app manifest node
      const teamsAppNode = rootChild.children![0];
      assert.isDefined(teamsAppNode.condition);
      assert.deepEqual(teamsAppNode.condition, {
        contains: CollaborationConstants.TeamsAppQuestionId,
      });

      // Check AAD app manifest node
      const aadAppNode = rootChild.children![1];
      assert.isDefined(aadAppNode.condition);
      assert.deepEqual(aadAppNode.condition, {
        contains: CollaborationConstants.AadAppQuestionId,
      });

      // Check email input node
      const emailNode = rootChild.children![2];
      assert.isDefined(emailNode.data);
      const emailQuestion = emailNode.data as TextInputQuestion;
      assert.include(emailQuestion.title!, "Add owner to");
    });

    it("should include agent option when ShareEnabled flag is true", () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      const node = grantPermissionQuestionNode();
      const appTypeQuestion = node.children![0].data as MultiSelectQuestion;

      assert.isDefined(appTypeQuestion.staticOptions);
      const options = appTypeQuestion.staticOptions as { id: string }[];
      assert.lengthOf(options, 3);
      assert.isTrue(options.some((o) => o.id === CollaborationConstants.AgentOptionId));
    });

    it("should not include agent option when ShareEnabled flag is false", () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      const node = grantPermissionQuestionNode();
      const appTypeQuestion = node.children![0].data as MultiSelectQuestion;

      assert.isDefined(appTypeQuestion.staticOptions);
      const options = appTypeQuestion.staticOptions as { id: string }[];
      assert.lengthOf(options, 2);
      assert.isFalse(options.some((o) => o.id === CollaborationConstants.AgentOptionId));
    });

    it("should require at least one app type selection", () => {
      const node = grantPermissionQuestionNode();
      if (!node.children) return;
      const appTypeQuestion = node.children[0].data as MultiSelectQuestion;
      const validation = appTypeQuestion.validation as StringArrayValidation;

      assert.isDefined(validation);
      assert.equal(validation.minItems, 1);
    });
  });

  describe("listCollaboratorQuestionNode", () => {
    it("should return question node with correct structure", () => {
      const node = listCollaboratorQuestionNode();

      // Check root node structure
      assert.equal(node.data.type, "group");
      assert.isDefined(node.children);
      assert.lengthOf(node.children!, 1);

      // Check root child node
      const rootChild = node.children![0];
      assert.isDefined(rootChild.condition);
      assert.isDefined(rootChild.data);
      assert.equal(rootChild.cliOptionDisabled, "self");
      assert.equal(rootChild.inputsDisabled, "self");
      assert.isDefined(rootChild.children);
      assert.lengthOf(rootChild.children!, 2); // Teams app and AAD app, no email input

      // Check that the condition function works correctly
      const conditionFn = rootChild.condition as (inputs: Inputs) => boolean;
      assert.isTrue(conditionFn({ platform: Platform.VSCode }));
      assert.isTrue(conditionFn({ platform: Platform.CLI }));

      // Check Teams app manifest node
      const teamsAppNode = rootChild.children![0];
      assert.isDefined(teamsAppNode.condition);
      assert.deepEqual(teamsAppNode.condition, {
        contains: CollaborationConstants.TeamsAppQuestionId,
      });

      // Check AAD app manifest node
      const aadAppNode = rootChild.children![1];
      assert.isDefined(aadAppNode.condition);
      assert.deepEqual(aadAppNode.condition, {
        contains: CollaborationConstants.AadAppQuestionId,
      });
    });

    it("should include agent option when ShareEnabled flag is true", () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(true);
      const node = listCollaboratorQuestionNode();
      const appTypeQuestion = node.children![0].data as MultiSelectQuestion;

      assert.isDefined(appTypeQuestion.staticOptions);
      const options = appTypeQuestion.staticOptions as { id: string }[];
      assert.lengthOf(options, 3);
      assert.isTrue(options.some((o) => o.id === CollaborationConstants.AgentOptionId));
    });

    it("should not include agent option when ShareEnabled flag is false", () => {
      sandbox.stub(featureFlagManager, "getBooleanValue").returns(false);
      const node = listCollaboratorQuestionNode();
      const appTypeQuestion = node.children![0].data as MultiSelectQuestion;

      assert.isDefined(appTypeQuestion.staticOptions);
      const options = appTypeQuestion.staticOptions as { id: string }[];
      assert.lengthOf(options, 2);
      assert.isFalse(options.some((o) => o.id === CollaborationConstants.AgentOptionId));
    });

    it("should require at least one app type selection", () => {
      const node = listCollaboratorQuestionNode();
      if (!node.children) return;
      const appTypeQuestion = node.children[0].data as MultiSelectQuestion;
      const validation = appTypeQuestion.validation as StringArrayValidation;

      assert.isDefined(validation);
      assert.equal(validation.minItems, 1);
    });
  });
});
