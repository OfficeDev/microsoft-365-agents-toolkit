// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { err, ok, OptionItem, SingleSelectQuestion, SystemError } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import * as sinon from "sinon";
import { teamsDevPortalClient } from "../../src";
import { setTools, TOOLS } from "../../src/common/globalVars";
import * as shareUtils from "../../src/component/driver/share/utils";
import { CollaborationUtil } from "../../src/core/collaborator";
import { QuestionNames } from "../../src/question/constants";
import {
  removeSharedAccessNode,
  selectUsersToRemoveSharedAccess,
  shareNode,
} from "../../src/question/share";
import { MockTools } from "../core/utils";

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
