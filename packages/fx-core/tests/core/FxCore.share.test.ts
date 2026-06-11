// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  err,
  Inputs,
  IProgressHandler,
  ok,
  Platform,
  SystemError,
  UserError,
} from "@microsoft/teamsfx-api";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import fs from "fs-extra";
import { createSandbox, match } from "sinon";
import { InputValidationError, MAX_EMAIL_NUMBER, teamsDevPortalClient } from "../../src";
import { ProjectModel } from "../../src/component/configManager/interface";
import { PackageService } from "../../src/component/m365/packageService";
import { envUtil } from "../../src/component/utils/envUtil";
import { metadataUtil } from "../../src/component/utils/metadataUtil";
import { pathUtils } from "../../src/component/utils/pathUtils";
import { CollaborationUtil } from "../../src/core/collaborator";
import { FxCore, fxCoreDeps } from "../../src/core/FxCore";
import { QuestionNames } from "../../src/question/questionNames";
import { ShareOperationOption, ShareScopeOption } from "../../src/question/share";
import { MockLogProvider, MockTools } from "./utils";

chai.use(chaiAsPromised);

describe("FxCore.shareApplication", () => {
  const sandbox = createSandbox();
  const tools = new MockTools();
  const logger = new MockLogProvider();
  const mockProjectModel: ProjectModel = {
    version: "1.10.0",
  };

  beforeEach(() => {});

  afterEach(() => {
    sandbox.restore();
  });

  describe("Share with tenant users", () => {
    it("share happy path", async () => {
      const shareWithTenantStub = sandbox
        .stub(fxCoreDeps, "shareWithTenant")
        .resolves(ok(undefined));

      sandbox
        .stub(fxCoreDeps, "parseShareAppActionYamlConfig")
        .resolves(ok({ teamsappId: "mockAppId", titleId: "mockTitleId", appId: "mockAppId" }));
      sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
        if (config.name === "env") {
          return ok({ type: "success", result: "dev" });
        } else {
          return ok({ type: "success", result: "" });
        }
      });
      const progressStartStub = sandbox.stub();
      const progressEndStub = sandbox.stub();
      sandbox.stub(tools.ui, "createProgressBar").returns({
        start: progressStartStub,
        end: progressEndStub,
      } as any as IProgressHandler);
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
      sandbox.stub(pathUtils, "getYmlFilePath").returns("m365agents.yml");
      sandbox.stub(fs, "pathExistsSync").returns(true);
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        [QuestionNames.ShareOperation]: ShareOperationOption.ShareWithUsers,
        [QuestionNames.ShareScope]: ShareScopeOption.ShareAppWithTenantUsers,
      };
      const fxCore = new FxCore(tools);
      const res = await fxCore.shareApplication(inputs);
      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(shareWithTenantStub.calledOnce);
    });
  });

  describe("Share with specific users", () => {
    it("share with specific users happy path", async () => {
      const addSharedUsersStub = sandbox.stub(fxCoreDeps, "addSharedUsers").resolves(ok(undefined));

      // Setup common stubs
      sandbox
        .stub(fxCoreDeps, "parseShareAppActionYamlConfig")
        .resolves(ok({ teamsappId: "mockAppId", titleId: "mockTitleId", appId: "mockAppId" }));
      sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
        if (config.name === "env") {
          return ok({ type: "success", result: "dev" });
        } else {
          return ok({ type: "success", result: "" });
        }
      });

      const progressStartStub = sandbox.stub();
      const progressEndStub = sandbox.stub();
      sandbox.stub(tools.ui, "createProgressBar").returns({
        start: progressStartStub,
        end: progressEndStub,
      } as any as IProgressHandler);
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
      sandbox.stub(pathUtils, "getYmlFilePath").returns("m365agents.yml");
      sandbox.stub(fs, "pathExistsSync").returns(true);

      const emails = "user1@example.com,user2@example.com";
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        [QuestionNames.ShareOperation]: ShareOperationOption.ShareWithUsers,
        [QuestionNames.ShareScope]: ShareScopeOption.ShareAppWithSpecificUsers,
        [QuestionNames.UserEmail]: emails,
      };

      const fxCore = new FxCore(tools);
      const res = await fxCore.shareApplication(inputs);

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(addSharedUsersStub.calledOnce);
      chai.assert.isTrue(
        addSharedUsersStub.calledWith(match.any, "mockTitleId", [
          "user1@example.com",
          "user2@example.com",
        ])
      );
    });

    it("returns error when emails are invalid", async () => {
      // Setup common stubs
      sandbox
        .stub(fxCoreDeps, "parseShareAppActionYamlConfig")
        .resolves(ok({ teamsappId: "mockAppId", titleId: "mockTitleId", appId: "mockAppId" }));
      sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
        if (config.name === "env") {
          return ok({ type: "success", result: "dev" });
        } else {
          return ok({ type: "success", result: "" });
        }
      });

      const progressStartStub = sandbox.stub();
      const progressEndStub = sandbox.stub();
      sandbox.stub(tools.ui, "createProgressBar").returns({
        start: progressStartStub,
        end: progressEndStub,
      } as any as IProgressHandler);
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
      sandbox.stub(pathUtils, "getYmlFilePath").returns("m365agents.yml");
      sandbox.stub(fs, "pathExistsSync").returns(true);

      // Case 1: No emails
      const noEmails: Inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        [QuestionNames.ShareOperation]: ShareOperationOption.ShareWithUsers,
        [QuestionNames.ShareScope]: ShareScopeOption.ShareAppWithSpecificUsers,
        [QuestionNames.UserEmail]: "",
      };

      const fxCore = new FxCore(tools);
      const res1 = await fxCore.shareApplication(noEmails);

      chai.assert.isTrue(res1.isErr());
      if (res1.isErr()) {
        chai.assert.instanceOf(res1.error, InputValidationError);
      }

      // Case 2: Too many emails
      const tooManyEmails: Inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        [QuestionNames.ShareOperation]: ShareOperationOption.ShareWithUsers,
        [QuestionNames.ShareScope]: ShareScopeOption.ShareAppWithSpecificUsers,
        [QuestionNames.UserEmail]: Array(MAX_EMAIL_NUMBER + 1)
          .fill("user@example.com")
          .join(","),
      };

      const res2 = await fxCore.shareApplication(tooManyEmails);

      chai.assert.isTrue(res2.isErr());
      if (res2.isErr()) {
        chai.assert.instanceOf(res2.error, InputValidationError);
      }
    });

    it("raw method returns No emails when parsed email list is empty", async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        [QuestionNames.ShareOperation]: ShareOperationOption.RemoveShareAccessFromUsers,
        [QuestionNames.UserEmail]: "   ,  ,",
      };

      const fxCore = new FxCore(tools);
      const shareApplicationRaw = (fxCore.shareApplication as any).original;
      const res = await shareApplicationRaw.call(fxCore, inputs, undefined);

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.instanceOf(res.error, InputValidationError);
        chai.assert.include(res.error.message, "No emails");
      }
    });
  });

  describe("Remove share access", () => {
    it("removes share access successfully", async () => {
      const removeShareAccessStub = sandbox
        .stub(fxCoreDeps, "removeShareAccess")
        .resolves(ok(undefined));

      // Setup common stubs
      sandbox
        .stub(fxCoreDeps, "parseShareAppActionYamlConfig")
        .resolves(ok({ teamsappId: "mockAppId", titleId: "mockTitleId", appId: "mockAppId" }));
      sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
        if (config.name === "env") {
          return ok({ type: "success", result: "dev" });
        } else {
          return ok({ type: "success", result: "" });
        }
      });

      const progressStartStub = sandbox.stub();
      const progressEndStub = sandbox.stub();
      sandbox.stub(tools.ui, "createProgressBar").returns({
        start: progressStartStub,
        end: progressEndStub,
      } as any as IProgressHandler);
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
      sandbox.stub(pathUtils, "getYmlFilePath").returns("m365agents.yml");
      sandbox.stub(fs, "pathExistsSync").returns(true);

      const emails = "user1@example.com,user2@example.com";
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        [QuestionNames.ShareOperation]: ShareOperationOption.RemoveShareAccessFromUsers,
        [QuestionNames.UserEmail]: emails,
      };

      const fxCore = new FxCore(tools);
      const res = await fxCore.shareApplication(inputs);

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(removeShareAccessStub.calledOnce);
      chai.assert.isTrue(
        removeShareAccessStub.calledWith(match.any, "mockTitleId", [
          "user1@example.com",
          "user2@example.com",
        ])
      );
    });
  });

  describe("Error cases", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("returns error for invalid share option", async () => {
      // Setup common stubs
      sandbox
        .stub(fxCoreDeps, "parseShareAppActionYamlConfig")
        .resolves(ok({ teamsappId: "mockAppId", titleId: "mockTitleId", appId: "mockAppId" }));
      sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
        if (config.name === "env") {
          return ok({ type: "success", result: "dev" });
        } else {
          return ok({ type: "success", result: "" });
        }
      });

      // Mock token provider
      sandbox
        .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
        .resolves(ok("mock-token"));

      const progressStartStub = sandbox.stub();
      const progressEndStub = sandbox.stub();
      sandbox.stub(tools.ui, "createProgressBar").returns({
        start: progressStartStub,
        end: progressEndStub,
      } as any as IProgressHandler);
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
      sandbox.stub(pathUtils, "getYmlFilePath").returns("m365agents.yml");
      sandbox.stub(fs, "pathExistsSync").returns(true);

      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        [QuestionNames.ShareOperation]: ShareOperationOption.ShareWithUsers,
        [QuestionNames.ShareScope]: "invalid-option" as any,
      };

      const fxCore = new FxCore(tools);
      const res = await fxCore.shareApplication(inputs);

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.instanceOf(res.error, InputValidationError);
        chai.assert.isTrue(res.error.message.indexOf("Invalid input 'scope'") > -1);
      }
    });

    it("returns error when parse yaml fails", async () => {
      // Setup common stubs
      sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
        if (config.name === "env") {
          return ok({ type: "success", result: "dev" });
        } else {
          return ok({ type: "success", result: "" });
        }
      });
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
      sandbox.stub(pathUtils, "getYmlFilePath").returns("m365agents.yml");
      sandbox.stub(fs, "pathExistsSync").returns(true);

      const parseError = new UserError({
        name: "TestError",
        source: "testSource",
        message: "Failed to parse yaml",
        error: new Error(),
      });

      sandbox.stub(fxCoreDeps, "parseShareAppActionYamlConfig").resolves(err(parseError));

      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        [QuestionNames.ShareOperation]: ShareOperationOption.ShareWithUsers,
        [QuestionNames.ShareScope]: ShareScopeOption.ShareAppWithTenantUsers,
      };

      const fxCore = new FxCore(tools);
      const res = await fxCore.shareApplication(inputs);

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.deepEqual(res.error, parseError);
      }
    });

    it("returns error when token acquisition fails", async () => {
      const tokenError = new UserError({
        name: "TokenError",
        source: "testSource",
        message: "Failed to get token",
        error: new Error(),
      });

      // Setup common stubs
      sandbox
        .stub(fxCoreDeps, "parseShareAppActionYamlConfig")
        .resolves(ok({ teamsappId: "mockAppId", titleId: "mockTitleId", appId: "mockAppId" }));
      // Setup common stubs
      sandbox.stub(metadataUtil, "parse").resolves(ok(mockProjectModel));
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "prod"]));
      sandbox.stub(envUtil, "readEnv").resolves(ok({}));
      sandbox.stub(envUtil, "writeEnv").resolves(ok(undefined));
      sandbox.stub(tools.ui, "selectOption").callsFake(async (config) => {
        if (config.name === "env") {
          return ok({ type: "success", result: "dev" });
        } else {
          return ok({ type: "success", result: "" });
        }
      });
      sandbox.stub(pathUtils, "getEnvFilePath").resolves(ok("."));
      sandbox.stub(pathUtils, "getYmlFilePath").returns("m365agents.yml");
      sandbox.stub(fs, "pathExistsSync").returns(true);

      // Mock token provider to fail
      sandbox
        .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
        .resolves(err(tokenError));

      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        [QuestionNames.ShareOperation]: ShareOperationOption.ShareWithUsers,
        [QuestionNames.ShareScope]: ShareScopeOption.ShareAppWithTenantUsers,
      };

      const fxCore = new FxCore(tools);
      const res = await fxCore.shareApplication(inputs);

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.equal(res.error, tokenError);
      }
    });
  });

  describe("FxCore.removeSharedAccess", () => {
    const projectPath = "./tests/plugins/resource/daTemplate/da-no-action-test-template";

    function stubRemoveSharedAccessBase(options?: {
      tokenResult?: ReturnType<typeof ok> | ReturnType<typeof err>;
      currentUserErr?: UserError;
      userInfoUndefined?: boolean;
      sameUser?: boolean;
      removePermissionErr?: UserError;
    }): void {
      const tokenResult = options?.tokenResult ?? ok("mock-token");
      sandbox
        .stub(fxCoreDeps, "parseShareAppActionYamlConfig")
        .resolves(ok({ teamsappId: "mockAppId", titleId: "mockTitleId", appId: "mockAppId" }));
      if (options?.currentUserErr) {
        sandbox.stub(CollaborationUtil, "getCurrentUserInfo").resolves(err(options.currentUserErr));
      } else {
        sandbox.stub(CollaborationUtil, "getCurrentUserInfo").resolves(
          ok({
            aadId: options?.sameUser ? "target-aad" : "current-aad",
            displayName: "current-user",
            userPrincipalName: "current@example.com",
          } as any)
        );
      }
      sandbox.stub(CollaborationUtil, "getUserInfo").resolves(
        options?.userInfoUndefined
          ? (undefined as any)
          : ({
              aadId: "target-aad",
              displayName: "target-user",
              userPrincipalName: "target@example.com",
            } as any)
      );
      sandbox.stub(teamsDevPortalClient, "removePermission").resolves();
      sandbox
        .stub(PackageService.GetSharedInstance(), "removePermission")
        .resolves(options?.removePermissionErr ? err(options.removePermissionErr) : ok(undefined));
      sandbox
        .stub(tools.tokenProvider.m365TokenProvider, "getAccessToken")
        .resolves(tokenResult as any);
    }

    it("remove shared access happy path", async () => {
      stubRemoveSharedAccessBase();
      const core = new FxCore(tools);
      const result = await core.removeSharedAccess({
        platform: Platform.VSCode,
        projectPath,
        nonInteractive: true,
        [QuestionNames.RemoveUsers]: "user1@example.com,user2@example.com",
      });
      chai.assert.isTrue(result.isOk());
    });

    it("remove shared access - parse error", async () => {
      sandbox
        .stub(fxCoreDeps, "parseShareAppActionYamlConfig")
        .resolves(err(new UserError("mockedSource", "mockedError", "mockedMessage")));
      const core = new FxCore(tools);
      const result = await core.removeSharedAccess({
        platform: Platform.VSCode,
        projectPath,
        nonInteractive: true,
        [QuestionNames.RemoveUsers]: ["user1@example.com"],
      });
      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert.include(result.error.message, "mockedMessage");
      }
    });

    it("remove shared access - token error", async () => {
      stubRemoveSharedAccessBase({
        tokenResult: err(new SystemError("mockedSource", "mockedError", "mockedMessage")),
      });
      const core = new FxCore(tools);
      const result = await core.removeSharedAccess({
        platform: Platform.VSCode,
        projectPath,
        nonInteractive: true,
        [QuestionNames.RemoveUsers]: ["user1@example.com"],
      });
      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert.include(result.error.message, "mockedMessage");
      }
    });

    it("remove shared access - getCurrentUserInfo", async () => {
      stubRemoveSharedAccessBase({
        currentUserErr: new UserError("mockedSource", "mockedError", "mockedMessage"),
      });
      const core = new FxCore(tools);
      const result = await core.removeSharedAccess({
        platform: Platform.VSCode,
        projectPath,
        nonInteractive: true,
        [QuestionNames.RemoveUsers]: ["user1@example.com"],
      });
      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert.include(result.error.message, "mockedMessage");
      }
    });

    it("remove shared access - get user info error", async () => {
      stubRemoveSharedAccessBase({ userInfoUndefined: true });
      const core = new FxCore(tools);
      const result = await core.removeSharedAccess({
        platform: Platform.VSCode,
        projectPath,
        nonInteractive: true,
        [QuestionNames.RemoveUsers]: ["user1@example.com"],
      });
      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert.include(result.error.message, "Invalid user");
      }
    });

    it("remove shared access - remove current user", async () => {
      stubRemoveSharedAccessBase({ sameUser: true });
      const core = new FxCore(tools);
      const result = await core.removeSharedAccess({
        platform: Platform.VSCode,
        projectPath,
        nonInteractive: true,
        [QuestionNames.RemoveUsers]: ["user1@example.com"],
      });
      chai.assert.isTrue(result.isErr());
    });

    it("remove shared access - mos grant permission error", async () => {
      stubRemoveSharedAccessBase({
        removePermissionErr: new UserError("mockedSource", "mockedError", "mockedMessage"),
      });
      const core = new FxCore(tools);
      const result = await core.removeSharedAccess({
        platform: Platform.VSCode,
        projectPath,
        nonInteractive: true,
        [QuestionNames.RemoveUsers]: ["user1@example.com"],
      });
      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert.include(result.error.message, "mockedMessage");
      }
    });
  });
});
