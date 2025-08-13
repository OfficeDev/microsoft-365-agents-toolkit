// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, Inputs, IProgressHandler, ok, Platform, UserError } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import fs from "fs-extra";
import "mocha";
import { createSandbox, match } from "sinon";
import { InputValidationError, MAX_EMAIL_NUMBER, PackageService } from "../../src";
import * as teamsDevPortalClient from "../../src/client/teamsDevPortalClient";
import { ProjectModel } from "../../src/component/configManager/interface";
import * as shareUtils from "../../src/component/driver/share/utils";
import { envUtil } from "../../src/component/utils/envUtil";
import { metadataUtil } from "../../src/component/utils/metadataUtil";
import { pathUtils } from "../../src/component/utils/pathUtils";
import * as collaboratorUtil from "../../src/core/collaborator";
import { FxCore } from "../../src/core/FxCore";
import * as shareCore from "../../src/core/share";
import { QuestionNames } from "../../src/question/questionNames";
import { ShareOperationOption, ShareScopeOption } from "../../src/question/share";
import { MockLogProvider, MockTools } from "./utils";

chai.use(chaiAsPromised);

describe("FxCore.shareApplication", () => {
  const sandbox = createSandbox();
  const tools = new MockTools();
  const logger = new MockLogProvider();
  const mockProjectModel: ProjectModel = {
    version: "1.0.0",
  };

  beforeEach(() => {});

  afterEach(() => {
    sandbox.restore();
  });

  describe("Share with tenant users", () => {
    it("share happy path", async () => {
      const shareWithTenantStub = sandbox
        .stub(shareCore, "shareWithTenant")
        .resolves(ok(undefined));

      sandbox
        .stub(shareUtils, "parseShareAppActionYamlConfig")
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
      sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);
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
      const addSharedUsersStub = sandbox.stub(shareCore, "addSharedUsers").resolves(ok(undefined));

      // Setup common stubs
      sandbox
        .stub(shareUtils, "parseShareAppActionYamlConfig")
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
      sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);

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
        .stub(shareUtils, "parseShareAppActionYamlConfig")
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
      sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);

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
  });

  describe("Remove share access", () => {
    it("removes share access successfully", async () => {
      const removeShareAccessStub = sandbox
        .stub(shareCore, "removeShareAccess")
        .resolves(ok(undefined));

      // Setup common stubs
      sandbox
        .stub(shareUtils, "parseShareAppActionYamlConfig")
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
      sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);

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

  describe("Share with owners", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("shares with owners successfully", async () => {
      // Mock getUserInfo from CollaborationUtil
      sandbox.stub(collaboratorUtil.CollaborationUtil, "getUserInfo").resolves({
        displayName: "Test User",
        userPrincipalName: "testuser@example.com",
        aadId: "mock-aad-id",
        tenantId: "mock-tenant-id",
        isAdministrator: false,
      });

      // Mock teams portal client
      const grantPermissionStub = sandbox
        .stub(teamsDevPortalClient.TeamsDevPortalClient.prototype, "grantPermission")
        .resolves();

      // Mock package service
      sandbox.stub(PackageService, "GetSharedInstance").returns({
        grantPermission: () => ok(undefined),
      } as any);

      // Setup common stubs
      sandbox
        .stub(shareUtils, "parseShareAppActionYamlConfig")
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
      sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);

      const emails = "owner@example.com";
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        [QuestionNames.ShareOperation]: ShareOperationOption.ShareWithUsers,
        [QuestionNames.ShareScope]: ShareScopeOption.ShareAppWithOwners,
        [QuestionNames.UserEmail]: emails,
      };

      const fxCore = new FxCore(tools);
      const res = await fxCore.shareApplication(inputs);

      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(grantPermissionStub.calledOnce);
    });

    it("returns error for invalid user when sharing with owners", async () => {
      // Mock getUserInfo to return undefined (invalid user)
      sandbox.stub(collaboratorUtil.CollaborationUtil, "getUserInfo").resolves(undefined);

      // Setup common stubs
      sandbox
        .stub(shareUtils, "parseShareAppActionYamlConfig")
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
      sandbox.stub(fs, "pathExistsSync").onFirstCall().returns(false).onSecondCall().returns(true);

      const emails = "invalid@example.com";
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: ".",
        [QuestionNames.ShareOperation]: ShareOperationOption.ShareWithUsers,
        [QuestionNames.ShareScope]: ShareScopeOption.ShareAppWithOwners,
        [QuestionNames.UserEmail]: emails,
      };

      const fxCore = new FxCore(tools);
      const res = await fxCore.shareApplication(inputs);

      chai.assert.isTrue(res.isErr());
      if (res.isErr()) {
        chai.assert.instanceOf(res.error, InputValidationError);
        chai.assert.isTrue(res.error.message.indexOf("Invalid user: invalid@example.com") > -1);
      }
    });
  });

  describe("Error cases", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("returns error for invalid share option", async () => {
      // Setup common stubs
      sandbox
        .stub(shareUtils, "parseShareAppActionYamlConfig")
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

      sandbox.stub(shareUtils, "parseShareAppActionYamlConfig").resolves(err(parseError));

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
        .stub(shareUtils, "parseShareAppActionYamlConfig")
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
});
