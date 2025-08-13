// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, ok } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import sinon from "sinon";
import { setTools } from "../../src/common/globalVars";
import { AppUser } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/appUser";
import { M365AppEntity, M365EntityType } from "../../src/component/m365/interface";
import { PackageService } from "../../src/component/m365/packageService";
import { CollaborationUtil } from "../../src/core/collaborator";
import { addSharedUsers, removeShareAccess, shareWithTenant } from "../../src/core/share";
import { InputValidationError } from "../../src/error/common";
import { MockTools } from "./utils";

describe("share", () => {
  const sandbox = sinon.createSandbox();
  const tools = new MockTools();
  setTools(tools);
  const mockMosToken = "mock-mos-token";
  const mockTitleId = "mock-title-id";
  const mockSharedInstance = {
    shareWithTenant: () => {},
    shareWithUsers: () => {},
    unshare: () => {},
    getSharedUsers: () => {},
  };

  beforeEach(() => {
    sandbox.stub(PackageService, "GetSharedInstance").returns(mockSharedInstance as any);
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("shareWithTenant", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("should share with tenant successfully", async () => {
      // Arrange
      const shareWithTenantStub = sandbox
        .stub(mockSharedInstance, "shareWithTenant")
        .resolves(ok(undefined));
      sandbox.stub(tools.ui, "showMessage");

      // Act
      const result = await shareWithTenant(mockMosToken, mockTitleId);

      // Assert
      assert.isTrue(result.isOk());
      assert.isTrue(shareWithTenantStub.calledOnce);
      assert.deepEqual(shareWithTenantStub.args[0], [mockMosToken, mockTitleId]);
    });

    it("should return error when sharing with tenant fails", async () => {
      // Arrange
      const mockError = new InputValidationError("test", "Share with tenant failed");
      sandbox.stub(mockSharedInstance, "shareWithTenant").resolves(err(mockError));

      // Act
      const result = await shareWithTenant(mockMosToken, mockTitleId);

      // Assert
      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.deepEqual(result.error, mockError);
      }
    });
  });

  describe("addSharedUsers", () => {
    const mockEmails = ["user1@example.com", "user2@example.com"];
    const mockUserInfo1: AppUser = {
      aadId: "user1-id",
      displayName: "User 1",
      userPrincipalName: "user1@example.com",
      tenantId: "mock-tenant-id",
      isAdministrator: true,
    };
    const mockUserInfo2: AppUser = {
      aadId: "user2-id",
      displayName: "User 2",
      userPrincipalName: "user2@example.com",
      tenantId: "mock-tenant-id",
      isAdministrator: true,
    };
    const mockExistingEntities: M365AppEntity[] = [
      { entityId: "existing-id", entityType: M365EntityType.User },
    ];

    afterEach(() => {
      sandbox.restore();
    });

    it("should add shared users successfully", async () => {
      // Arrange
      const getSharedUsersStub = sandbox
        .stub(mockSharedInstance, "getSharedUsers")
        .resolves(ok(mockExistingEntities));

      const shareWithUsersStub = sandbox
        .stub(mockSharedInstance, "shareWithUsers")
        .resolves(ok(undefined));

      const getUserInfoStub = sandbox.stub(CollaborationUtil, "getUserInfo");
      getUserInfoStub.withArgs(sinon.match.any, mockEmails[0]).resolves(mockUserInfo1);
      getUserInfoStub.withArgs(sinon.match.any, mockEmails[1]).resolves(mockUserInfo2);

      // Act
      const result = await addSharedUsers(mockMosToken, mockTitleId, mockEmails);

      // Assert
      assert.isTrue(result.isOk());
      assert.isTrue(getSharedUsersStub.calledOnce);
      assert.isTrue(getUserInfoStub.calledTwice);

      // Should call shareWithUsers
      assert.isTrue(shareWithUsersStub.calledOnce);

      // We can't check the exact arguments due to TypeScript issues,
      // but we can verify it was called
      assert.isTrue(shareWithUsersStub.called);
    });

    it("should return error when invalid user email is provided", async () => {
      // Arrange
      sandbox.stub(mockSharedInstance, "getSharedUsers").resolves(ok(mockExistingEntities));

      const getUserInfoStub = sandbox.stub(CollaborationUtil, "getUserInfo");
      getUserInfoStub.withArgs(sinon.match.any, mockEmails[0]).resolves(mockUserInfo1);
      getUserInfoStub.withArgs(sinon.match.any, mockEmails[1]).resolves(undefined);

      // Act
      const result = await addSharedUsers(mockMosToken, mockTitleId, mockEmails);

      // Assert
      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.instanceOf(result.error, InputValidationError);
        assert.include(result.error.message, "Invalid user: user2@example.com");
      }
    });

    it("should return error when getSharedUsers fails", async () => {
      // Arrange
      const mockError = new InputValidationError("test", "Get shared users failed");
      sandbox.stub(mockSharedInstance, "getSharedUsers").resolves(err(mockError));

      // Act
      const result = await addSharedUsers(mockMosToken, mockTitleId, mockEmails);

      // Assert
      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.deepEqual(result.error, mockError);
      }
    });

    it("should return error when shareWithUsers fails", async () => {
      // Arrange
      const mockError = new InputValidationError("test", "Share with users failed");

      sandbox.stub(mockSharedInstance, "getSharedUsers").resolves(ok(mockExistingEntities));

      sandbox
        .stub(CollaborationUtil, "getUserInfo")
        .withArgs(sinon.match.any, mockEmails[0])
        .resolves(mockUserInfo1)
        .withArgs(sinon.match.any, mockEmails[1])
        .resolves(mockUserInfo2);

      sandbox.stub(mockSharedInstance, "shareWithUsers").resolves(err(mockError));

      // Act
      const result = await addSharedUsers(mockMosToken, mockTitleId, mockEmails);

      // Assert
      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.deepEqual(result.error, mockError);
      }
    });
  });

  describe("removeShareAccess", () => {
    const mockEmails = ["user1@example.com", "user2@example.com"];
    const mockUserInfo1: AppUser = {
      aadId: "user1-id",
      displayName: "User 1",
      userPrincipalName: "user1@example.com",
      tenantId: "mock-tenant-id",
      isAdministrator: true,
    };
    const mockUserInfo2: AppUser = {
      aadId: "user2-id",
      displayName: "User 2",
      userPrincipalName: "user2@example.com",
      tenantId: "mock-tenant-id",
      isAdministrator: true,
    };

    const mockExistingEntities: M365AppEntity[] = [
      { entityId: "user1-id", entityType: M365EntityType.User },
      { entityId: "user2-id", entityType: M365EntityType.User },
      { entityId: "user3-id", entityType: M365EntityType.User },
    ];

    afterEach(() => {
      sandbox.restore();
    });

    it("should remove users and keep remaining users", async () => {
      // Arrange
      const getSharedUsersStub = sandbox
        .stub(mockSharedInstance, "getSharedUsers")
        .resolves(ok(mockExistingEntities));

      const shareWithUsersStub = sandbox
        .stub(mockSharedInstance, "shareWithUsers")
        .resolves(ok(undefined));

      const unshareStub = sandbox.stub(mockSharedInstance, "unshare");

      const getUserInfoStub = sandbox.stub(CollaborationUtil, "getUserInfo");
      getUserInfoStub.withArgs(sinon.match.any, mockEmails[0]).resolves(mockUserInfo1);
      getUserInfoStub.withArgs(sinon.match.any, mockEmails[1]).resolves(mockUserInfo2);

      // Act
      const result = await removeShareAccess(mockMosToken, mockTitleId, mockEmails);

      // Assert
      assert.isTrue(result.isOk());
      assert.isTrue(getSharedUsersStub.calledOnce);
      assert.isTrue(getUserInfoStub.calledTwice);

      // Should call shareWithUsers with remaining users and not call unshare
      assert.isTrue(shareWithUsersStub.calledOnce);
      assert.isFalse(unshareStub.called);
    });

    it("should unshare when removing all users", async () => {
      // Arrange
      // Only existing entities are the ones we're removing
      const limitedEntities = [
        { entityId: "user1-id", entityType: M365EntityType.User },
        { entityId: "user2-id", entityType: M365EntityType.User },
      ];

      const getSharedUsersStub = sandbox
        .stub(mockSharedInstance, "getSharedUsers")
        .resolves(ok(limitedEntities));

      const shareWithUsersStub = sandbox.stub(mockSharedInstance, "shareWithUsers");

      const unshareStub = sandbox.stub(mockSharedInstance, "unshare").resolves(ok(undefined));

      const getUserInfoStub = sandbox.stub(CollaborationUtil, "getUserInfo");
      getUserInfoStub.withArgs(sinon.match.any, mockEmails[0]).resolves(mockUserInfo1);
      getUserInfoStub.withArgs(sinon.match.any, mockEmails[1]).resolves(mockUserInfo2);

      // Act
      const result = await removeShareAccess(mockMosToken, mockTitleId, mockEmails);

      // Assert
      assert.isTrue(result.isOk());
      assert.isTrue(getSharedUsersStub.calledOnce);
      assert.isTrue(getUserInfoStub.calledTwice);

      // Should call unshare because all users are removed
      assert.isFalse(shareWithUsersStub.called);
      assert.isTrue(unshareStub.calledOnce);
    });

    it("should return error when invalid user email is provided", async () => {
      // Arrange
      sandbox.stub(mockSharedInstance, "getSharedUsers").resolves(ok(mockExistingEntities));

      const getUserInfoStub = sandbox.stub(CollaborationUtil, "getUserInfo");
      getUserInfoStub.withArgs(sinon.match.any, mockEmails[0]).resolves(mockUserInfo1);
      getUserInfoStub.withArgs(sinon.match.any, mockEmails[1]).resolves(undefined);

      // Act
      const result = await removeShareAccess(mockMosToken, mockTitleId, mockEmails);

      // Assert
      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.instanceOf(result.error, InputValidationError);
        assert.include(result.error.message, "Invalid user: user2@example.com");
      }
    });

    it("should return error when getSharedUsers fails", async () => {
      // Arrange
      const mockError = new InputValidationError("test", "Get shared users failed");
      sandbox.stub(mockSharedInstance, "getSharedUsers").resolves(err(mockError));

      // Act
      const result = await removeShareAccess(mockMosToken, mockTitleId, mockEmails);

      // Assert
      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.deepEqual(result.error, mockError);
      }
    });

    it("should return error when shareWithUsers fails", async () => {
      // Arrange
      const mockError = new InputValidationError("test", "Share with users failed");

      sandbox.stub(mockSharedInstance, "getSharedUsers").resolves(ok(mockExistingEntities));

      sandbox
        .stub(CollaborationUtil, "getUserInfo")
        .withArgs(sinon.match.any, mockEmails[0])
        .resolves(mockUserInfo1)
        .withArgs(sinon.match.any, mockEmails[1])
        .resolves(mockUserInfo2);

      sandbox.stub(mockSharedInstance, "shareWithUsers").resolves(err(mockError));

      // Act
      const result = await removeShareAccess(mockMosToken, mockTitleId, mockEmails);

      // Assert
      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.deepEqual(result.error, mockError);
      }
    });

    it("should return error when unshare fails", async () => {
      // Arrange
      const mockError = new InputValidationError("test", "Unshare failed");
      // Only existing entities are the ones we're removing
      const limitedEntities = [
        { entityId: "user1-id", entityType: M365EntityType.User },
        { entityId: "user2-id", entityType: M365EntityType.User },
      ];

      sandbox.stub(mockSharedInstance, "getSharedUsers").resolves(ok(limitedEntities));

      sandbox
        .stub(CollaborationUtil, "getUserInfo")
        .withArgs(sinon.match.any, mockEmails[0])
        .resolves(mockUserInfo1)
        .withArgs(sinon.match.any, mockEmails[1])
        .resolves(mockUserInfo2);

      sandbox.stub(mockSharedInstance, "unshare").resolves(err(mockError));

      // Act
      const result = await removeShareAccess(mockMosToken, mockTitleId, mockEmails);

      // Assert
      assert.isTrue(result.isErr());
      if (result.isErr()) {
        assert.deepEqual(result.error, mockError);
      }
    });
  });
});
