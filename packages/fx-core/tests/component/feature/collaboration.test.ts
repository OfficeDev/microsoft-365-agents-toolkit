// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError } from "@microsoft/teamsfx-api";
import axios from "axios";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import "mocha";
import { err, ok } from "neverthrow";
import * as sinon from "sinon";
import { GraphClient } from "../../../src/client/graphClient";
import { teamsDevPortalClient } from "../../../src/client/teamsDevPortalClient";
import { setTools } from "../../../src/common/globalVars";
import { AadAppClient } from "../../../src/client/aadAppClient";
import { AppUser } from "../../../src/component/driver/teamsApp/interfaces/appdefinitions/appUser";
import {
  AadCollaboration,
  AgentCollaboration,
  TeamsCollaboration,
} from "../../../src/component/feature/collaboration";
import { M365AppDefinition } from "../../../src/component/m365/interface";
import { PackageService } from "../../../src/component/m365/packageService";
import { MockedM365Provider, MockTools } from "../../core/utils";
import { MockedLogProvider, MockedV2Context } from "../../plugins/solution/util";

chai.use(chaiAsPromised);
const expect = chai.expect;

describe("AadCollaboration", async () => {
  const m365TokenProvider = new MockedM365Provider();
  const logProvider = new MockedLogProvider();
  const aadCollaboration = new AadCollaboration(m365TokenProvider, logProvider);
  const sandbox = sinon.createSandbox();
  const context = new MockedV2Context();
  const expectedObjectId = "00000000-0000-0000-0000-000000000000";
  const expectedUserId = "expectedUserId";

  afterEach(() => {
    sandbox.restore();
  });

  it("grant permission: should add owner", async () => {
    sandbox.stub(AadAppClient.prototype, "addOwner").resolves();

    const result = await aadCollaboration.grantPermission(
      context,
      expectedObjectId,
      expectedUserId
    );
    expect(result.isOk() && result.value[0].resourceId == expectedObjectId).to.be.true;
  });

  it("list collaborator: should return all owners", async () => {
    sandbox.stub(AadAppClient.prototype, "getOwners").resolves([
      {
        resourceId: expectedObjectId,
        displayName: "displayName",
        userPrincipalName: "userPrincipalName",
        userObjectId: expectedUserId,
      },
    ]);

    const result = await aadCollaboration.listCollaborator(context, expectedObjectId);
    expect(result.isOk() && result.value[0].resourceId == expectedObjectId).to.be.true;
  });

  it("check permission: should return owner if user is Microsoft Entra owner", async () => {
    sandbox.stub(AadAppClient.prototype, "getOwners").resolves([
      {
        resourceId: expectedUserId,
        displayName: "displayName",
        userPrincipalName: "userPrincipalName",
        userObjectId: expectedUserId,
      },
    ]);

    const result = await aadCollaboration.checkPermission(
      context,
      expectedObjectId,
      expectedUserId
    );
    expect(result.isOk() && result.value[0].roles![0] == "Owner").to.be.true;
  });

  it("check permission: should return no permission if user is not Microsoft Entra owner", async () => {
    sandbox.stub(AadAppClient.prototype, "getOwners").resolves([
      {
        resourceId: expectedUserId,
        displayName: "displayName",
        userPrincipalName: "userPrincipalName",
        userObjectId: expectedUserId,
      },
    ]);

    const result = await aadCollaboration.checkPermission(context, expectedObjectId, "id");
    expect(result.isOk() && result.value[0].roles![0] == "No Permission").to.be.true;
  });

  it("grant permission errors: should return HttpClientError for 4xx errors", async () => {
    sandbox.stub(AadAppClient.prototype, "addOwner").rejects({
      message: "Request failed with status code 404",
      response: {
        status: 400,
        data: {},
      },
    });
    sandbox.stub(axios, "isAxiosError").returns(true);

    const result = await aadCollaboration.grantPermission(
      context,
      expectedObjectId,
      expectedUserId
    );
    expect(result.isErr() && result.error.name == "HttpClientError").to.be.true;
  });

  it("grant permission errors: should return AppIdNotExist for 404 errors", async () => {
    sandbox.stub(AadAppClient.prototype, "addOwner").rejects({
      message: "Request failed with status code 404",
      response: {
        status: 404,
        data: {},
      },
    });
    sandbox.stub(axios, "isAxiosError").returns(true);

    const result = await aadCollaboration.grantPermission(
      context,
      expectedObjectId,
      expectedUserId
    );
    expect(result.isErr() && result.error.name == "AppIdNotExist").to.be.true;
  });

  it("grant permission errors: should return HttpServerError for 5xx errors", async () => {
    sandbox.stub(AadAppClient.prototype, "addOwner").rejects({
      message: "Request failed with status code 500",
      response: {
        status: 500,
        data: {},
      },
    });
    sandbox.stub(axios, "isAxiosError").returns(true);

    const result = await aadCollaboration.grantPermission(
      context,
      expectedObjectId,
      expectedUserId
    );
    expect(result.isErr() && result.error.name == "HttpServerError").to.be.true;
  });

  it("grant permission errors: should return UnhandledError for unknown errors", async () => {
    sandbox.stub(AadAppClient.prototype, "addOwner").rejects({
      message: "Request failed with status code 500",
      response: {
        status: 500,
        data: {},
      },
    });
    sandbox.stub(axios, "isAxiosError").returns(false);

    const result = await aadCollaboration.grantPermission(
      context,
      expectedObjectId,
      expectedUserId
    );
    expect(result.isErr() && result.error.name == "UnhandledError").to.be.true;
  });

  it("list collaborator errors: should return HttpClientError for 4xx errors", async () => {
    sandbox.stub(AadAppClient.prototype, "getOwners").rejects({
      message: "Request failed with status code 404",
      response: {
        status: 400,
        data: {},
      },
    });
    sandbox.stub(axios, "isAxiosError").returns(true);

    const result = await aadCollaboration.listCollaborator(context, expectedObjectId);
    expect(result.isErr() && result.error.name == "HttpClientError").to.be.true;
  });

  it("list collaborator errors: should return AppIdNotExist for 404 errors", async () => {
    sandbox.stub(AadAppClient.prototype, "getOwners").rejects({
      message: "Request failed with status code 404",
      response: {
        status: 404,
        data: {},
      },
    });
    sandbox.stub(axios, "isAxiosError").returns(true);

    const result = await aadCollaboration.listCollaborator(context, expectedObjectId);
    expect(result.isErr() && result.error.name == "AppIdNotExist").to.be.true;
  });

  it("list collaborator errors: should return HttpServerError for 5xx errors", async () => {
    sandbox.stub(AadAppClient.prototype, "getOwners").rejects({
      message: "Request failed with status code 500",
      response: {
        status: 500,
        data: {},
      },
    });
    sandbox.stub(axios, "isAxiosError").returns(true);

    const result = await aadCollaboration.listCollaborator(context, expectedObjectId);
    expect(result.isErr() && result.error.name == "HttpServerError").to.be.true;
  });

  it("list collaborator errors: should return UnhandledError for unknown errors", async () => {
    sandbox.stub(AadAppClient.prototype, "getOwners").rejects({
      message: "Request failed with status code 500",
      response: {
        status: 500,
        data: {},
      },
    });
    sandbox.stub(axios, "isAxiosError").returns(false);

    const result = await aadCollaboration.listCollaborator(context, expectedObjectId);
    expect(result.isErr() && result.error.name == "UnhandledError").to.be.true;
  });

  it("check permission errors: should return HttpClientError for 4xx errors", async () => {
    sandbox.stub(AadAppClient.prototype, "getOwners").rejects({
      message: "Request failed with status code 404",
      response: {
        status: 400,
        data: {},
      },
    });
    sandbox.stub(axios, "isAxiosError").returns(true);

    const result = await aadCollaboration.checkPermission(
      context,
      expectedObjectId,
      expectedUserId
    );
    expect(result.isErr() && result.error.name == "HttpClientError").to.be.true;
  });

  it("check permission errors: should return AppIdNotExist for 404 errors", async () => {
    sandbox.stub(AadAppClient.prototype, "getOwners").rejects({
      message: "Request failed with status code 404",
      response: {
        status: 404,
        data: {},
      },
    });
    sandbox.stub(axios, "isAxiosError").returns(true);

    const result = await aadCollaboration.checkPermission(
      context,
      expectedObjectId,
      expectedUserId
    );
    expect(result.isErr() && result.error.name == "AppIdNotExist").to.be.true;
  });

  it("check permission errors: should return HttpServerError for 5xx errors", async () => {
    sandbox.stub(AadAppClient.prototype, "getOwners").rejects({
      message: "Request failed with status code 500",
      response: {
        status: 500,
        data: {},
      },
    });
    sandbox.stub(axios, "isAxiosError").returns(true);

    const result = await aadCollaboration.checkPermission(
      context,
      expectedObjectId,
      expectedUserId
    );
    expect(result.isErr() && result.error.name == "HttpServerError").to.be.true;
  });

  it("check permission errors: should return UnhandledError for unknown errors", async () => {
    sandbox.stub(AadAppClient.prototype, "getOwners").rejects({
      message: "Request failed with status code 500",
      response: {
        status: 500,
        data: {},
      },
    });
    sandbox.stub(axios, "isAxiosError").returns(false);

    const result = await aadCollaboration.checkPermission(
      context,
      expectedObjectId,
      expectedUserId
    );
    expect(result.isErr() && result.error.name == "UnhandledError").to.be.true;
  });
});

describe("TeamsCollaboration", async () => {
  const context = new MockedV2Context();
  const m365TokenProvider = new MockedM365Provider();
  const teamsCollaboration = new TeamsCollaboration(m365TokenProvider);
  const sandbox = sinon.createSandbox();
  const expectedAppId = "00000000-0000-0000-0000-000000000000";
  const expectedUserId = "expectedUserId";
  const expectedUserInfo: AppUser = {
    tenantId: "tenantId",
    aadId: expectedUserId,
    displayName: "displayName",
    userPrincipalName: "userPrincipalName",
    isAdministrator: true,
  };

  afterEach(() => {
    sandbox.restore();
  });

  it("grant permission: should add owner", async () => {
    sandbox.stub(teamsDevPortalClient, "grantPermission").resolves();

    const result = await teamsCollaboration.grantPermission(
      context,
      expectedAppId,
      expectedUserInfo
    );
    expect(result.isOk() && result.value[0].resourceId == expectedAppId).to.be.true;
  });

  it("list collaborator: should return all owners", async () => {
    sandbox.stub(teamsDevPortalClient, "getUserList").resolves([expectedUserInfo]);

    const result = await teamsCollaboration.listCollaborator(context, expectedAppId);
    expect(result.isOk() && result.value[0].resourceId == expectedAppId).to.be.true;
  });

  it("check permission: should return admin if user is teams app owner", async () => {
    sandbox.stub(teamsDevPortalClient, "checkPermission").resolves("Administrator");

    const result = await teamsCollaboration.checkPermission(
      context,
      expectedAppId,
      expectedUserInfo
    );
    expect(result.isOk() && result.value[0].roles![0] == "Administrator").to.be.true;
  });

  it("check permission: should return no permission if user is not Microsoft Entra owner", async () => {
    sandbox.stub(teamsDevPortalClient, "checkPermission").resolves("No permission");

    const result = await teamsCollaboration.checkPermission(
      context,
      expectedAppId,
      expectedUserInfo
    );
    expect(result.isOk() && result.value[0].roles![0] == "No permission").to.be.true;
  });

  it("list collaborator errors: should return HttpClientError for 4xx errors", async () => {
    sandbox.stub(teamsDevPortalClient, "getUserList").rejects({
      innerError: {
        message: "Request failed with status code 400",
        response: {
          status: 400,
          data: {},
        },
      },
    });

    const result = await teamsCollaboration.listCollaborator(context, expectedAppId);
    expect(result.isErr() && result.error.name == "HttpClientError").to.be.true;
  });

  it("list collaborator errors: should return AppIdNotExist for 404 errors", async () => {
    sandbox.stub(teamsDevPortalClient, "getUserList").rejects({
      innerError: {
        message: "Request failed with status code 404",
        response: {
          status: 404,
          data: {},
        },
      },
    });

    const result = await teamsCollaboration.listCollaborator(context, expectedAppId);
    expect(result.isErr() && result.error.name == "AppIdNotExist").to.be.true;
  });

  it("list collaborator errors: should return HttpServerError for 5xx errors", async () => {
    sandbox.stub(teamsDevPortalClient, "getUserList").rejects({
      innerError: {
        message: "Request failed with status code 500",
        response: {
          status: 500,
          data: {},
        },
      },
    });

    const result = await teamsCollaboration.listCollaborator(context, expectedAppId);
    expect(result.isErr() && result.error.name == "HttpServerError").to.be.true;
  });

  it("list collaborator errors: should return unhandledErrors", async () => {
    sandbox.stub(teamsDevPortalClient, "getUserList").rejects({
      message: "Request failed with status code 500",
    });

    const result = await teamsCollaboration.listCollaborator(context, expectedAppId);
    expect(result.isErr() && result.error.name == "UnhandledError").to.be.true;
  });

  it("grant permission errors: should return HttpClientError for 4xx errors", async () => {
    sandbox.stub(teamsDevPortalClient, "grantPermission").rejects({
      innerError: {
        message: "Request failed with status code 400",
        response: {
          status: 400,
          data: {},
        },
      },
    });

    const result = await teamsCollaboration.grantPermission(
      context,
      expectedAppId,
      expectedUserInfo
    );
    expect(result.isErr() && result.error.name == "HttpClientError").to.be.true;
  });

  it("grant permission errors: should return AppIdNotExist for 404 errors", async () => {
    sandbox.stub(teamsDevPortalClient, "grantPermission").rejects({
      innerError: {
        message: "Request failed with status code 404",
        response: {
          status: 404,
          data: {},
        },
      },
    });

    const result = await teamsCollaboration.grantPermission(
      context,
      expectedAppId,
      expectedUserInfo
    );
    expect(result.isErr() && result.error.name == "AppIdNotExist").to.be.true;
  });

  it("grant permission errors: should return HttpServerError for 5xx errors", async () => {
    sandbox.stub(teamsDevPortalClient, "grantPermission").rejects({
      innerError: {
        message: "Request failed with status code 500",
        response: {
          status: 500,
          data: {},
        },
      },
    });

    const result = await teamsCollaboration.grantPermission(
      context,
      expectedAppId,
      expectedUserInfo
    );
    expect(result.isErr() && result.error.name == "HttpServerError").to.be.true;
  });

  it("grant permission errors: should return unhandledErrors", async () => {
    sandbox.stub(teamsDevPortalClient, "grantPermission").rejects({
      message: "Request failed with status code 500",
    });

    const result = await teamsCollaboration.grantPermission(
      context,
      expectedAppId,
      expectedUserInfo
    );
    expect(result.isErr() && result.error.name == "UnhandledError").to.be.true;
  });

  it("check permission errors: should return HttpClientError for 4xx errors", async () => {
    sandbox.stub(teamsDevPortalClient, "checkPermission").rejects({
      innerError: {
        message: "Request failed with status code 400",
        response: {
          status: 400,
          data: {},
        },
      },
    });

    const result = await teamsCollaboration.checkPermission(
      context,
      expectedAppId,
      expectedUserInfo
    );
    expect(result.isErr() && result.error.name == "HttpClientError").to.be.true;
  });

  it("check permission errors: should return AppIdNotExist for 404 errors", async () => {
    sandbox.stub(teamsDevPortalClient, "checkPermission").rejects({
      innerError: {
        message: "Request failed with status code 404",
        response: {
          status: 404,
          data: {},
        },
      },
    });

    const result = await teamsCollaboration.checkPermission(
      context,
      expectedAppId,
      expectedUserInfo
    );
    expect(result.isErr() && result.error.name == "AppIdNotExist").to.be.true;
  });

  it("check permission errors: should return HttpServerError for 5xx errors", async () => {
    sandbox.stub(teamsDevPortalClient, "checkPermission").rejects({
      innerError: {
        message: "Request failed with status code 500",
        response: {
          status: 500,
          data: {},
        },
      },
    });

    const result = await teamsCollaboration.checkPermission(
      context,
      expectedAppId,
      expectedUserInfo
    );
    expect(result.isErr() && result.error.name == "HttpServerError").to.be.true;
  });

  it("check permission errors: should return unhandledErrors", async () => {
    sandbox.stub(teamsDevPortalClient, "checkPermission").rejects({
      message: "Request failed with status code 500",
    });

    const result = await teamsCollaboration.checkPermission(
      context,
      expectedAppId,
      expectedUserInfo
    );
    expect(result.isErr() && result.error.name == "UnhandledError").to.be.true;
  });
});

describe("AgentCollaboration", async () => {
  const context = new MockedV2Context();
  const m365TokenProvider = new MockedM365Provider();
  setTools(new MockTools());
  const agentCollaboration = new AgentCollaboration(m365TokenProvider);
  const sandbox = sinon.createSandbox();
  const expectedTitleId = "test-title-id";
  const expectedUserId = "expectedUserId";
  const expectedUserInfo: AppUser = {
    tenantId: "tenantId",
    aadId: expectedUserId,
    displayName: "displayName",
    userPrincipalName: "userPrincipalName",
    isAdministrator: true,
  };

  afterEach(() => {
    sandbox.restore();
  });

  it("grant permission: should add owner", async () => {
    sandbox.stub(m365TokenProvider, "getAccessToken").resolves(ok("test-token"));

    const packageServiceStub = sandbox.stub(PackageService.GetSharedInstance(), "addOwner");
    packageServiceStub.resolves(ok(undefined));

    const result = await agentCollaboration.grantPermission(
      context,
      expectedTitleId,
      expectedUserInfo
    );
    expect(result.isOk() && result.value[0].resourceId == expectedTitleId).to.be.true;
    expect(result.isOk() && result.value[0].roles![0] == "Owner").to.be.true;
  });

  it("list collaborator: should return all owners", async () => {
    sandbox.stub(m365TokenProvider, "getAccessToken").resolves(ok("test-token"));

    const packageServiceStub = sandbox.stub(PackageService.GetSharedInstance(), "previewApp");
    packageServiceStub.resolves(
      ok({
        owners: [
          {
            entityId: expectedUserId,
            entityType: "User",
          },
        ],
      } as M365AppDefinition)
    );

    sandbox.stub(GraphClient.prototype, "getUserInfoFromId").resolves({
      id: expectedUserId,
      displayName: "displayName",
      userPrincipalName: "userPrincipalName",
      mail: "test@mail.com",
    });

    const result = await agentCollaboration.listCollaborator(context, expectedTitleId);
    expect(result.isOk() && result.value[0].resourceId == expectedTitleId).to.be.true;
    expect(result.isOk() && result.value[0].userObjectId == expectedUserId).to.be.true;
  });

  it("grant permission errors: should return error from token provider", async () => {
    const expectedError = new Error("token error");
    sandbox.stub(m365TokenProvider, "getAccessToken").resolves(err(expectedError as FxError));

    const result = await agentCollaboration.grantPermission(
      context,
      expectedTitleId,
      expectedUserInfo
    );
    expect(result.isErr() && result.error === expectedError).to.be.true;
  });

  it("grant permission errors: should return error from package service", async () => {
    sandbox.stub(m365TokenProvider, "getAccessToken").resolves(ok("test-token"));

    const expectedError = new Error("package service error");
    const packageServiceStub = sandbox.stub(PackageService.GetSharedInstance(), "addOwner");
    packageServiceStub.resolves(err(expectedError as FxError));

    const result = await agentCollaboration.grantPermission(
      context,
      expectedTitleId,
      expectedUserInfo
    );
    expect(result.isErr() && result.error === expectedError).to.be.true;
  });

  it("list collaborator: should skip users when getUserInfoFromId returns undefined", async () => {
    sandbox.stub(m365TokenProvider, "getAccessToken").resolves(ok("test-token"));

    const packageServiceStub = sandbox.stub(PackageService.GetSharedInstance(), "previewApp");
    packageServiceStub.resolves(
      ok({
        owners: [
          {
            entityId: expectedUserId,
            entityType: "User",
          },
        ],
      } as M365AppDefinition)
    );

    sandbox.stub(GraphClient.prototype, "getUserInfoFromId").resolves(undefined);

    const result = await agentCollaboration.listCollaborator(context, expectedTitleId);
    expect(result.isOk() && result.value.length === 0).to.be.true;
  });

  it("list collaborator errors: should return error from token provider", async () => {
    const expectedError = new Error("token error");
    sandbox.stub(m365TokenProvider, "getAccessToken").resolves(err(expectedError as FxError));

    const result = await agentCollaboration.listCollaborator(context, expectedTitleId);
    expect(result.isErr() && result.error === expectedError).to.be.true;
  });

  it("list collaborator errors: should return error from package service", async () => {
    sandbox.stub(m365TokenProvider, "getAccessToken").resolves(ok("test-token"));

    const expectedError = new Error("package service error");
    const packageServiceStub = sandbox.stub(PackageService.GetSharedInstance(), "previewApp");
    packageServiceStub.resolves(err(expectedError as FxError));

    const result = await agentCollaboration.listCollaborator(context, expectedTitleId);
    expect(result.isErr() && result.error === expectedError).to.be.true;
  });
});
