// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { TeamsAppManifest, err, ok } from "@microsoft/teamsfx-api";
import axios, { AxiosResponse } from "axios";
import * as chai from "chai";
import "mocha";
import mockedEnv from "mocked-env";
import { createSandbox } from "sinon";
import { v4 as uuid } from "uuid";
import { RetryHandler, teamsDevPortalClient } from "../../src/client/teamsDevPortalClient";
import { setTools } from "../../src/common/globalVars";
import * as telemetry from "../../src/common/telemetry";
import { Constants, ErrorMessages } from "../../src/component/driver/teamsApp/constants";
import { AppStudioError } from "../../src/component/driver/teamsApp/errors";
import {
  ApiSecretRegistration,
  ApiSecretRegistrationAppType,
  ApiSecretRegistrationUpdate,
} from "../../src/component/driver/teamsApp/interfaces/ApiSecretRegistration";
import { AsyncAppValidationStatus } from "../../src/component/driver/teamsApp/interfaces/AsyncAppValidationResponse";
import {
  OauthRegistration,
  OauthRegistrationAppType,
  OauthRegistrationTargetAudience,
  OauthRegistrationUserAccessType,
} from "../../src/component/driver/teamsApp/interfaces/OauthRegistration";
import { PublishingState } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/IPublishingAppDefinition";
import { AppDefinition } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
import { AppUser } from "../../src/component/driver/teamsApp/interfaces/appdefinitions/appUser";
import { AppStudioResultFactory } from "../../src/component/driver/teamsApp/results";
import { manifestUtils } from "../../src/component/driver/teamsApp/utils/ManifestUtils";
import { IBotRegistration } from "../../src/component/resource/botService/appStudio/interfaces/IBotRegistration";
import { ErrorNames } from "../../src/component/resource/botService/constants";
import {
  DeveloperPortalAPIFailedSystemError,
  DeveloperPortalAPIFailedUserError,
} from "../../src/error/teamsApp";
import { Messages } from "../component/resource/botService/messages";
import { MockTools } from "../core/utils";
import { getDefaultString } from "../../src/common/localizeUtils";
import { HelpLinks } from "../../src/common/constants";
import { expect } from "chai";

describe("TeamsDevPortalClient Test", () => {
  const tools = new MockTools();
  const sandbox = createSandbox();
  setTools(tools);
  const token = "appStudioToken";
  const appDef: AppDefinition = {
    appName: "fake",
    teamsAppId: uuid(),
    userList: [],
  };

  const appApiRegistration: ApiSecretRegistration = {
    id: "fakeId",
    description: "An Api Key registration for auth",
    clientSecrets: [
      {
        id: uuid(),
        value: "fakeValue",
        isValueRedacted: false,
      },
    ],
    applicableToApps: ApiSecretRegistrationAppType.AnyApp,
    targetUrlsShouldStartWith: ["https://www.example.com"],
  };

  const fakeOauthRegistration: OauthRegistration = {
    description: "fake-description",
    scopes: ["fake-scope"],
    clientId: "fake-client-id",
    clientSecret: "fake-client-secret",
    authorizationEndpoint: "fake-authorization-url",
    tokenExchangeEndpoint: "fake-token-endpoint",
    tokenRefreshEndpoint: "fake-refresh-endpoint",
    applicableToApps: OauthRegistrationAppType.AnyApp,
    targetAudience: OauthRegistrationTargetAudience.AnyTenant,
    manageableByUsers: [
      {
        userId: "fake-user-id",
        accessType: OauthRegistrationUserAccessType.ReadWrite,
      },
    ],
    targetUrlsShouldStartWith: ["fake-domain"],
  };

  const sampleBot: IBotRegistration = {
    botId: "00000000-0000-0000-0000-000000000000",
    name: "ttttttt-local-debug",
    description: "",
    iconUrl:
      "https://docs.botframework.com/static/devportal/client/images/bot-framework-default.png",
    messagingEndpoint: "https://1111-222-222-333-44.ngrok.io/api/messages",
    callingEndpoint: "",
  };
  beforeEach(() => {
    sandbox.stub(RetryHandler, "RETRIES").value(1);
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("setRegionEndpointByToken", () => {
    it("Happy path", async () => {
      sandbox.stub(RetryHandler, "Retry").resolves({
        status: 200,
        data: {
          regionGtms: {
            teamsDevPortal: "https://xxx.xxx.xxx",
          },
        },
      });
      await teamsDevPortalClient.setRegionEndpointByToken("https://xxx.xxx.xxx");
      chai.assert.equal(teamsDevPortalClient.regionEndpoint, "https://xxx.xxx.xxx");
    });
    it("Not set region for int endpoint", async () => {
      teamsDevPortalClient.regionEndpoint = undefined;
      const restore = mockedEnv({
        APP_STUDIO_ENV: "int",
      });
      await teamsDevPortalClient.setRegionEndpointByToken("https://xxx.xxx.xxx");
      chai.assert.isUndefined(teamsDevPortalClient.regionEndpoint);
      restore();
    });
  });
  describe("publishTeamsApp", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      const response = {
        data: {
          id: "fakeId",
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").resolves(response);

      const res = await teamsDevPortalClient.publishTeamsApp(token, "fakeId", Buffer.from(""));
      chai.assert.equal(res, response.data.id);
    });
    it("return undefined response", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      sandbox.stub(fakeAxiosInstance, "post").resolves(undefined);
      try {
        await teamsDevPortalClient.publishTeamsApp(token, "fakeId", Buffer.from(""));
      } catch (e) {
        chai.assert.equal(e.name, DeveloperPortalAPIFailedSystemError.name);
        chai.assert.isTrue(e.message.includes(AppStudioError.TeamsAppPublishFailedError.name));
      }
    });
    it("return no data", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      const response = {};
      sandbox.stub(fakeAxiosInstance, "post").resolves(response);
      try {
        await teamsDevPortalClient.publishTeamsApp(token, "fakeId", Buffer.from(""));
      } catch (e) {
        chai.assert.equal(e.name, DeveloperPortalAPIFailedSystemError.name);
        chai.assert.isTrue(e.message.includes(AppStudioError.TeamsAppPublishFailedError.name));
      }
    });
    it("return no data with correlation id", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      const xCorrelationId = "fakeCorrelationId";
      const response = {
        headers: {
          "x-correlation-id": xCorrelationId,
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").resolves(response);
      try {
        await teamsDevPortalClient.publishTeamsApp(token, "fakeId", Buffer.from(""));
      } catch (e) {
        chai.assert.equal(e.name, DeveloperPortalAPIFailedSystemError.name);
        chai.assert.isTrue(e.message.includes(AppStudioError.TeamsAppPublishFailedError.name));
        chai.assert.isTrue(e.message.includes(xCorrelationId));
      }
    });
    it("API Failure", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "error",
        message: "fake message",
      };
      sandbox.stub(fakeAxiosInstance, "post").throws(error);

      try {
        await teamsDevPortalClient.publishTeamsApp(token, "fakeId", Buffer.from(""));
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });

    it("should contain x-correlation-id on BadeRequest with 2xx status code", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const xCorrelationId = "fakeCorrelationId";
      const response = {
        data: {
          error: "BadRequest",
        },
        message: "fake message",
        headers: {
          "x-correlation-id": xCorrelationId,
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").resolves(response);

      try {
        await teamsDevPortalClient.publishTeamsApp(token, "fakeId", Buffer.from(""));
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
        chai.assert.include(error.message, xCorrelationId);
      }
    });

    it("Bad gateway", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const postResponse = {
        data: {
          error: {
            code: "BadGateway",
            message: "fakeMessage",
          },
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").resolves(postResponse);

      const getResponse = {
        data: {
          value: [
            {
              appDefinitions: [
                {
                  lastModifiedDateTime: new Date(),
                  publishingState: PublishingState.submitted,
                  teamsAppId: uuid(),
                  displayName: "fakeApp",
                },
              ],
            },
          ],
        },
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(getResponse);

      const res = await teamsDevPortalClient.publishTeamsApp(token, "fakeId", Buffer.from(""));
      chai.assert.equal(res, getResponse.data.value[0].appDefinitions[0].teamsAppId);
    });

    it("AppdefinitionsAlreadyExists - update", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const publishResponse = {
        data: {
          error: {
            code: "Conflict",
            message: "Conflict",
            innerError: {
              code: "AppDefinitionAlreadyExists",
            },
          },
        },
      };

      const updateResponse = {
        data: {
          teamsAppId: "fakeId",
        },
      };
      sandbox
        .stub(fakeAxiosInstance, "post")
        .onFirstCall()
        .resolves(publishResponse)
        .onSecondCall()
        .resolves(updateResponse);
      sandbox.stub(teamsDevPortalClient, "publishTeamsAppUpdate").resolves("fakeId");

      const getResponse = {
        data: {
          value: [
            {
              appDefinitions: [
                {
                  lastModifiedDateTime: new Date(),
                  publishingState: PublishingState.submitted,
                  teamsAppId: uuid(),
                  displayName: "fakeApp",
                },
              ],
            },
          ],
        },
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(getResponse);

      const res = await teamsDevPortalClient.publishTeamsApp(token, "fakeId", Buffer.from(""));
      chai.assert.equal(res, "fakeId");
    });

    it("AppdefinitionsAlreadyExists - failed", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const postResponse = {
        data: {
          error: {
            code: "Conflict",
            message: "Conflict",
            innerError: {
              code: "AppDefinitionAlreadyExists",
            },
          },
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").resolves(postResponse);

      try {
        await teamsDevPortalClient.publishTeamsApp(token, "fakeId", Buffer.from(""));
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
        chai.assert.isTrue(
          error.message.includes(AppStudioError.TeamsAppPublishConflictError.name)
        );
      }
    });
  });

  describe("import Teams app", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: appDef,
      };
      sandbox.stub(fakeAxiosInstance, "post").resolves(response);

      teamsDevPortalClient.regionEndpoint = "https://dev.teams.microsoft.com/amer";

      const res = await teamsDevPortalClient.importApp(token, Buffer.from(""));
      chai.assert.equal(res, appDef);
    });

    it("Happy path - with wrong region", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: appDef,
      };
      sandbox.stub(fakeAxiosInstance, "post").resolves(response);
      teamsDevPortalClient.regionEndpoint = "https://dev.teams.microsoft.com";
      const res = await teamsDevPortalClient.importApp(token, Buffer.from(""));
      chai.assert.equal(res, appDef);
    });

    it("409 conflict", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        response: {
          status: 409,
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").throws(error);

      try {
        await teamsDevPortalClient.importApp(token, Buffer.from(""));
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedUserError.name);
        chai.assert.isTrue(error.message.includes(AppStudioError.TeamsAppCreateConflictError.name));
        chai.assert.equal(error.helpLink, HelpLinks.SwitchTenant);
      }
    });

    it("422 conflict", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        response: {
          status: 422,
          data: "Unable import, App already exists and published. publishStatus: 'LobStore'",
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").throws(error);

      try {
        await teamsDevPortalClient.importApp(token, Buffer.from(""));
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedUserError.name);
        chai.assert.isTrue(
          error.message.includes(AppStudioError.TeamsAppCreateConflictWithPublishedAppError.name)
        );
      }
    });

    it("422 conflict with unknown data", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        response: {
          status: 422,
          data: "Unknown",
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").throws(error);

      try {
        await teamsDevPortalClient.importApp(token, Buffer.from(""));
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
        chai.assert.isFalse(
          error.message.includes(AppStudioError.TeamsAppCreateConflictWithPublishedAppError.name)
        );
        chai.assert.isTrue(
          error.message.includes(getDefaultString("error.appstudio.apiFailed.name.common"))
        );
      }
    });

    it("422 other error", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        response: {
          status: 422,
          data: "fake error message",
          headers: {
            "x-correlation-id": uuid(),
          },
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").throws(error);

      try {
        await teamsDevPortalClient.importApp(token, Buffer.from(""));
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });

    it("invalid Teams app id", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      sandbox
        .stub(manifestUtils, "extractManifestFromArchivedFile")
        .returns(ok(new TeamsAppManifest()));

      const error = {
        response: {
          status: 400,
          data: "App Id must be a GUID",
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").throws(error);

      try {
        await teamsDevPortalClient.importApp(token, Buffer.from(""));
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedUserError.name);
        chai.assert.isTrue(error.message.includes(AppStudioError.InvalidTeamsAppIdError.name));
      }
    });

    it("extract manifet failed", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      const fileNotFoundError = AppStudioResultFactory.UserError(
        AppStudioError.FileNotFoundError.name,
        AppStudioError.FileNotFoundError.message(Constants.MANIFEST_FILE)
      );
      sandbox
        .stub(manifestUtils, "extractManifestFromArchivedFile")
        .returns(err(fileNotFoundError));

      const error = {
        response: {
          status: 400,
          data: "App Id must be a GUID",
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").throws(error);

      try {
        await teamsDevPortalClient.importApp(token, Buffer.from(""));
      } catch (error) {
        chai.assert.equal(error.name, AppStudioError.FileNotFoundError.name);
      }
    });

    it("400 bad reqeust", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        response: {
          staus: 400,
          data: "BadRequest",
          headers: {
            "x-correlation-id": uuid(),
          },
        },
        message: "fake message",
      };
      sandbox.stub(fakeAxiosInstance, "post").throws(error);

      try {
        await teamsDevPortalClient.importApp(token, Buffer.from(""));
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });

    it("return error when no response data", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const res = {
        response: {
          staus: 200,
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").resolves(res);

      try {
        await teamsDevPortalClient.importApp(token, Buffer.from(""));
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });
  });

  describe("getApp", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: appDef,
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);

      const res = await teamsDevPortalClient.getApp(token, appDef.teamsAppId!);
      chai.assert.equal(res, appDef);
    });

    it("404 not found", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "404",
        message: "fake message",
      };
      sandbox.stub(fakeAxiosInstance, "get").throws(error);

      try {
        await teamsDevPortalClient.getApp(token, appDef.teamsAppId!);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });

    it("region - 404", async () => {
      teamsDevPortalClient.regionEndpoint = "https://dev.teams.microsoft.com/amer";
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        response: {
          status: 404,
          headers: {
            "x-correlation-id": "fakeCorrelationId",
          },
        },
      };
      sandbox.stub(fakeAxiosInstance, "get").throws(error);

      try {
        await teamsDevPortalClient.getApp(token, appDef.teamsAppId!);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      } finally {
        teamsDevPortalClient.setRegionEndpoint(undefined as unknown as string);
      }
    });

    it("app id not match", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: appDef,
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      try {
        await teamsDevPortalClient.getApp(token, "anotherId");
      } catch (e) {
        chai.assert.isTrue(e.message.includes("cannot get the app definition with app ID"));
      }
    });
  });
  describe("getStaggedApp", () => {
    it("happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      const response = {
        data: {
          value: [
            {
              appDefinitions: [
                {
                  publishingState: PublishingState.submitted,
                  teamsAppId: "xx",
                  displayName: "xx",
                  lastModifiedDateTime: null,
                },
              ],
            },
          ],
        },
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      const res = await teamsDevPortalClient.getStaggedApp(token, "fake");
      chai.assert.equal(res?.teamsAppId, "xx");
    });
    it("not found", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      const response = {
        data: {
          value: [],
        },
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      const res = await teamsDevPortalClient.getStaggedApp(token, "fake");
      chai.assert.isUndefined(res);
    });
  });
  describe("getAppPackage", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: "fakeData",
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);

      const res = await teamsDevPortalClient.getAppPackage(token, appDef.teamsAppId!);
      chai.assert.equal(res, "fakeData");
    });

    it("404 not found", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "404",
        message: "fake message",
      };
      sandbox.stub(fakeAxiosInstance, "get").throws(error);

      try {
        await teamsDevPortalClient.getAppPackage(token, appDef.teamsAppId!);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });

    it("No data", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      const response = {
        data: undefined,
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      try {
        await teamsDevPortalClient.getAppPackage(token, appDef.teamsAppId!);
      } catch (e) {
        chai.assert.isTrue(e instanceof DeveloperPortalAPIFailedSystemError);
      }
    });
  });

  describe("partner center app validation", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: {
          status: "Accepted",
          errors: [],
          warnings: [],
          notes: [],
          addInDetails: {
            displayName: "fakeApp",
            developerName: "Teams",
            version: "0.0.1",
            manifestVersion: "1.16",
          },
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").resolves(response);

      const res = await teamsDevPortalClient.partnerCenterAppPackageValidation(
        token,
        Buffer.from("")
      );
      chai.assert.equal(res, response.data);
    });

    it("422", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "422",
        message: "Invalid zip",
      };
      sandbox.stub(fakeAxiosInstance, "post").throws(error);

      try {
        await teamsDevPortalClient.partnerCenterAppPackageValidation(token, Buffer.from(""));
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });
  });

  describe("Check exists in tenant", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: true,
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);

      const res = await teamsDevPortalClient.checkExistsInTenant(token, appDef.teamsAppId!);
      chai.assert.isTrue(res);
    });
    it("data false", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: false,
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);

      const res = await teamsDevPortalClient.checkExistsInTenant(token, appDef.teamsAppId!);
      chai.assert.isFalse(res);
    });
    it("404 not found", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "404",
        message: "fake message",
      };
      sandbox.stub(fakeAxiosInstance, "get").throws(error);

      try {
        await teamsDevPortalClient.checkExistsInTenant(token, appDef.teamsAppId!);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });
  });

  describe("publishTeamsAppUpdate", () => {
    it("Happy path", async () => {
      sandbox.stub(teamsDevPortalClient, "getStaggedApp").resolves({
        publishingState: PublishingState.submitted,
        teamsAppId: "xx",
        displayName: "xx",
        lastModifiedDateTime: null,
      });
      sandbox.stub(RetryHandler, "Retry").resolves({ data: { teamsAppId: "xx" } });
      const res = await teamsDevPortalClient.publishTeamsAppUpdate(token, "", Buffer.from(""));
      chai.assert.equal(res, "xx");
    });
    it("return no data", async () => {
      sandbox.stub(teamsDevPortalClient, "getStaggedApp").resolves({
        publishingState: PublishingState.submitted,
        teamsAppId: "xx",
        displayName: "xx",
        lastModifiedDateTime: null,
      });
      sandbox.stub(RetryHandler, "Retry").resolves({ data: { teamsAppId: "xx" } });
      try {
        await teamsDevPortalClient.publishTeamsAppUpdate(token, "", Buffer.from(""));
      } catch (e) {
        chai.assert.isTrue(e.name === AppStudioError.TeamsAppPublishFailedError.name);
      }
    });
    it("should contain x-correlation-id on BadeRequest with 2xx status code", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const xCorrelationId = "fakeCorrelationId";
      const postResponse = {
        data: {
          error: "BadRequest",
        },
        message: "fake message",
        headers: {
          "x-correlation-id": xCorrelationId,
        },
      };

      sandbox.stub(fakeAxiosInstance, "post").resolves(postResponse);

      const getResponse = {
        data: {
          value: [
            {
              appDefinitions: [
                {
                  publishingState: PublishingState.submitted,
                  teamsAppId: "xx",
                  displayName: "xx",
                  lastModifiedDateTime: null,
                },
              ],
            },
          ],
        },
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(getResponse);

      try {
        await teamsDevPortalClient.publishTeamsAppUpdate(token, "", Buffer.from(""));
      } catch (error) {
        chai.assert.include(error.message, xCorrelationId);
      }
    });
    it("API Failure", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "error",
        message: "fake message",
      };
      sandbox.stub(fakeAxiosInstance, "post").throws(error);

      const getResponse = {
        data: {
          value: [
            {
              appDefinitions: [
                {
                  publishingState: PublishingState.submitted,
                  teamsAppId: "xx",
                  displayName: "xx",
                  lastModifiedDateTime: null,
                },
              ],
            },
          ],
        },
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(getResponse);

      try {
        await teamsDevPortalClient.publishTeamsAppUpdate(token, "", Buffer.from(""));
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });
    it("Bad Request", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const xCorrelationId = "fakeCorrelationId";
      const postResponse = {
        data: {
          errorMessage: "BadRequest",
        },
        message: "fake message",
        headers: {
          "x-correlation-id": xCorrelationId,
        },
      };

      sandbox.stub(fakeAxiosInstance, "post").resolves(postResponse);

      const getResponse = {
        data: {
          value: [
            {
              appDefinitions: [
                {
                  publishingState: PublishingState.submitted,
                  teamsAppId: "xx",
                  displayName: "xx",
                  lastModifiedDateTime: null,
                },
              ],
            },
          ],
        },
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(getResponse);

      try {
        await teamsDevPortalClient.publishTeamsAppUpdate(token, "", Buffer.from(""));
      } catch (error) {
        chai.assert.include(error.message, xCorrelationId);
        chai.assert.include(error.message, "BadRequest");
      }
    });
  });

  describe("wrapResponse", () => {
    it("should return an error with e.message if it exists", () => {
      const e = new Error("Error from e");
      const error = teamsDevPortalClient.wrapResponse(e, undefined);
      expect(error.message).to.equal("Error from e");
    });

    it("should return an error with e.message and response are missing", () => {
      const e = new Error("");
      const error = teamsDevPortalClient.wrapResponse(e, undefined);
      expect(error.message).to.equal("");
    });

    it("should return an error with response.data.error.message if e.message is missing and response.data.error.message exists", () => {
      const e = new Error("");
      const response = {
        data: { error: { message: "Error from response.data.error" }, errorMessage: "" },
      } as any;
      const error = teamsDevPortalClient.wrapResponse(e, response);
      expect(error.message).to.equal("Error from response.data.error");
      expect(error.response).to.equal(response);
      expect(error.request).to.equal(response.request);
    });

    it("should return an error with response.data.errorMessage if both e.message and response.data.error.message are missing", () => {
      const e = new Error("");
      const response = {
        data: { error: { message: "" }, errorMessage: "Error from response.data.errorMessage" },
      } as any;
      const error = teamsDevPortalClient.wrapResponse(e, response);
      expect(error.message).to.equal("Error from response.data.errorMessage");
      expect(error.response).to.equal(response);
      expect(error.request).to.equal(response.request);
    });

    it("should return an error with empty message if all messages are missing", () => {
      const e = new Error("");
      const response = { data: { error: { message: "" }, errorMessage: "" } } as any;
      const error = teamsDevPortalClient.wrapResponse(e, response);
      expect(error.message).to.equal("");
      expect(error.response).to.equal(response);
      expect(error.request).to.equal(response.request);
    });
  });

  describe("grantPermission", () => {
    it("no need to grant", async () => {
      sandbox.stub(teamsDevPortalClient, "getApp").resolves(appDef);
      sandbox.stub(teamsDevPortalClient, "checkUser").returns(true);
      try {
        await teamsDevPortalClient.grantPermission(token, "fake", {
          tenantId: uuid(),
          aadId: uuid(),
          displayName: "fake",
          userPrincipalName: "fake",
          isAdministrator: false,
        });
      } catch (e) {
        chai.assert.fail(Messages.ShouldNotReachHere);
      }
    });
    it("API Failure", async () => {
      sandbox.stub(teamsDevPortalClient, "getApp").resolves(appDef);
      sandbox.stub(teamsDevPortalClient, "checkUser").returns(false);
      sandbox.stub(RetryHandler, "Retry").rejects(new Error());
      const appUser: AppUser = {
        tenantId: uuid(),
        aadId: uuid(),
        displayName: "fake",
        userPrincipalName: "fake",
        isAdministrator: false,
      };
      try {
        await teamsDevPortalClient.grantPermission(token, appDef.teamsAppId!, appUser);
      } catch (e) {
        chai.assert.isTrue(e instanceof DeveloperPortalAPIFailedSystemError);
      }
    });
    it("response no data", async () => {
      sandbox.stub(teamsDevPortalClient, "getApp").resolves(appDef);
      sandbox.stub(teamsDevPortalClient, "checkUser").returns(false);
      sandbox.stub(RetryHandler, "Retry").resolves({
        data: undefined,
      });
      const appUser: AppUser = {
        tenantId: uuid(),
        aadId: uuid(),
        displayName: "fake",
        userPrincipalName: "fake",
        isAdministrator: false,
      };
      try {
        await teamsDevPortalClient.grantPermission(token, appDef.teamsAppId!, appUser);
      } catch (e) {
        chai.assert.isTrue(e.message.includes(ErrorMessages.GrantPermissionFailed));
      }
    });
    it("happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const newAppUser: AppUser = {
        tenantId: "new-tenant-id",
        aadId: "new-aad-id",
        displayName: "fake",
        userPrincipalName: "fake",
        isAdministrator: false,
      };
      const teamsAppId = appDef.teamsAppId!;
      const appDefWithUser: AppDefinition = {
        appName: "fake",
        teamsAppId: teamsAppId,
        userList: [
          {
            tenantId: "fake-tenant-id",
            aadId: "fake-aad-id",
            displayName: "fake",
            userPrincipalName: "fake",
            isAdministrator: false,
          },
        ],
      };
      const appDefWithUserAdded: AppDefinition = {
        appName: "fake",
        teamsAppId: teamsAppId,
        userList: [
          {
            tenantId: "fake-tenant-id",
            aadId: "fake-aad-id",
            displayName: "fake",
            userPrincipalName: "fake",
            isAdministrator: false,
          },
          newAppUser,
        ],
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves({
        data: appDefWithUser,
      });
      sandbox.stub(fakeAxiosInstance, "post").resolves({
        data: appDefWithUserAdded,
      });

      await teamsDevPortalClient.grantPermission(token, appDef.teamsAppId!, newAppUser);
    });
  });

  describe("removePermission", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const userToRemove: AppUser = {
        tenantId: "fakeTenantId",
        aadId: "testUserId",
        displayName: "Test User",
        userPrincipalName: "test@test.com",
        isAdministrator: false,
      };

      // Mock getApp response
      const getAppResponse = {
        data: {
          ...appDef,
          userList: [userToRemove],
        },
      };

      // Mock post response for removing permission
      const postResponse = {
        data: {
          ...appDef,
          userList: [],
        },
      };

      const getStub = sandbox.stub(fakeAxiosInstance, "get").resolves(getAppResponse);
      const postStub = sandbox.stub(fakeAxiosInstance, "post").resolves(postResponse);

      await teamsDevPortalClient.removePermission(token, appDef.teamsAppId!, userToRemove);

      chai.assert.isTrue(getStub.calledOnce);
      chai.assert.isTrue(postStub.calledOnce);
      chai.assert.equal(
        postStub.firstCall.args[0],
        `/api/appdefinitions/${appDef.teamsAppId!}/owner`
      );
    });

    it("User not exists", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const userToRemove: AppUser = {
        tenantId: "fakeTenantId",
        aadId: "nonExistUserId",
        displayName: "Non Exist User",
        userPrincipalName: "nonexist@test.com",
        isAdministrator: false,
      };

      // Mock getApp response with empty user list
      const getAppResponse = {
        data: {
          ...appDef,
          userList: [],
        },
      };

      sandbox.stub(fakeAxiosInstance, "get").resolves(getAppResponse);
      const postStub = sandbox.stub(fakeAxiosInstance, "post");

      // Should return directly without making post request
      await teamsDevPortalClient.removePermission(token, appDef.teamsAppId!, userToRemove);

      chai.assert.isTrue(postStub.notCalled);
    });

    it("API Failure", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const userToRemove: AppUser = {
        tenantId: "fakeTenantId",
        aadId: "testUserId",
        displayName: "Test User",
        userPrincipalName: "test@test.com",
        isAdministrator: false,
      };

      // Mock getApp response
      const getAppResponse = {
        data: {
          ...appDef,
          userList: [userToRemove],
        },
      };

      sandbox.stub(fakeAxiosInstance, "get").resolves(getAppResponse);

      const error = {
        name: "Error",
        message: "API call failed",
        response: {
          status: 500,
          data: { error: "Internal server error" },
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").throws(error);

      try {
        await teamsDevPortalClient.removePermission(token, appDef.teamsAppId!, userToRemove);
        chai.assert.fail("Should throw error");
      } catch (e) {
        chai.assert.equal(e.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });

    it("Empty response", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const userToRemove: AppUser = {
        tenantId: "fakeTenantId",
        aadId: "testUserId",
        displayName: "Test User",
        userPrincipalName: "test@test.com",
        isAdministrator: false,
      };

      // Mock getApp response
      const getAppResponse = {
        data: {
          ...appDef,
          userList: [userToRemove],
        },
      };

      sandbox.stub(fakeAxiosInstance, "get").resolves(getAppResponse);

      // Post response is empty
      const postResponse = {};
      sandbox.stub(fakeAxiosInstance, "post").resolves(postResponse);

      try {
        await teamsDevPortalClient.removePermission(token, appDef.teamsAppId!, userToRemove);
        chai.assert.fail("Should throw error");
      } catch (e) {
        chai.assert.equal(e.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });

    it("User not removed after API call", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const userToRemove: AppUser = {
        tenantId: "fakeTenantId",
        aadId: "testUserId",
        displayName: "Test User",
        userPrincipalName: "test@test.com",
        isAdministrator: false,
      };

      // Mock getApp response
      const getAppResponse = {
        data: {
          ...appDef,
          userList: [userToRemove],
        },
      };

      // Mock post response where user still exists
      const postResponse = {
        data: {
          ...appDef,
          userList: [userToRemove],
        },
      };

      sandbox.stub(fakeAxiosInstance, "get").resolves(getAppResponse);
      sandbox.stub(fakeAxiosInstance, "post").resolves(postResponse);

      try {
        await teamsDevPortalClient.removePermission(token, appDef.teamsAppId!, userToRemove);
        chai.assert.fail("Should throw error");
      } catch (e) {
        chai.assert.equal(e.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });
  });

  describe("getUserList", () => {
    it("happy path", async () => {
      sandbox.stub(teamsDevPortalClient, "getApp").resolves({
        userList: [
          {
            tenantId: "fake-tenant-id",
            aadId: "fake-aad-id",
            displayName: "fake",
            userPrincipalName: "fake",
            isAdministrator: false,
          },
        ],
      });
      const res = await teamsDevPortalClient.getUserList(token, appDef.teamsAppId!);
      chai.assert.equal(res!.length, 1);
    });
  });

  describe("checkPermission", () => {
    it("getUserList error", async () => {
      sandbox.stub(teamsDevPortalClient, "getUserList").rejects(new Error());
      const res = await teamsDevPortalClient.checkPermission(
        token,
        appDef.teamsAppId!,
        "fakeUesrId"
      );
      chai.assert.equal(res, Constants.PERMISSIONS.noPermission);
    });
    it("aadId not match", async () => {
      sandbox.stub(teamsDevPortalClient, "getUserList").resolves([
        {
          tenantId: "fake-tenant-id",
          aadId: "fake-aad-id",
          displayName: "fake",
          userPrincipalName: "fake",
          isAdministrator: false,
        },
      ]);
      const res = await teamsDevPortalClient.checkPermission(token, "any-id", "fakeUesrId");
      chai.assert.equal(res, Constants.PERMISSIONS.noPermission);
    });
    it("is admin", async () => {
      sandbox.stub(teamsDevPortalClient, "getUserList").resolves([
        {
          tenantId: "fake-tenant-id",
          aadId: "fake-aad-id",
          displayName: "fake",
          userPrincipalName: "fake",
          isAdministrator: true,
        },
      ]);
      const res = await teamsDevPortalClient.checkPermission(token, "any-id", "fake-aad-id");
      chai.assert.equal(res, Constants.PERMISSIONS.admin);
    });
    it("is operative", async () => {
      sandbox.stub(teamsDevPortalClient, "getUserList").resolves([
        {
          tenantId: "fake-tenant-id",
          aadId: "fake-aad-id",
          displayName: "fake",
          userPrincipalName: "fake",
          isAdministrator: false,
        },
      ]);
      const res = await teamsDevPortalClient.checkPermission(token, "any-id", "fake-aad-id");
      chai.assert.equal(res, Constants.PERMISSIONS.operative);
    });
  });

  describe("getApiKeyRegistration", () => {
    it("404 not found", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "404",
        message: "fake message",
      };
      sandbox.stub(fakeAxiosInstance, "get").throws(error);

      try {
        await teamsDevPortalClient.getApiKeyRegistrationById(token, "fakeId");
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });

    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: appApiRegistration,
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);

      const res = await teamsDevPortalClient.getApiKeyRegistrationById(token, "fakeId");
      chai.assert.equal(res, appApiRegistration);
    });
  });

  describe("createApiKeyRegistration", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: appApiRegistration,
      };
      sandbox.stub(fakeAxiosInstance, "post").resolves(response);

      const res = await teamsDevPortalClient.createApiKeyRegistration(token, appApiRegistration);
      chai.assert.equal(res, appApiRegistration);
    });

    it("Graph API failure", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        response: {
          staus: 400,
          data: {
            statusCode: 400,
            errorMessage:
              "Unsuccessful response received from Teams Graph Service. Error Message: System.Net.Http.HttpConnectionResponseContent",
          },
          headers: {
            "x-correlation-id": uuid(),
          },
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").throws(error);

      try {
        await teamsDevPortalClient.createApiKeyRegistration(token, appApiRegistration);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });
  });

  describe("updateApiKeyRegistration", () => {
    const appApiRegistration: ApiSecretRegistrationUpdate = {
      description: "fake description",
      applicableToApps: ApiSecretRegistrationAppType.AnyApp,
      targetUrlsShouldStartWith: ["https://www.example.com"],
    };
    it("404 not found", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "404",
        message: "fake message",
      };
      sandbox.stub(fakeAxiosInstance, "patch").throws(error);

      try {
        await teamsDevPortalClient.updateApiKeyRegistration(token, appApiRegistration, "fakeId");
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });

    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: appApiRegistration,
      };
      sandbox.stub(fakeAxiosInstance, "patch").resolves(response);

      const res = await teamsDevPortalClient.updateApiKeyRegistration(
        token,
        appApiRegistration,
        "fakeId"
      );
      chai.assert.equal(res, appApiRegistration);
    });
  });

  describe("createOauthRegistration", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: {
          configurationRegistrationId: {
            oAuthConfigId: "fakeId",
          },
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").resolves(response);

      const res = await teamsDevPortalClient.createOauthRegistration(token, fakeOauthRegistration);
      chai.assert.equal(res.configurationRegistrationId.oAuthConfigId, "fakeId");
    });

    it("Graph API failure", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        response: {
          staus: 400,
          data: {
            statusCode: 400,
            errorMessage:
              "Unsuccessful response received from Teams Graph Service. Error Message: System.Net.Http.HttpConnectionResponseContent",
          },
          headers: {
            "x-correlation-id": uuid(),
          },
        },
      };
      sandbox.stub(fakeAxiosInstance, "get").throws(error);

      try {
        await teamsDevPortalClient.createOauthRegistration(token, fakeOauthRegistration);
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });
  });

  describe("getOauthRegistration", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: fakeOauthRegistration,
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);

      const res = await teamsDevPortalClient.getOauthRegistrationById(token, "fakeId");
      chai.assert.equal(res, fakeOauthRegistration);
    });

    it("Graph API failure", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "404",
        message: "fake message",
      };
      sandbox.stub(fakeAxiosInstance, "get").throws(error);

      try {
        await teamsDevPortalClient.getOauthRegistrationById(token, "fakeId");
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });
  });

  describe("updateOauthRegistration", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: fakeOauthRegistration,
      };
      sandbox.stub(fakeAxiosInstance, "patch").resolves(response);

      const res = await teamsDevPortalClient.updateOauthRegistration(
        token,
        fakeOauthRegistration,
        "fakeId"
      );
      chai.assert.equal(res, fakeOauthRegistration);
    });

    it("Graph API failure", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "404",
        message: "fake message",
      };
      sandbox.stub(fakeAxiosInstance, "patch").throws(error);

      try {
        await teamsDevPortalClient.updateOauthRegistration(token, fakeOauthRegistration, "fakeId");
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });
  });

  describe("list Teams app", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: [appDef],
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      teamsDevPortalClient.setRegionEndpoint("https://dev.teams.microsoft.com/amer");
      const res = await teamsDevPortalClient.listApps(token);
      chai.assert.deepEqual(res, [appDef]);
    });
    it("Error - no region", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: [appDef],
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      teamsDevPortalClient.setRegionEndpoint("");
      try {
        await teamsDevPortalClient.listApps(token);
        chai.assert.fail("should throw error");
      } catch (e) {
        chai.assert.isTrue(e instanceof Error);
      }
    });
    it("Error - api failure", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      sandbox.stub(fakeAxiosInstance, "get").rejects(new Error());
      teamsDevPortalClient.setRegionEndpoint("https://dev.teams.microsoft.com/amer");
      try {
        await teamsDevPortalClient.listApps(token);
        chai.assert.fail("should throw error");
      } catch (e) {
        chai.assert.isTrue(e instanceof DeveloperPortalAPIFailedSystemError);
      }
    });
    it("Error - no data", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: undefined,
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      teamsDevPortalClient.setRegionEndpoint("https://dev.teams.microsoft.com/amer");
      try {
        await teamsDevPortalClient.listApps(token);
        chai.assert.fail("should throw error");
      } catch (e) {
        chai.assert.isTrue(
          e.message.includes(
            "Unable to make API call to Developer Portal: API failed, cannot get the app definitions, API name: list-app, X-Correlation-ID: undefined. This may be due to a temporary service error. Try again after a few minutes."
          )
        );
      }
    });
  });

  describe("delete Teams app", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      const response = {
        data: true,
      };
      sandbox.stub(fakeAxiosInstance, "delete").resolves(response);
      teamsDevPortalClient.setRegionEndpoint("https://dev.teams.microsoft.com/amer");
      const res = await teamsDevPortalClient.deleteApp(token, "testid");
      chai.assert.isTrue(res);
    });
    it("Error - no region", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: [appDef],
      };
      sandbox.stub(fakeAxiosInstance, "delete").resolves(response);
      teamsDevPortalClient.setRegionEndpoint("");
      try {
        await teamsDevPortalClient.deleteApp(token, "testid");
        chai.assert.fail("should throw error");
      } catch (e) {
        chai.assert.isTrue(e instanceof Error);
      }
    });
    it("Error - api failure", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      sandbox.stub(fakeAxiosInstance, "delete").rejects(new Error());
      teamsDevPortalClient.setRegionEndpoint("https://dev.teams.microsoft.com/amer");
      try {
        await teamsDevPortalClient.deleteApp(token, "testid");
        chai.assert.fail("should throw error");
      } catch (e) {
        chai.assert.isTrue(e instanceof DeveloperPortalAPIFailedSystemError);
      }
    });
    it("Error - no data", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: undefined,
      };
      sandbox.stub(fakeAxiosInstance, "delete").resolves(response);
      teamsDevPortalClient.setRegionEndpoint("https://dev.teams.microsoft.com/amer");
      try {
        await teamsDevPortalClient.deleteApp(token, "testid");
        chai.assert.fail("should throw error");
      } catch (e) {
        chai.assert.isTrue(e.message.includes("cannot delete the app: " + "testid"));
      }
    });
  });

  describe("Submit async app validation request", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      const response = {
        data: {
          appValidationId: uuid(),
          status: AsyncAppValidationStatus.Created,
        },
      };
      sandbox.stub(fakeAxiosInstance, "post").resolves(response);
      const res = await teamsDevPortalClient.submitAppValidationRequest(token, "fakeId");
      chai.assert.equal(res.appValidationId, response.data.appValidationId);
    });
  });

  describe("Get async app validation request list", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      const response = {
        data: {
          continuationToken: "",
          appValidations: [],
        },
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      const res = await teamsDevPortalClient.getAppValidationRequestList(token, "fakeId");
      chai.assert.equal(res.appValidations!.length, 0);
    });

    it("404 not found", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "404",
        message: "fake message",
      };
      sandbox.stub(fakeAxiosInstance, "post").throws(error);

      try {
        await teamsDevPortalClient.submitAppValidationRequest(token, "fakeId");
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });
  });

  describe("Get async app validation result details", () => {
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      const response = {
        data: {
          appValidationId: "fakeId",
          appId: "fakeAppId",
          status: AsyncAppValidationStatus.Completed,
          appVersion: "1.0.0",
          manifestVersion: "1.16",
          createdAt: Date(),
          updatedAt: Date(),
          validationResults: {
            successes: [],
            warnings: [],
            failures: [],
            skipped: [],
          },
        },
      };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      const res = await teamsDevPortalClient.getAppValidationById(token, "fakeId");
      chai.assert.equal(res.appValidationId, "fakeId");
    });

    it("404 not found", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "404",
        message: "fake message",
      };
      sandbox.stub(fakeAxiosInstance, "get").throws(error);

      try {
        await teamsDevPortalClient.getAppValidationRequestList(token, "fakeId");
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });

    it("404 not found", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const error = {
        name: "404",
        message: "fake message",
      };
      sandbox.stub(fakeAxiosInstance, "get").throws(error);

      try {
        await teamsDevPortalClient.getAppValidationById(token, "fakeId");
      } catch (error) {
        chai.assert.equal(error.name, DeveloperPortalAPIFailedSystemError.name);
      }
    });
  });

  describe("getBotRegistration", () => {
    it("Should return a valid bot registration", async () => {
      // Arrange
      sandbox.stub(RetryHandler, "Retry").resolves({
        status: 200,
        data: sampleBot,
      });
      // Act
      const res = await teamsDevPortalClient.getBotRegistration("anything", "anything");

      // Assert
      chai.assert.isTrue(res !== undefined);
      chai.assert.isTrue(res?.botId === sampleBot.botId);
    });

    it("Should return a undefined when 404 was throwed out", async () => {
      // Arrange
      const mockAxiosInstance = axios.create();
      sandbox.stub(mockAxiosInstance, "get").rejects({
        response: {
          status: 404,
        },
      });
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);

      // Act
      const res = await teamsDevPortalClient.getBotRegistration("anything", "anything");

      // Assert
      chai.assert.isUndefined(res);
    });

    it("Should throw NotAllowedToAcquireToken error when 401 was throwed out", async () => {
      // Arrange
      const mockAxiosInstance = axios.create();
      sandbox.stub(mockAxiosInstance, "get").rejects({
        response: {
          status: 401,
        },
      });
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);

      // Act & Assert
      try {
        await teamsDevPortalClient.getBotRegistration("anything", "anything");
        chai.assert.fail(Messages.ShouldNotReachHere);
      } catch (e) {
        chai.assert.isTrue(e.name === ErrorNames.ACQUIRE_BOT_FRAMEWORK_TOKEN_ERROR);
      }
    });

    it("Should throw DeveloperPortalAPIFailed error when other exceptions (500) were throwed out", async () => {
      // Arrange
      sandbox.stub(RetryHandler, "Retry").rejects({
        response: {
          headers: {
            "x-correlation-id": "anything",
          },
          status: 500,
        },
      });

      // Act & Assert
      try {
        await teamsDevPortalClient.getBotRegistration("anything", "anything");
        chai.assert.fail(Messages.ShouldNotReachHere);
      } catch (e) {
        chai.assert.isTrue(e.name === DeveloperPortalAPIFailedSystemError.name);
      }
    });
  });

  describe("createBotRegistration", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("Bot registration should be created successfully", async () => {
      // Arrange
      sandbox.stub(teamsDevPortalClient, "getBotRegistration").resolves(undefined);
      const mockAxiosInstance = axios.create();
      sandbox.stub(mockAxiosInstance, "post").resolves({
        status: 200,
        data: sampleBot,
      });
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);

      // Act & Assert
      try {
        await teamsDevPortalClient.createBotRegistration("anything", sampleBot);
      } catch (e) {
        chai.assert.fail(Messages.ShouldNotReachHere);
      }
    });

    it("Bot registration creation should be skipped (existing bot case).", async () => {
      // Arrange
      sandbox.stub(teamsDevPortalClient, "getBotRegistration").resolves(sampleBot);

      // Act & Assert
      try {
        await teamsDevPortalClient.createBotRegistration("anything", sampleBot);
      } catch (e) {
        chai.assert.fail(Messages.ShouldNotReachHere);
      }
    });

    it("BotFrameworkNotAllowedToAcquireToken error should be throwed out (401)", async () => {
      // Arrange
      sandbox.stub(teamsDevPortalClient, "getBotRegistration").resolves(undefined);
      const mockAxiosInstance = axios.create();
      sandbox.stub(mockAxiosInstance, "post").rejects({
        response: {
          status: 401,
        },
      });
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);

      // Act & Assert
      try {
        await teamsDevPortalClient.createBotRegistration("anything", sampleBot);
        chai.assert.fail(Messages.ShouldNotReachHere);
      } catch (e) {
        chai.assert.isTrue(e.name === ErrorNames.ACQUIRE_BOT_FRAMEWORK_TOKEN_ERROR);
      }
    });

    it("BotFrameworkForbiddenResult error should be throwed out (403)", async () => {
      // Arrange
      sandbox.stub(teamsDevPortalClient, "getBotRegistration").resolves(undefined);
      const mockAxiosInstance = axios.create();
      sandbox.stub(mockAxiosInstance, "post").rejects({
        response: {
          status: 403,
        },
      });
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);

      // Act & Assert
      try {
        await teamsDevPortalClient.createBotRegistration("anything", sampleBot);
        chai.assert.fail(Messages.ShouldNotReachHere);
      } catch (e) {
        chai.assert.isTrue(e.name === ErrorNames.FORBIDDEN_RESULT_BOT_FRAMEWORK_ERROR);
      }
    });

    it("BotFrameworkConflictResult error should be throwed out (429)", async () => {
      // Arrange
      sandbox.stub(teamsDevPortalClient, "getBotRegistration").resolves(undefined);
      const mockAxiosInstance = axios.create();
      sandbox.stub(mockAxiosInstance, "post").rejects({
        response: {
          status: 429,
        },
      });
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);

      // Act & Assert
      try {
        await teamsDevPortalClient.createBotRegistration("anything", sampleBot);
        chai.assert.fail(Messages.ShouldNotReachHere);
      } catch (e) {
        chai.assert.isTrue(e.name === ErrorNames.CONFLICT_RESULT_BOT_FRAMEWORK_ERROR);
      }
    });

    it("DeveloperPortalAPIFailed error should be throwed out (500)", async () => {
      // Arrange
      sandbox.stub(teamsDevPortalClient, "getBotRegistration").resolves(undefined);
      sandbox.stub(RetryHandler, "Retry").rejects({
        response: {
          headers: {
            "x-correlation-id": "anything",
          },
          status: 500,
        },
      });

      // Act & Assert
      try {
        await teamsDevPortalClient.createBotRegistration("anything", sampleBot);
        chai.assert.fail(Messages.ShouldNotReachHere);
      } catch (e) {
        chai.assert.isTrue(e.name === DeveloperPortalAPIFailedSystemError.name);
      }
    });
  });

  describe("updateBotRegistration", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("Bot registration should be updated successfully", async () => {
      // Arrange
      const mockAxiosInstance = axios.create();
      sandbox.stub(mockAxiosInstance, "post").resolves({
        status: 200,
        data: sampleBot,
      });
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);

      // Act & Assert
      try {
        await teamsDevPortalClient.updateBotRegistration("anything", sampleBot);
      } catch (e) {
        chai.assert.fail(Messages.ShouldNotReachHere);
      }
    });

    it("BotFrameworkNotAllowedToAcquireToken error should be throwed out (401)", async () => {
      // Arrange
      const mockAxiosInstance = axios.create();
      sandbox.stub(mockAxiosInstance, "post").rejects({
        response: {
          status: 401,
        },
      });
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);

      // Act & Assert
      try {
        await teamsDevPortalClient.updateBotRegistration("anything", sampleBot);
        chai.assert.fail(Messages.ShouldNotReachHere);
      } catch (e) {
        chai.assert.isTrue(e.name === ErrorNames.ACQUIRE_BOT_FRAMEWORK_TOKEN_ERROR);
      }
    });

    it("BotFrameworkForbiddenResult error should be throwed out (403)", async () => {
      // Arrange
      const mockAxiosInstance = axios.create();
      sandbox.stub(mockAxiosInstance, "post").rejects({
        response: {
          status: 403,
        },
      });
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);

      // Act & Assert
      try {
        await teamsDevPortalClient.updateBotRegistration("anything", sampleBot);
        chai.assert.fail(Messages.ShouldNotReachHere);
      } catch (e) {
        chai.assert.isTrue(e.name === ErrorNames.FORBIDDEN_RESULT_BOT_FRAMEWORK_ERROR);
      }
    });

    it("BotFrameworkConflictResult error should be throwed out (429)", async () => {
      // Arrange
      const mockAxiosInstance = axios.create();
      sandbox.stub(mockAxiosInstance, "post").rejects({
        response: {
          status: 429,
        },
      });
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);

      // Act & Assert
      try {
        await teamsDevPortalClient.updateBotRegistration("anything", sampleBot);
        chai.assert.fail(Messages.ShouldNotReachHere);
      } catch (e) {
        chai.assert.isTrue(e.name === ErrorNames.CONFLICT_RESULT_BOT_FRAMEWORK_ERROR);
      }
    });

    it("DeveloperPortalAPIFailed error should be throwed out (500)", async () => {
      // Arrange
      sandbox.stub(RetryHandler, "Retry").rejects({
        response: {
          headers: {
            "x-correlation-id": "anything",
          },
          status: 500,
        },
      });

      // Act & Assert
      try {
        await teamsDevPortalClient.updateBotRegistration("anything", sampleBot);
        chai.assert.fail(Messages.ShouldNotReachHere);
      } catch (e) {
        chai.assert.isTrue(e.name === DeveloperPortalAPIFailedSystemError.name);
      }
    });
  });

  describe("updateMessageEndpoint", () => {
    afterEach(() => {
      sandbox.restore();
    });

    it("Message endpoint should be updated successfully", async () => {
      // Arrange
      sandbox.stub(teamsDevPortalClient, "getBotRegistration").resolves(sampleBot);
      sandbox.stub(teamsDevPortalClient, "updateBotRegistration").resolves();
      // Act & Assert
      try {
        await teamsDevPortalClient.updateMessageEndpoint("anything", "anything", "anything");
      } catch (e) {
        chai.assert.fail(Messages.ShouldNotReachHere);
      }
    });

    it("BotRegistrationNotFound error should be throwed out", async () => {
      // Arrange
      sandbox.stub(teamsDevPortalClient, "getBotRegistration").resolves(undefined);
      // Act & Assert
      try {
        await teamsDevPortalClient.updateMessageEndpoint("anything", "anything", "anything");
        chai.assert.fail(Messages.ShouldNotReachHere);
      } catch (e) {
        chai.assert.isTrue(e.name === ErrorNames.BOT_REGISTRATION_NOTFOUND_ERROR);
      }
    });
  });

  describe("listBots", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("happy", async () => {
      // Arrange
      const mockAxiosInstance = axios.create();
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);
      sandbox.stub(mockAxiosInstance, "get").resolves({
        status: 200,
        data: [sampleBot],
      });
      // Act & Assert
      try {
        const res = await teamsDevPortalClient.listBots("anything");
        chai.assert.deepEqual(res, [sampleBot]);
      } catch (e) {
        chai.assert.fail(Messages.ShouldNotReachHere);
      }
    });
    it("invalid response", async () => {
      // Arrange
      const mockAxiosInstance = axios.create();
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);
      sandbox.stub(mockAxiosInstance, "get").resolves({
        status: 200,
      });
      // Act & Assert
      try {
        await teamsDevPortalClient.listBots("anything");
        chai.assert.fail(Messages.ShouldNotReachHere);
      } catch (e) {}
    });
    it("api failure", async () => {
      // Arrange
      const mockAxiosInstance = axios.create();
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);
      sandbox.stub(mockAxiosInstance, "get").resolves({ response: { status: 404 } });
      // Act & Assert
      try {
        await teamsDevPortalClient.listBots("anything");
        chai.assert.fail(Messages.ShouldNotReachHere);
      } catch (e) {
        chai.assert.isTrue(e instanceof DeveloperPortalAPIFailedSystemError);
      }
    });
  });
  describe("deleteBot", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("happy", async () => {
      // Arrange
      const mockAxiosInstance = axios.create();
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);
      sandbox.stub(mockAxiosInstance, "delete").resolves({
        status: 200,
      });
      // Act & Assert
      try {
        await teamsDevPortalClient.deleteBot("anything", "anything");
      } catch (e) {
        chai.assert.fail(Messages.ShouldNotReachHere);
      }
    });
    it("throw error", async () => {
      // Arrange
      const mockAxiosInstance = axios.create();
      sandbox.stub(teamsDevPortalClient, "createRequesterWithToken").returns(mockAxiosInstance);
      sandbox.stub(mockAxiosInstance, "delete").rejects({ response: { status: 404 } });
      // Act & Assert
      try {
        await teamsDevPortalClient.deleteBot("anything", "anything");
        chai.assert.fail(Messages.ShouldNotReachHere);
      } catch (e) {
        chai.assert.isTrue(e instanceof Error);
      }
    });
  });
  describe("getSideloadingStatus()", () => {
    let mockGet: () => AxiosResponse;
    let events: number;
    let errors: number;
    beforeEach(() => {
      const mockInstance = axios.create();
      sandbox.stub(mockInstance, "get").callsFake(async () => mockGet());
      sandbox.stub(axios, "create").returns(mockInstance);

      events = 0;
      sandbox.stub(telemetry, "sendTelemetryEvent").callsFake(() => {
        ++events;
      });

      errors = 0;
      sandbox.stub(telemetry, "sendTelemetryErrorEvent").callsFake(() => {
        ++errors;
      });
    });
    it("sideloading enabled", async () => {
      mockGet = () => {
        return {
          status: 200,
          data: {
            value: {
              isSideloadingAllowed: true,
            },
          },
        } as AxiosResponse;
      };

      const result = await teamsDevPortalClient.getSideloadingStatus("fake-token");

      chai.assert.isDefined(result);
      chai.assert.isTrue(result);
      chai.assert.equal(events, 1);
      chai.assert.equal(errors, 0);
    });
    it("status > 400", async () => {
      mockGet = () => {
        return {
          status: 404,
        } as AxiosResponse;
      };
      const result = await teamsDevPortalClient.getSideloadingStatus("fake-token");
      chai.assert.isUndefined(result);
    });
    it("sideloading not enabled", async () => {
      mockGet = () => {
        return {
          status: 200,
          data: {
            value: {
              isSideloadingAllowed: false,
            },
          },
        } as AxiosResponse;
      };

      const result = await teamsDevPortalClient.getSideloadingStatus("fake-token");

      chai.assert.isDefined(result);
      chai.assert.isFalse(result);
      chai.assert.equal(events, 1);
      chai.assert.equal(errors, 0);
    });

    it("sideloading unknown", async () => {
      mockGet = () => {
        return {
          status: 200,
          data: {
            value: {
              foo: "bar",
            },
          },
        } as AxiosResponse;
      };

      const result = await teamsDevPortalClient.getSideloadingStatus("fake-token");

      chai.assert.isUndefined(result);
      chai.assert.equal(events, 0);
      chai.assert.equal(errors, 1);
    });

    it("error and retry", async () => {
      sandbox.stub(RetryHandler, "Retry").rejects(new Error());
      const res = await teamsDevPortalClient.getSideloadingStatus("fake-token");
      chai.assert.isUndefined(res);
    });
  });
  describe("getBotId", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it("happy", async () => {
      sandbox.stub(teamsDevPortalClient, "getApp").resolves({
        bots: [
          {
            botId: "mocked-bot-id",
            needsChannelSelector: false,
            isNotificationOnly: false,
            supportsFiles: false,
            supportsCalling: false,
            supportsVideo: false,
            scopes: [],
            teamCommands: [],
            personalCommands: [],
            groupChatCommands: [],
          },
        ],
      });
      try {
        const res = await teamsDevPortalClient.getBotId("token", "anything");
        chai.assert.equal(res, "mocked-bot-id");
      } catch (e) {
        chai.assert.fail(Messages.ShouldNotReachHere);
      }
    });
    it("empty bots", async () => {
      sandbox.stub(teamsDevPortalClient, "getApp").resolves({
        bots: [],
      });
      try {
        const res = await teamsDevPortalClient.getBotId("token", "anything");
        chai.assert.isUndefined(res);
      } catch (e) {
        chai.assert.fail(Messages.ShouldNotReachHere);
      }
    });
    it("no bots", async () => {
      sandbox.stub(teamsDevPortalClient, "getApp").resolves({});
      try {
        const res = await teamsDevPortalClient.getBotId("token", "anything");
        chai.assert.isUndefined(res);
      } catch (e) {
        chai.assert.fail(Messages.ShouldNotReachHere);
      }
    });
  });
});
