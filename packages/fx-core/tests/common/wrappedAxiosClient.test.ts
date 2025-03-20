// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios, { AxiosInstance } from "axios";
import * as chai from "chai";
import "mocha";
import * as sinon from "sinon";
import { v4 as uuid } from "uuid";
import { getAppStudioEndpoint } from "../../src/common/constants";
import { setTools } from "../../src/common/globalVars";
import { WrappedAxiosClient } from "../../src/common/wrappedAxiosClient";
import { APP_STUDIO_API_NAMES } from "../../src/component/driver/teamsApp/constants";
import { MockTools } from "../core/utils";
import { MOS3ApiDefinitions } from "../../src/component/m365/serviceConstant";

describe("Wrapped Axios Client Test", () => {
  const mockTools = new MockTools();
  beforeEach(() => {
    setTools(mockTools);
  });

  afterEach(() => {
    sinon.restore();
  });

  it("create", async () => {
    const testAxiosInstance = {
      interceptors: {
        request: {
          use: sinon.stub(),
        },
        response: {
          use: sinon.stub(),
        },
      },
    } as any as AxiosInstance;
    sinon.stub(axios, "create").returns(testAxiosInstance);
    WrappedAxiosClient.create();
  });

  it("No telemetry reporter", async () => {
    setTools({} as any);

    const mockedRequest = {
      method: "POST",
      baseURL: getAppStudioEndpoint(),
      url: "/amer/api/appdefinitions/v2/import",
      params: {
        overwriteIfAppAlreadyExists: true,
      },
      status: 200,
      data: {},
    } as any;
    WrappedAxiosClient.onRequest(mockedRequest);

    const mockedResponse = {
      request: {
        method: "GET",
        host: getAppStudioEndpoint(),
        path: "/api/appdefinitions/manifest",
      },
      config: {
        params: {},
      },
      status: 200,
      data: {},
    } as any;
    WrappedAxiosClient.onResponse(mockedResponse);

    const mockedError = {
      request: {
        method: "GET",
        host: getAppStudioEndpoint(),
        path: "/api/appdefinitions/fakeId",
      },
      config: {
        data: "Invalid JSON",
      },
      response: {
        status: 404,
        headers: {
          "x-ms-correlation-id": uuid(),
        },
      },
    } as any;
    WrappedAxiosClient.onRejected(mockedError);
  });

  it("TOOLS not initialized", async () => {
    setTools(undefined as any);

    const mockedRequest = {
      method: "POST",
      baseURL: getAppStudioEndpoint(),
      url: "/amer/api/appdefinitions/v2/import",
      params: {
        overwriteIfAppAlreadyExists: true,
      },
      status: 200,
      data: {},
    } as any;
    WrappedAxiosClient.onRequest(mockedRequest);

    const mockedResponse = {
      request: {
        method: "GET",
        host: getAppStudioEndpoint(),
        path: "/api/appdefinitions/manifest",
      },
      config: {
        params: {},
      },
      status: 200,
      data: {},
    } as any;
    WrappedAxiosClient.onResponse(mockedResponse);

    const mockedError = {
      request: {
        method: "GET",
        host: getAppStudioEndpoint(),
        path: "/api/appdefinitions/fakeId",
      },
      config: {},
      response: {
        status: 404,
        headers: {
          "x-ms-correlation-id": uuid(),
        },
      },
    } as any;
    WrappedAxiosClient.onRejected(mockedError);
  });

  it("TDP API start telemetry", async () => {
    const mockedRequest = {
      method: "POST",
      baseURL: getAppStudioEndpoint(),
      url: "/amer/api/appdefinitions/v2/import",
      params: {
        overwriteIfAppAlreadyExists: true,
      },
      status: 200,
      data: {},
    } as any;
    const telemetryChecker = sinon.spy(mockTools.telemetryReporter, "sendTelemetryEvent");

    WrappedAxiosClient.onRequest(mockedRequest);
    chai.expect(telemetryChecker.calledOnce).to.be.true;
  });

  it("Dependency API start telemetry", async () => {
    const mockedRequest = {
      method: "POST",
      baseURL: "https://example.com",
      url: "",
      params: {},
      status: 200,
      data: {},
    } as any;
    const telemetryChecker = sinon.spy(mockTools.telemetryReporter, "sendTelemetryEvent");

    WrappedAxiosClient.onRequest(mockedRequest);
    chai.expect(telemetryChecker.calledOnce).to.be.true;
  });

  it("TDP API success response", async () => {
    const mockedResponse = {
      request: {
        method: "GET",
        host: getAppStudioEndpoint(),
        path: "/api/appdefinitions/manifest",
      },
      config: {
        params: {},
      },
      status: 200,
      data: {},
    } as any;
    const telemetryChecker = sinon.spy(mockTools.telemetryReporter, "sendTelemetryEvent");

    WrappedAxiosClient.onResponse(mockedResponse);
    chai.expect(telemetryChecker.calledOnce).to.be.true;
  });

  it("Dependency API success response", async () => {
    const mockedResponse = {
      request: {
        method: "GET",
        host: "https://example.com",
        path: "",
      },
      config: {
        params: {},
      },
      status: 200,
      data: {},
    } as any;
    const telemetryChecker = sinon.spy(mockTools.telemetryReporter, "sendTelemetryEvent");

    WrappedAxiosClient.onResponse(mockedResponse);
    chai.expect(telemetryChecker.calledOnce).to.be.true;
  });

  it("TDP API error response", async () => {
    const mockedError = {
      request: {
        method: "GET",
        host: getAppStudioEndpoint(),
        path: "/api/appdefinitions/fakeId",
      },
      config: {},
      response: {
        status: 404,
        headers: {
          "x-ms-correlation-id": uuid(),
        },
      },
    } as any;
    const telemetryChecker = sinon.spy(mockTools.telemetryReporter, "sendTelemetryErrorEvent");

    WrappedAxiosClient.onRejected(mockedError);
    chai.expect(telemetryChecker.calledOnce).to.be.true;
  });

  it("Dependency API error response", async () => {
    const mockedError = {
      request: {
        method: "GET",
        host: "https://example.com",
        path: "",
      },
      config: {
        data: '{"botId":"fakeId"}',
      },
      response: {
        status: 400,
      },
    } as any;
    const telemetryChecker = sinon.spy(mockTools.telemetryReporter, "sendTelemetryErrorEvent");

    WrappedAxiosClient.onRejected(mockedError);
    chai.expect(telemetryChecker.calledOnce).to.be.true;
  });

  it("MOS API error response", async () => {
    const mockedError = {
      request: {
        method: "POST",
        host: "https://titles.prod.mos.microsoft.com",
        path: "/dev/v1/users/packages",
      },
      config: {},
      response: {
        status: 400,
        data: {
          code: "BadRequest",
          message: "Invalid request",
        },
      },
    } as any;
    const telemetryChecker = sinon.spy(mockTools.telemetryReporter, "sendTelemetryErrorEvent");
    WrappedAxiosClient.onRejected(mockedError);
    chai.expect(telemetryChecker.calledOnce).to.be.true;
  });

  it("Create bot API start telemetry", async () => {
    const mockedRequest = {
      method: "POST",
      baseURL: getAppStudioEndpoint(),
      url: "/api/botframework",
      params: {},
      status: 200,
      data: {
        botId: "fakeId",
      },
    } as any;
    const telemetryChecker = sinon.spy(mockTools.telemetryReporter, "sendTelemetryEvent");

    WrappedAxiosClient.onRequest(mockedRequest);
    chai.expect(telemetryChecker.calledOnce).to.be.true;
  });

  it("Update bot API start telemetry", async () => {
    const mockedRequest = {
      method: "POST",
      baseURL: getAppStudioEndpoint(),
      url: `/api/botframework/${uuid()}`,
      params: {},
      status: 200,
      data: {},
    } as any;
    const telemetryChecker = sinon.spy(mockTools.telemetryReporter, "sendTelemetryEvent");

    WrappedAxiosClient.onRequest(mockedRequest);
    chai.expect(telemetryChecker.calledOnce).to.be.true;
  });

  it("Convert API name", async () => {
    const fakeId = uuid();

    let apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + "/api/appdefinitions/partnerCenterAppPackageValidation",
      "POST"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.VALIDATE_APP_PACKAGE);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/appdefinitions/${fakeId}/manifest`,
      "GET"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.GET_APP_PACKAGE);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/appdefinitions/${fakeId}/owner`,
      "POST"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.UPDATE_OWNER);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/appdefinitions/${fakeId}`,
      "GET"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.GET_APP);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/appdefinitions/${fakeId}`,
      "DELETE"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.DELETE_APP);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/publishing/${fakeId}`,
      "GET"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.GET_PUBLISHED_APP);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/publishing`,
      "POST"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.PUBLISH_APP);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/publishing/${fakeId}/appdefinitions`,
      "POST"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.UPDATE_PUBLISHED_APP);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/usersettings/mtUserAppPolicy`,
      "GET"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.CHECK_SIDELOADING_STATUS);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/v1.0/apiSecretRegistrations/${fakeId}`,
      "GET"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.GET_API_KEY);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/v1.0/apiSecretRegistrations/${fakeId}`,
      "PATCH"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.UPDATE_API_KEY);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/v1.0/apiSecretRegistrations`,
      "POST"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.CREATE_API_KEY);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/botframework/${fakeId}`,
      "GET"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.GET_BOT);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/botframework/${fakeId}`,
      "POST"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.UPDATE_BOT);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/botframework/${fakeId}`,
      "DELETE"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.DELETE_BOT);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/botframework`,
      "GET"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.LIST_BOT);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/botframework`,
      "POST"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.CREATE_BOT);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/v1.0/appvalidations/appdefinition/validate`,
      "POST"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.SUBMIT_APP_VALIDATION);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() +
        `/api/v1.0/appvalidations/appdefinitions/efe81961-44bc-49ae-99f8-1476caef994c`,
      "GET"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.GET_APP_VALIDATION_REQUESTS);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/v1.0/appvalidations/2512d616-8aac-461f-8af0-23e9b09ec650`,
      "GET"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.GET_APP_VALIDATION_RESULT);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/v1.0/oAuthConfigurations`,
      "POST"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.CREATE_OAUTH);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/v1.0/oAuthConfigurations/${fakeId}`,
      "GET"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.GET_OAUTH);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/v1.0/oAuthConfigurations/${fakeId}`,
      "PATCH"
    );
    chai.assert.equal(apiName, APP_STUDIO_API_NAMES.UPDATE_OAUTH);

    apiName = WrappedAxiosClient.convertUrlToApiName(
      getAppStudioEndpoint() + `/api/v1.0/oAuthConfigurations/${fakeId}`,
      ""
    );
    chai.assert.notEqual(apiName, APP_STUDIO_API_NAMES.UPDATE_OAUTH);

    apiName = WrappedAxiosClient.convertUrlToApiName(getAppStudioEndpoint() + `unknown`, "GET");
    chai.assert.equal(apiName, (getAppStudioEndpoint() + `unknown`).replace(/\//g, `-`));

    apiName = WrappedAxiosClient.convertUrlToApiName(
      "https://authsvc.teams.microsoft.com/v1.0/users/region",
      "POST"
    );
    chai.assert.equal(apiName, "get-region");

    apiName = WrappedAxiosClient.convertUrlToApiName("https://example.com", "GET");
    chai.assert.equal(apiName, "https:--example.com");
  });

  it("Convert API Definition for MOS API", async () => {
    const fakeId = uuid();

    let modApiDef = WrappedAxiosClient.convertMethodUrlToApiDefForMOS(
      "GET",
      "/config/v1/environment"
    );
    chai.assert.deepEqual(modApiDef, MOS3ApiDefinitions.GetConfigEnv);

    modApiDef = WrappedAxiosClient.convertMethodUrlToApiDefForMOS(
      "POST",
      "/dev/v1/users/packages/addins"
    );
    chai.assert.deepEqual(modApiDef, MOS3ApiDefinitions.PostPackageAddin);

    modApiDef = WrappedAxiosClient.convertMethodUrlToApiDefForMOS(
      "GET",
      `/dev/v1/users/packages/status/${fakeId}`
    );
    chai.assert.deepEqual(modApiDef, MOS3ApiDefinitions.GetDevStatus);

    modApiDef = WrappedAxiosClient.convertMethodUrlToApiDefForMOS(
      "POST",
      "/builder/v1/users/packages?scope=Personal"
    );
    chai.assert.deepEqual(modApiDef, MOS3ApiDefinitions.PostBuilderPackage);

    modApiDef = WrappedAxiosClient.convertMethodUrlToApiDefForMOS(
      "GET",
      `/builder/v1/users/packages/status/${fakeId}`
    );
    chai.assert.deepEqual(modApiDef, MOS3ApiDefinitions.GetBuilderStatus);

    modApiDef = WrappedAxiosClient.convertMethodUrlToApiDefForMOS("POST", "/dev/v1/users/packages");
    chai.assert.deepEqual(modApiDef, MOS3ApiDefinitions.PostDevPackage);

    modApiDef = WrappedAxiosClient.convertMethodUrlToApiDefForMOS(
      "POST",
      "/dev/v1/users/packages/acquisitions"
    );
    chai.assert.deepEqual(modApiDef, MOS3ApiDefinitions.PostDevPackageAcquisitions);

    modApiDef = WrappedAxiosClient.convertMethodUrlToApiDefForMOS(
      "GET",
      `/marketplace/v1/users/titles/${fakeId}/sharingInfo`
    );
    chai.assert.deepEqual(modApiDef, MOS3ApiDefinitions.GetShareInfo);

    modApiDef = WrappedAxiosClient.convertMethodUrlToApiDefForMOS(
      "GET",
      "/catalog/v1/users/titles/launchInfo"
    );
    chai.assert.deepEqual(modApiDef, MOS3ApiDefinitions.GetCatalogLaunchInfo);

    modApiDef = WrappedAxiosClient.convertMethodUrlToApiDefForMOS(
      "DELETE",
      `/catalog/v1/users/acquisitions/${fakeId}`
    );
    chai.assert.deepEqual(modApiDef, MOS3ApiDefinitions.DeleteCatalogAcquisitions);

    modApiDef = WrappedAxiosClient.convertMethodUrlToApiDefForMOS(
      "GET",
      `/catalog/v1/users/titles/${fakeId}/launchInfo`
    );
    chai.assert.deepEqual(modApiDef, MOS3ApiDefinitions.GetLaunchInfoByTitle);

    modApiDef = WrappedAxiosClient.convertMethodUrlToApiDefForMOS(
      "GET",
      "/catalog/v1/users/uitypes"
    );
    chai.assert.deepEqual(modApiDef, MOS3ApiDefinitions.GetCatalogUITypes);

    modApiDef = WrappedAxiosClient.convertMethodUrlToApiDefForMOS("GET", "/abcdef/v1/users/xxxxx");
    chai.assert.isUndefined(modApiDef);
  });
});
