// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import axios, { AxiosInstance, AxiosRequestConfig, AxiosResponse } from "axios";
import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import fs from "fs-extra";
import "mocha";
import sinon from "sinon";
import { NotExtendedToM365Error } from "../../../src/component/m365/errors";
import { AppScope, PackageService } from "../../../src/component/m365/packageService";
import { setTools } from "../../../src/common/globalVars";
import { UnhandledError } from "../../../src/error/common";
import { MockLogProvider } from "../../core/utils";

chai.use(chaiAsPromised);

describe("Package Service", () => {
  const sandbox = sinon.createSandbox();
  const logger = new MockLogProvider();
  let axiosDeleteResponses: Record<string, unknown> = {};
  let axiosGetResponses: Record<string, unknown> = {};
  let axiosPostResponses: Record<string, unknown> = {};
  const testAxiosInstance = {
    defaults: {
      headers: {
        common: {},
      },
    },
    interceptors: {
      request: {
        use: sandbox.stub(),
      },
      response: {
        use: sandbox.stub(),
      },
    },
    delete: function <T = any, R = AxiosResponse<T>>(
      url: string,
      config?: AxiosRequestConfig
    ): Promise<R> {
      const response = axiosDeleteResponses[url] as any;
      return response.message !== undefined ? Promise.reject(response) : Promise.resolve(response);
    },
    get: function <T = any, R = AxiosResponse<T>>(url: string): Promise<R> {
      const response = axiosGetResponses[url] as any;
      return response.message !== undefined ? Promise.reject(response) : Promise.resolve(response);
    },
    post: function <T = any, R = AxiosResponse<T>>(
      url: string,
      data?: any,
      config?: AxiosRequestConfig
    ): Promise<R> {
      const response = axiosPostResponses[url] as any;
      return response.message !== undefined ? Promise.reject(response) : Promise.resolve(response);
    },
  } as any as AxiosInstance;

  afterEach(() => {
    sandbox.restore();
  });

  beforeEach(() => {
    axiosDeleteResponses = {};
    axiosGetResponses = {};
    axiosPostResponses = {};
    sandbox.stub(fs, "readFile").callsFake((file) => {
      return Promise.resolve(Buffer.from("test"));
    });
    sandbox.stub(axios, "create").returns(testAxiosInstance);

    setTools({} as any);
    process.env["TEAMSFX_BUILDER_API"] = "1";
  });

  it("GetSharedInstance happy path", () => {
    let instance = PackageService.GetSharedInstance();
    chai.assert.isDefined(instance);
    instance = PackageService.GetSharedInstance();
    chai.assert.isDefined(instance);
  });

  it("sideLoadXmlManifest happy path with 200 return code", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosPostResponses["/dev/v1/users/packages/addins"] = {
      status: 200,
      data: {
        titleId: "test-title-id",
        appId: "test-app-id",
      },
    };

    const infoStub = sandbox.stub(logger, "info").returns();
    const verboseStub = sandbox.stub(logger, "verbose").returns();
    let packageService = new PackageService("https://test-endpoint", logger);
    let actualError: Error | undefined;
    try {
      const result = await packageService.sideLoadXmlManifest("test-token", "test-path");
      chai.assert.equal(result[0], "test-title-id");
      chai.assert.equal(result[1], "test-app-id");
      chai.assert.isTrue(infoStub.calledWith("TitleId: test-title-id"));
      chai.assert.isTrue(infoStub.calledWith("AppId: test-app-id"));
      chai.assert.isTrue(verboseStub.calledWith("Sideloading done."));
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);

    // Test with logger undefined
    packageService = new PackageService("https://test-endpoint", undefined);
    actualError = undefined;
    try {
      const result = await packageService.sideLoadXmlManifest("test-token", "test-path");
      chai.assert.equal(result[0], "test-title-id");
      chai.assert.equal(result[1], "test-app-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
  });

  it("sideLoadXmlManifest happy path with 202 return code", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosPostResponses["/dev/v1/users/packages/addins"] = {
      status: 202,
      data: {
        statusId: "test-status-id",
      },
    };

    axiosGetResponses["/dev/v1/users/packages/status/test-status-id"] = {
      status: 200,
      data: {
        titleId: "test-title-id",
        appId: "test-app-id",
      },
    };

    const infoStub = sandbox.stub(logger, "info").returns();
    const verboseStub = sandbox.stub(logger, "verbose").returns();
    const debugStub = sandbox.stub(logger, "debug").returns();
    let packageService = new PackageService("https://test-endpoint", logger);
    let actualError: Error | undefined;
    try {
      const result = await packageService.sideLoadXmlManifest("test-token", "test-path");
      chai.assert.equal(result[0], "test-title-id");
      chai.assert.equal(result[1], "test-app-id");
      chai.assert.isTrue(
        debugStub.calledWith("Acquiring package with statusId: test-status-id ...")
      );
      chai.assert.isTrue(debugStub.calledWith("Package status: 200 ..."));
      chai.assert.isTrue(infoStub.calledWith("TitleId: test-title-id"));
      chai.assert.isTrue(infoStub.calledWith("AppId: test-app-id"));
      chai.assert.isTrue(verboseStub.calledWith("Sideloading done."));
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);

    // Test with logger undefined
    packageService = new PackageService("https://test-endpoint", undefined);
    actualError = undefined;
    try {
      const result = await packageService.sideLoadXmlManifest("test-token", "test-path");
      chai.assert.equal(result[0], "test-title-id");
      chai.assert.equal(result[1], "test-app-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
  });

  it("sideLoadXmlManifest happy path with xml api 200 return code, status api with 202 on first try and 200 on second try", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosPostResponses["/dev/v1/users/packages/addins"] = {
      status: 202,
      data: {
        statusId: "test-status-id",
      },
    };

    sandbox
      .stub(testAxiosInstance, "get")
      .withArgs("/dev/v1/users/packages/status/test-status-id", {
        baseURL: "https://test-url",
        headers: { Authorization: `Bearer test-token` },
      })
      .onFirstCall()
      .resolves({
        status: 202,
      })
      .onSecondCall()
      .resolves({
        status: 200,
        data: {
          titleId: "test-title-id",
          appId: "test-app-id",
        },
      })
      .withArgs("/config/v1/environment", {
        baseURL: "https://test-endpoint",
        headers: { Authorization: `Bearer test-token` },
      })
      .resolves({
        data: {
          titlesServiceUrl: "https://test-url",
        },
      });

    const infoStub = sandbox.stub(logger, "info").returns();
    const verboseStub = sandbox.stub(logger, "verbose").returns();
    const debugStub = sandbox.stub(logger, "debug").returns();
    const packageService = new PackageService("https://test-endpoint", logger);
    let actualError: Error | undefined;
    try {
      const result = await packageService.sideLoadXmlManifest("test-token", "test-path");
      chai.assert.equal(result[0], "test-title-id");
      chai.assert.equal(result[1], "test-app-id");
      chai.assert.isTrue(
        debugStub.calledWith("Acquiring package with statusId: test-status-id ...")
      );
      chai.assert.isTrue(debugStub.calledWith("Package status: 200 ..."));
      chai.assert.isTrue(infoStub.calledWith("TitleId: test-title-id"));
      chai.assert.isTrue(infoStub.calledWith("AppId: test-app-id"));
      chai.assert.isTrue(verboseStub.calledWith("Sideloading done."));
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
  });

  it("sideLoadXmlManifest xml api with non 200/202 return code", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosPostResponses["/dev/v1/users/packages/addins"] = {
      status: 203,
      data: {
        statusId: "test-status-id",
      },
    };

    const infoStub = sandbox.stub(logger, "info").returns();
    const verboseStub = sandbox.stub(logger, "verbose").returns();
    const debugStub = sandbox.stub(logger, "debug").returns();
    const errorStub = sandbox.stub(logger, "error").returns();
    const packageService = new PackageService("https://test-endpoint", logger);
    let actualError: Error | undefined;
    try {
      const result = await packageService.sideLoadXmlManifest("test-token", "test-path");
      chai.assert.equal(result[0], "test-title-id");
      chai.assert.equal(result[1], "test-app-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isFalse(debugStub.calledWith("Package status: 200 ..."));
    chai.assert.isFalse(infoStub.calledWith("TitleId: test-title-id"));
    chai.assert.isFalse(infoStub.calledWith("AppId: test-app-id"));
    chai.assert.isFalse(verboseStub.calledWith("Sideloading done."));
    // chai.assert.isTrue(errorStub.calledWith("Sideloading failed."));

    chai.assert.isDefined(actualError);
  });

  it("sideLoadXmlManifest xml upload api throws error with response", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    const error: any = new Error("test-post");
    error.response = {
      data: {},
    };
    axiosPostResponses["/dev/v1/users/packages/addins"] = error;

    const errorStub = sandbox.stub(logger, "error").returns();
    const packageService = new PackageService("https://test-endpoint", logger);
    let actualError: Error | undefined;
    try {
      await packageService.sideLoadXmlManifest("test-token", "test-path");
    } catch (error: any) {
      actualError = error;
    }
    // chai.assert.isTrue(errorStub.calledWith(`${JSON.stringify(error.response.data)}`));
    // chai.assert.isTrue(errorStub.calledWith(`Sideloading failed.`));
    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError?.message.includes("test-post"));
  });

  it("sideLoadXmlManifest xml upload api throws error without response", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    const error: Error = new Error("test-post");
    axiosPostResponses["/dev/v1/users/packages/addins"] = error;

    const errorStub = sandbox.stub(logger, "error").returns();
    const packageService = new PackageService("https://test-endpoint", logger);
    let actualError: Error | undefined;
    try {
      await packageService.sideLoadXmlManifest("test-token", "test-path");
    } catch (error: any) {
      actualError = error;
    }
    // chai.assert.isTrue(errorStub.calledWith(`test-post`));
    chai.assert.isDefined(actualError);
    // chai.assert.isTrue(actualError?.message.includes("test-post"));
  });

  it("sideLoading happy path", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosPostResponses["/dev/v1/users/packages"] = {
      data: {
        operationId: "test-operation-id",
        titlePreview: {
          titleId: "test-title-id-preview",
        },
      },
    };
    axiosPostResponses["/builder/v1/users/packages"] = {
      data: {
        statusId: "test-status-id-builder-api",
        titlePreview: {
          titleId: "test-title-id-preview-builder-api",
        },
      },
    };
    axiosPostResponses["/dev/v1/users/packages/acquisitions"] = {
      data: {
        statusId: "test-status-id",
      },
    };
    axiosGetResponses["/dev/v1/users/packages/status/test-status-id"] = {
      status: 200,
      data: {
        titleId: "test-title-id",
        appId: "test-app-id",
      },
    };
    axiosGetResponses["/builder/v1/users/packages/status/test-status-id-builder-api"] = {
      status: 200,
      data: {
        titleId: "test-title-id-builder-api",
        appId: "test-app-id-builder-api",
      },
    };
    axiosGetResponses["/marketplace/v1/users/titles/test-title-id-builder-api/sharingInfo"] = {
      data: {
        unifiedStoreLink: "https://test-share-link",
      },
    };

    let packageService = new PackageService("https://test-endpoint");
    sandbox.stub(packageService, "getManifestFromZip" as keyof PackageService).returns({} as any);
    let actualError: Error | undefined;
    try {
      const result = await packageService.sideLoading("test-token", "test-path");
      chai.assert.equal(result[0], "test-title-id");
      chai.assert.equal(result[1], "test-app-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
    packageService = new PackageService("https://test-endpoint", logger);
    sandbox.stub(packageService, "getManifestFromZip" as keyof PackageService).returns({
      $schema:
        "https://developer.microsoft.com/json-schemas/teams/v1.19/MicrosoftTeams.schema.json",
      manifestVersion: "1.19",
      version: "1.0.0",
      id: "${{TEAMS_APP_ID}}",
      developer: {
        name: "Teams App, Inc.",
        websiteUrl: "https://www.example.com",
        privacyUrl: "https://www.example.com/privacy",
        termsOfUseUrl: "https://www.example.com/termofuse",
      },
      icons: {
        color: "color.png",
        outline: "outline.png",
      },
      name: {
        short: "test-manifest",
        full: "test-manifest full name",
      },
      description: {
        short: "Short description for test-manifest",
        full: "Full description for test-manifest",
      },
      accentColor: "#FFFFFF",
      composeExtensions: [],
      permissions: ["identity", "messageTeamMembers"],
    } as any);
    try {
      const result = await packageService.sideLoading("test-token", "test-path");
      chai.assert.equal(result[0], "test-title-id");
      chai.assert.equal(result[1], "test-app-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);

    packageService = new PackageService("https://test-endpoint", logger);
    sandbox.stub(packageService, "getManifestFromZip" as keyof PackageService).returns({
      copilotAgents: {
        declarativeAgents: [
          {
            id: "declarativeAgent",
            file: "declarativeAgent.json",
          },
        ],
      },
    } as any);
    try {
      const result = await packageService.sideLoading("test-token", "test-path", AppScope.Shared);
      chai.assert.equal(result[0], "test-title-id-builder-api");
      chai.assert.equal(result[1], "test-app-id-builder-api");
      chai.assert.equal(result[2], "https://test-share-link");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);

    // without logger
    packageService = new PackageService("https://test-endpoint");
    sandbox.stub(packageService, "getManifestFromZip" as keyof PackageService).returns({
      copilotAgents: {
        declarativeAgents: [
          {
            id: "declarativeAgent",
            file: "declarativeAgent.json",
          },
        ],
      },
    } as any);
    try {
      const result = await packageService.sideLoading("test-token", "test-path");
      chai.assert.equal(result[0], "test-title-id-builder-api");
      chai.assert.equal(result[1], "test-app-id-builder-api");
      chai.assert.equal(result[2], "");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);

    packageService = new PackageService("https://test-endpoint");
    try {
      const result = await packageService.sideLoading(
        "test-token",
        "./tests/component/m365/success.zip"
      );
      chai.assert.equal(result[0], "test-title-id");
      chai.assert.equal(result[1], "test-app-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
  });

  it("sideloading throws error in get status", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };

    axiosPostResponses["/builder/v1/users/packages"] = {
      data: {
        statusId: "test-status-id-builder-api",
        titlePreview: {
          titleId: "test-title-id-preview-builder-api",
        },
      },
    };
    let actualError: Error | undefined;
    const packageService = new PackageService("https://test-endpoint", logger);
    sandbox.stub(packageService, "getManifestFromZip" as keyof PackageService).returns({
      copilotAgents: {
        declarativeAgents: [
          {
            id: "declarativeAgent",
            file: "declarativeAgent.json",
          },
        ],
      },
    } as any);
    try {
      const result = await packageService.sideLoading("test-token", "test-path", AppScope.Shared);
    } catch (error: any) {
      actualError = error;
    }
    chai.assert.isDefined(actualError);

    const expectedError = new Error("test-status") as any;
    expectedError.response = {
      data: {
        foo: "bar",
      },
      headers: {
        traceresponse: "tracing-id",
      },
    };
    axiosGetResponses["/builder/v1/users/packages/status/test-status-id-builder-api"] =
      expectedError;
    actualError = undefined;
    try {
      const result = await packageService.sideLoading("test-token", "test-path", AppScope.Shared);
    } catch (error: any) {
      actualError = error;
    }
    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError?.message.includes("test-status"));
  });
  it("sideLoading throws expected error", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosPostResponses["/dev/v1/users/packages"] = new Error("test-post");
    axiosPostResponses["/builder/v1/users/packages"] = new Error("test-post-builder-api");

    let packageService = new PackageService("https://test-endpoint");
    sandbox.stub(packageService, "getManifestFromZip" as keyof PackageService).returns({} as any);
    let actualError: Error | undefined;
    try {
      await packageService.sideLoading("test-token", "test-path");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError?.message.includes("test-post"));

    packageService = new PackageService("https://test-endpoint", logger);
    sandbox.stub(packageService, "getManifestFromZip" as keyof PackageService).returns({} as any);
    actualError = undefined;
    try {
      await packageService.sideLoading("test-token", "test-path");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError?.message.includes("test-post"));

    packageService = new PackageService("https://test-endpoint", logger);
    sandbox.stub(packageService, "getManifestFromZip" as keyof PackageService).returns({
      copilotAgents: {
        declarativeAgents: [
          {
            id: "declarativeAgent",
            file: "declarativeAgent.json",
          },
        ],
      },
    } as any);
    actualError = undefined;
    try {
      await packageService.sideLoading("test-token", "test-path");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError?.message.includes("test-post-builder-api"));

    packageService = new PackageService("https://test-endpoint", logger);
    sandbox
      .stub(packageService, "getManifestFromZip" as keyof PackageService)
      .returns(undefined as any);
    actualError = undefined;
    try {
      await packageService.sideLoading("test-token", "test-path");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(
      actualError?.message.includes("Invalid app package zip. manifest.json is missing")
    );
  });

  it("sideLoading throws expected reponse error", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    const expectedError = new Error("test-post") as any;
    expectedError.response = {
      data: {
        foo: "bar",
      },
      headers: {
        traceresponse: "tracing-id",
      },
    };
    axiosPostResponses["/dev/v1/users/packages"] = expectedError;

    let packageService = new PackageService("https://test-endpoint");
    sandbox.stub(packageService, "getManifestFromZip" as keyof PackageService).returns({} as any);
    let actualError: any;
    try {
      await packageService.sideLoading("test-token", "test-path");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError.message.includes("test-post"));

    packageService = new PackageService("https://test-endpoint", logger);
    sandbox.stub(packageService, "getManifestFromZip" as keyof PackageService).returns({} as any);
    actualError = undefined;
    try {
      await packageService.sideLoading("test-token", "test-path");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError.message.includes("test-post"));
  });

  it("sideLoading badrequest as user error", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    const expectedError = new Error("test-post") as any;
    expectedError.response = {
      data: {
        foo: "bar",
      },
      headers: {
        traceresponse: "tracing-id",
      },
      status: 400,
    };
    axiosPostResponses["/dev/v1/users/packages"] = expectedError;

    const packageService = new PackageService("https://test-endpoint");
    sandbox.stub(packageService, "getManifestFromZip" as keyof PackageService).returns({} as any);
    let actualError: any;
    try {
      await packageService.sideLoading("test-token", "test-path");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError.message.includes("test-post"));
    chai.assert.isTrue(actualError instanceof UserError);
  });

  it("sideLoading Builder API Feature Flag turned off", async () => {
    process.env["TEAMSFX_BUILDER_API"] = "0";
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosPostResponses["/dev/v1/users/packages"] = {
      data: {
        operationId: "test-operation-id",
        titlePreview: {
          titleId: "test-title-id-preview",
        },
      },
    };
    axiosPostResponses["/dev/v1/users/packages/acquisitions"] = {
      data: {
        statusId: "test-status-id",
      },
    };
    axiosGetResponses["/dev/v1/users/packages/status/test-status-id"] = {
      status: 200,
      data: {
        titleId: "test-title-id",
        appId: "test-app-id",
      },
    };
    axiosGetResponses["/marketplace/v1/users/titles/test-title-id-builder-api/sharingInfo"] = {
      data: {
        unifiedStoreLink: "https://test-share-link",
      },
    };

    let actualError: Error | undefined;
    const packageService = new PackageService("https://test-endpoint", logger);
    sandbox.stub(packageService, "getManifestFromZip" as keyof PackageService).returns({
      copilotAgents: {
        declarativeAgents: [
          {
            id: "declarativeAgent",
            file: "declarativeAgent.json",
          },
        ],
      },
    } as any);
    try {
      const result = await packageService.sideLoading("test-token", "test-path", AppScope.Shared);
      chai.assert.equal(result[0], "test-title-id");
      chai.assert.equal(result[1], "test-app-id");
      chai.assert.equal(result[2], "");
    } catch (error: any) {
      actualError = error;
    }
    chai.assert.isUndefined(actualError);
  });

  it("retrieveTitleId happy path", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosPostResponses["/catalog/v1/users/titles/launchInfo"] = {
      data: {
        acquisition: {
          titleId: "test-title-id",
        },
      },
    };

    const packageService = new PackageService("https://test-endpoint");
    const titleId = await packageService.retrieveTitleId("test-token", "test-manifest-id");

    chai.assert.equal(titleId, "test-title-id");
  });

  it("retrieveTitleId throws expected error", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosPostResponses["/catalog/v1/users/titles/launchInfo"] = new Error("test-post");

    const packageService = new PackageService("https://test-endpoint");
    let actualError: Error | undefined;
    try {
      await packageService.retrieveTitleId("test-token", "test-manifest-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError?.message.includes("test-post"));
  });

  it("retrieveTitleId throws expected response error", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    const expectedError = new Error("test-post") as any;
    expectedError.response = {
      data: {},
    };
    axiosPostResponses["/catalog/v1/users/titles/launchInfo"] = expectedError;

    const packageService = new PackageService("https://test-endpoint");
    let actualError: any;
    try {
      await packageService.retrieveTitleId("test-token", "test-manifest-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError.message.includes("test-post"));
  });

  it("retrieveAppId happy path", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosPostResponses["/catalog/v1/users/titles/launchInfo"] = {
      data: {
        acquisition: {
          appId: "test-app-id",
        },
      },
    };

    {
      const packageService = new PackageService("https://test-endpoint");
      const appId = await packageService.retrieveAppId("test-token", "test-manifest-id");

      chai.assert.equal(appId, "test-app-id");
    }

    {
      const packageService = new PackageService("https://test-endpoint", new MockLogProvider());
      const appId = await packageService.retrieveAppId("test-token", "test-manifest-id");

      chai.assert.equal(appId, "test-app-id");
    }
  });

  it("retrieveAppId throws expected error", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosPostResponses["/catalog/v1/users/titles/launchInfo"] = new Error("test-post");

    {
      const packageService = new PackageService("https://test-endpoint");
      let actualError: Error | undefined;
      try {
        await packageService.retrieveAppId("test-token", "test-manifest-id");
      } catch (error: any) {
        actualError = error;
      }

      chai.assert.isDefined(actualError);
      chai.assert.isTrue(actualError?.message.includes("test-post"));
    }

    {
      const packageService = new PackageService("https://test-endpoint", new MockLogProvider());
      let actualError: Error | undefined;
      try {
        await packageService.retrieveAppId("test-token", "test-manifest-id");
      } catch (error: any) {
        actualError = error;
      }

      chai.assert.isDefined(actualError);
      chai.assert.isTrue(actualError?.message.includes("test-post"));
    }
  });

  it("retrieveAppId throws expected response error", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    const expectedError = new Error("test-post") as any;
    expectedError.response = {
      data: {},
    };
    axiosPostResponses["/catalog/v1/users/titles/launchInfo"] = expectedError;

    {
      const packageService = new PackageService("https://test-endpoint");
      let actualError: any;
      try {
        await packageService.retrieveAppId("test-token", "test-manifest-id");
      } catch (error: any) {
        actualError = error;
      }

      chai.assert.isDefined(actualError);
      chai.assert.isTrue(actualError.message.includes("test-post"));
    }

    {
      const packageService = new PackageService("https://test-endpoint", new MockLogProvider());
      let actualError: any;
      try {
        await packageService.retrieveAppId("test-token", "test-manifest-id");
      } catch (error: any) {
        actualError = error;
      }

      chai.assert.isDefined(actualError);
      chai.assert.isTrue(actualError instanceof UnhandledError);
    }
  });

  it("unacquire happy path", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosDeleteResponses["/catalog/v1/users/acquisitions/test-title-id"] = {};

    let packageService = new PackageService("https://test-endpoint");
    let actualError: Error | undefined;
    try {
      await packageService.unacquire("test-token", "test-title-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);

    packageService = new PackageService("https://test-endpoint", logger);
    actualError = undefined;
    try {
      await packageService.unacquire("test-token", "test-title-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
  });
  it("unacquire throws expected error", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosDeleteResponses["/catalog/v1/users/acquisitions/test-title-id"] = new Error("test-delete");

    let packageService = new PackageService("https://test-endpoint");
    let actualError: Error | undefined;
    try {
      await packageService.unacquire("test-token", "test-title-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError?.message.includes("test-delete"));

    packageService = new PackageService("https://test-endpoint", logger);
    actualError = undefined;
    try {
      await packageService.unacquire("test-token", "test-title-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError?.message.includes("test-delete"));
  });

  it("unacquire throws expected response error", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    const expectedError = new Error("test-post") as any;
    expectedError.response = {
      data: {},
    };
    axiosDeleteResponses["/catalog/v1/users/acquisitions/test-title-id"] = expectedError;

    const packageService = new PackageService("https://test-endpoint");
    let actualError: any;
    try {
      await packageService.unacquire("test-token", "test-title-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError instanceof UnhandledError);
  });

  it("getLaunchInfoByTitleId happy path", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosGetResponses["/catalog/v1/users/titles/test-title-id/launchInfo"] = {
      data: {
        foo: "bar",
      },
    };

    const packageService = new PackageService("https://test-endpoint");
    const launchInfo = await packageService.getLaunchInfoByTitleId("test-token", "test-title-id");

    chai.assert.deepEqual(launchInfo, { foo: "bar" });
  });
  it("getLaunchInfoByManifestId throws expected error", async () => {
    const packageService = new PackageService("https://test-endpoint");
    sandbox.stub(testAxiosInstance, "post").rejects({ response: { status: 404 } });
    sandbox.stub(packageService, "getTitleServiceUrl").resolves("https://test-url");
    try {
      await packageService.getLaunchInfoByManifestId("test-token", "test-manifest-id");
      chai.assert.fail("should not reach here");
    } catch (e) {
      chai.assert.isTrue(e instanceof NotExtendedToM365Error);
    }
  });
  it("getLaunchInfoByTitleId throws expected error", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosGetResponses["/catalog/v1/users/titles/test-title-id/launchInfo"] = new Error("test-get");

    const packageService = new PackageService("https://test-endpoint");
    let actualError: Error | undefined;
    try {
      await packageService.getLaunchInfoByTitleId("test-token", "test-title-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError?.message.includes("test-get"));
  });

  it("getLaunchInfoByTitleId throws expected response error", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    const expectedError = new Error("test-post") as any;
    expectedError.response = {
      data: {},
    };
    axiosGetResponses["/catalog/v1/users/titles/test-title-id/launchInfo"] = expectedError;

    const packageService = new PackageService("https://test-endpoint");
    let actualError: any;
    try {
      await packageService.getLaunchInfoByTitleId("test-token", "test-title-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError instanceof UnhandledError);
  });

  it("getTitleServiceUrl throws expected error", async () => {
    axiosGetResponses["/config/v1/environment"] = new Error("test-service-url-error");

    const packageService = new PackageService("https://test-endpoint");
    let actualError: Error | undefined;
    try {
      await packageService.getLaunchInfoByTitleId("test-token", "test-title-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError?.message.includes("test-service-url-error"));
  });

  it("getTitleServiceUrl throws invalid url error", async () => {
    let packageService = new PackageService("{{test-endpoint}}");
    let actualError: Error | undefined;
    try {
      await packageService.getLaunchInfoByTitleId("test-token", "test-title-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError?.message.includes("Invalid URL"));

    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "{{test-url}}",
      },
    };

    packageService = new PackageService("https://test-endpoint");
    actualError = undefined;
    try {
      await packageService.getLaunchInfoByTitleId("test-token", "test-title-id");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
  });

  it("getActiveExperiences happy path", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosGetResponses["/catalog/v1/users/uitypes"] = {
      data: {
        activeExperiences: ["foo", "bar"],
      },
    };

    let packageService = new PackageService("https://test-endpoint");
    let actualError: Error | undefined;
    let result: string[] | undefined;
    try {
      result = await packageService.getActiveExperiences("test-token");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
    chai.assert.deepEqual(result, ["foo", "bar"]);

    packageService = new PackageService("https://test-endpoint", logger);
    actualError = undefined;
    try {
      result = await packageService.getActiveExperiences("test-token");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
    chai.assert.deepEqual(result, ["foo", "bar"]);
  });

  it("getActiveExperiences stale", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosGetResponses["/catalog/v1/users/uitypes"] = {
      data: {
        activeExperiences: ["foo", "bar"],
        nextInterval: 1,
      },
    };

    let packageService = new PackageService("https://test-endpoint");
    let actualError: Error | undefined;
    let result: string[] | undefined;
    try {
      result = await packageService.getActiveExperiences("test-token", true);
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
    chai.assert.deepEqual(result, ["foo", "bar"]);

    const debugStub = sandbox.stub(logger, "debug").returns();

    packageService = new PackageService("https://test-endpoint", logger);
    try {
      result = await packageService.getActiveExperiences("test-token", true);
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
    chai.assert.deepEqual(result, ["foo", "bar"]);
    chai.assert.equal(5, debugStub.getCalls().length);
  });

  it("getActiveExperiences throws expected error", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosGetResponses["/catalog/v1/users/uitypes"] = new Error("test-get");

    const packageService = new PackageService("https://test-endpoint");
    let actualError: Error | undefined;
    try {
      await packageService.getActiveExperiences("test-token");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError?.message.includes("test-get"));
  });

  it("getActiveExperiences throws expected response error", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    const expectedError = new Error("test-get") as any;
    expectedError.response = {
      data: {},
    };
    axiosGetResponses["/catalog/v1/users/uitypes"] = expectedError;

    let packageService = new PackageService("https://test-endpoint");
    let actualError: Error | undefined;
    try {
      await packageService.getActiveExperiences("test-token");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError instanceof UnhandledError);

    packageService = new PackageService("https://test-endpoint", logger);
    actualError = undefined;
    try {
      await packageService.getActiveExperiences("test-token");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isDefined(actualError);
    chai.assert.isTrue(actualError instanceof UnhandledError);
  });

  it("getCopilotStatus happy path", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosGetResponses["/catalog/v1/users/uitypes"] = {
      data: {
        activeExperiences: ["foo", "bar"],
      },
    };

    const packageService = new PackageService("https://test-endpoint");
    let actualError: Error | undefined;
    let result: boolean | undefined;
    try {
      result = await packageService.getCopilotStatus("test-token");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
    chai.assert.isFalse(result);
  });

  it("getCopilotStatus bad response", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosGetResponses["/catalog/v1/users/uitypes"] = {
      foo: "bar",
    };

    const packageService = new PackageService("https://test-endpoint");
    let actualError: Error | undefined;
    let result: boolean | undefined;
    try {
      result = await packageService.getCopilotStatus("test-token");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
    chai.assert.isUndefined(result);
  });

  it("getCopilotStatus returns undefined on error", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosGetResponses["/catalog/v1/users/uitypes"] = new Error("test-get");

    let packageService = new PackageService("https://test-endpoint");
    let actualError: Error | undefined;
    let result: boolean | undefined;
    try {
      result = await packageService.getCopilotStatus("test-token");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
    chai.assert.isUndefined(result);

    packageService = new PackageService("https://test-endpoint", logger);
    actualError = undefined;
    try {
      result = await packageService.getCopilotStatus("test-token");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
    chai.assert.isUndefined(result);
  });

  it("getCopilotStatus returns undefined on error with trace", async () => {
    const packageService = new PackageService("https://test-endpoint");
    (packageService as any).getActiveExperiences = async (_: string) => {
      const error = new Error();
      (error as any).response = {
        headers: {
          traceresponse: "test-trace",
        },
      };
      throw error;
    };
    let actualError: Error | undefined;
    let result: boolean | undefined;
    try {
      result = await packageService.getCopilotStatus("test-token");
    } catch (error: any) {
      actualError = error;
    }

    chai.assert.isUndefined(actualError);
    chai.assert.isUndefined(result);
  });

  it("get share link happy path", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    axiosGetResponses["/marketplace/v1/users/titles/test-title-id/sharingInfo"] = {
      data: {
        unifiedStoreLink: "https://test-share-link",
      },
    };
    const packageService = new PackageService("https://test-endpoint");
    const shareLink = await packageService.getShareLink("test-token", "test-title-id");
    chai.assert.equal(shareLink, "https://test-share-link");
  });

  it("get share link - failure", async () => {
    axiosGetResponses["/config/v1/environment"] = {
      data: {
        titlesServiceUrl: "https://test-url",
      },
    };
    const packageService = new PackageService("https://test-endpoint");
    let actualError: boolean | undefined;
    try {
      const shareLink = await packageService.getShareLink("test-token", "test-title-id");
    } catch (error: any) {
      actualError = error;
    }
    chai.assert.isDefined(actualError);
  });
});
