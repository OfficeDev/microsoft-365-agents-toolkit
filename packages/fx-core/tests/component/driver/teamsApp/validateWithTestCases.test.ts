// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, ok, Platform, SystemError, TeamsAppManifest } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import chai from "chai";
import fs from "fs-extra";
import "mocha";
import * as sinon from "sinon";
import { teamsDevPortalClient } from "../../../../src/client/teamsDevPortalClient";
import { setTools } from "../../../../src/common/globalVars";
import * as commonTools from "../../../../src/common/utils";
import { Constants } from "../../../../src/component/driver/teamsApp/constants";
import { AppStudioError } from "../../../../src/component/driver/teamsApp/errors";
import {
  AsyncAppValidationResponse,
  AsyncAppValidationStatus,
} from "../../../../src/component/driver/teamsApp/interfaces/AsyncAppValidationResponse";
import { AsyncAppValidationResultsResponse } from "../../../../src/component/driver/teamsApp/interfaces/AsyncAppValidationResultsResponse";
import { ValidateWithTestCasesArgs } from "../../../../src/component/driver/teamsApp/interfaces/ValidateWithTestCasesArgs";
import { teamsappMgr } from "../../../../src/component/driver/teamsApp/teamsappMgr";
import { ValidateWithTestCasesDriver } from "../../../../src/component/driver/teamsApp/validateTestCases";
import { metadataUtil } from "../../../../src/component/utils/metadataUtil";
import { InvalidActionInputError, UserCancelError } from "../../../../src/error/common";
import { MockedM365Provider, MockTools } from "../../../core/utils";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";

describe("teamsApp/validateWithTestCases", async () => {
  const tools = new MockTools();
  setTools(tools);

  const teamsAppDriver = new ValidateWithTestCasesDriver();

  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
    projectPath: "./",
  };

  beforeEach(() => {
    sinon.stub(commonTools, "waitSeconds").resolves();
  });

  afterEach(() => {
    sinon.restore();
  });

  it("file not found - app package", async () => {
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(AppStudioError.FileNotFoundError.name, result.error.name);
    }
  });

  it("file not found - manifest.json", async () => {
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(AppStudioError.FileNotFoundError.name, result.error.name);
    }
  });

  it("invalid param error", async () => {
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof InvalidActionInputError);
    }
  });

  it("Failed to get token", async () => {
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakePath",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");
    sinon
      .stub(mockedDriverContext.m365TokenProvider, "getAccessToken")
      .resolves(err(new SystemError({})));

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
  });

  it("Invalid validation result response - Null details", async () => {
    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves(undefined);
    const mockSubmitValidationResponse: AsyncAppValidationResponse = {
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    };
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const invalidValidationResultResponseJson: any = {
      appValidationId: "appValidationId123",
      appId: "appId123",
      status: "Completed",
      appVersion: "1.0.0",
      manifestVersion: "1.0.0",
      createdAt: "2024-03-27T12:00:00.000Z",
      updatedAt: "2024-03-27T12:00:00.000Z",
      validationResults: {
        successes: null,
        warnings: null,
        failures: null,
        skipped: null,
      },
    };
    const invalidValidationResultResponse: AsyncAppValidationResultsResponse = <
      AsyncAppValidationResultsResponse
    >invalidValidationResultResponseJson;
    sinon
      .stub(teamsDevPortalClient, "getAppValidationById")
      .resolves(invalidValidationResultResponse);
    await teamsAppDriver.runningBackgroundJob(
      args,
      mockedDriverContext,
      "test_token",
      mockSubmitValidationResponse,
      "test_id"
    );
    chai.assert(
      mockedDriverContext.logProvider.msg.includes("Validation request completed, status:")
    );
  });

  it("Invalid validation result response - Null validation results", async () => {
    const mockSubmitValidationResponse: AsyncAppValidationResponse = {
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    };
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const invalidValidationResultResponseJson: any = {
      appValidationId: "appValidationId123",
      appId: "appId123",
      status: "Completed",
      appVersion: "1.0.0",
      manifestVersion: "1.0.0",
      createdAt: "2024-03-27T12:00:00.000Z",
      updatedAt: "2024-03-27T12:00:00.000Z",
      validationResults: null,
    };
    const invalidValidationResultResponse: AsyncAppValidationResultsResponse = <
      AsyncAppValidationResultsResponse
    >invalidValidationResultResponseJson;
    sinon
      .stub(teamsDevPortalClient, "getAppValidationById")
      .resolves(invalidValidationResultResponse);
    await teamsAppDriver.runningBackgroundJob(
      args,
      mockedDriverContext,
      "test_token",
      mockSubmitValidationResponse,
      "test_id"
    );
    chai.assert(
      mockedDriverContext.logProvider.msg.includes("Validation request completed, status:")
    );
  });

  it("Valid validation result response", async () => {
    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({
      appValidations: [
        {
          id: "fakeId",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Completed,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
        {
          id: "fakeId2",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Aborted,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      ],
    });
    const mockSubmitValidationResponse: AsyncAppValidationResponse = {
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    };
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };
    sinon.stub(teamsDevPortalClient, "getAppValidationById").resolves({
      status: AsyncAppValidationStatus.Completed,
      appValidationId: "fakeId",
      appId: "fakeAppId",
      appVersion: "1.0.0",
      manifestVersion: "1.17",
      validationResults: {
        successes: [
          {
            title: "Validation_Success_Example",
            message: "Success validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        warnings: [
          {
            title: "Validation_Warning_Example",
            message: "Warning validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        failures: [
          {
            title: "Validation_Failure_Example",
            message: "Failure validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        skipped: [
          {
            title: "Validation_Skipped_Example",
            message: "Skipped validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    });
    await teamsAppDriver.runningBackgroundJob(
      args,
      mockedDriverContext,
      "test_token",
      mockSubmitValidationResponse,
      "test_id"
    );
    chai.assert(
      mockedDriverContext.logProvider.msg.includes("Validation request completed, status:")
    );
    chai.assert(
      mockedDriverContext.logProvider.msg.includes("1 failed, 1 warning, 1 skipped, 1 passed")
    );
  });

  it("Duplicate validations - InProgress", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({
      appValidations: [
        {
          id: "fakeId",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Completed,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
        {
          id: "fakeId2",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.InProgress,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      ],
    });
    sinon.stub(teamsDevPortalClient, "submitAppValidationRequest").throws("should not be called");
    sinon.stub(teamsDevPortalClient, "getAppValidationById").throws("should not be called");

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("Duplicate validations - Created", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({
      appValidations: [
        {
          id: "fakeId",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Completed,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
        {
          id: "fakeId2",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Created,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      ],
    });
    sinon.stub(teamsDevPortalClient, "submitAppValidationRequest").throws("should not be called");
    sinon.stub(teamsDevPortalClient, "getAppValidationById").throws("should not be called");

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("Duplicate validations - CLI", async () => {
    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({
      appValidations: [
        {
          id: "fakeId",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Completed,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
        {
          id: "fakeId2",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.InProgress,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      ],
    });
    sinon.stub(teamsDevPortalClient, "submitAppValidationRequest").throws("should not be called");
    sinon.stub(teamsDevPortalClient, "getAppValidationById").throws("should not be called");

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("Invalid list validation response", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({});
    sinon.stub(teamsDevPortalClient, "submitAppValidationRequest").resolves({
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    });

    sinon.stub(teamsDevPortalClient, "getAppValidationById").resolves({
      status: AsyncAppValidationStatus.Completed,
      appValidationId: "fakeId",
      appId: "fakeAppId",
      appVersion: "1.0.0",
      manifestVersion: "1.17",
      validationResults: {
        successes: [
          {
            title: "Validation_Success_Example",
            message: "Success validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        warnings: [],
        failures: [],
        skipped: [],
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    });

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("Happy path", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({
      appValidations: [
        {
          id: "fakeId",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Completed,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
        {
          id: "fakeId2",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Aborted,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      ],
    });
    sinon.stub(teamsDevPortalClient, "submitAppValidationRequest").resolves({
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    });

    sinon.stub(teamsDevPortalClient, "getAppValidationById").resolves({
      status: AsyncAppValidationStatus.Completed,
      appValidationId: "fakeId",
      appId: "fakeAppId",
      appVersion: "1.0.0",
      manifestVersion: "1.17",
      validationResults: {
        successes: [
          {
            title: "Validation_Success_Example",
            message: "Success validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        warnings: [
          {
            title: "Validation_Warning_Example",
            message: "Warning validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        failures: [
          {
            title: "Validation_Failure_Example",
            message: "Failure validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        skipped: [
          {
            title: "Validation_Skipped_Example",
            message: "Skipped validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    });

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("Aborted", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({
      appValidations: [
        {
          id: "fakeId",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Completed,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
        {
          id: "fakeId2",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Aborted,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      ],
    });
    sinon.stub(teamsDevPortalClient, "submitAppValidationRequest").resolves({
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    });

    sinon.stub(teamsDevPortalClient, "getAppValidationById").resolves({
      status: AsyncAppValidationStatus.Aborted,
      appValidationId: "fakeId",
      appId: "fakeAppId",
      appVersion: "1.0.0",
      manifestVersion: "1.17",
      validationResults: {
        failures: [],
        warnings: [],
        successes: [],
        skipped: [],
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    });

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: false,
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("Happy path - CLI", async () => {
    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({
      appValidations: [
        {
          id: "fakeId",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Completed,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
        {
          id: "fakeId2",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Aborted,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      ],
    });
    sinon.stub(teamsDevPortalClient, "submitAppValidationRequest").resolves({
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    });

    sinon.stub(teamsDevPortalClient, "getAppValidationById").resolves({
      status: AsyncAppValidationStatus.Completed,
      appValidationId: "fakeId",
      appId: "fakeAppId",
      appVersion: "1.0.0",
      manifestVersion: "1.17",
      validationResults: {
        failures: [],
        warnings: [],
        successes: [],
        skipped: [],
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    });

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("CLI - succeed", async () => {
    sinon.stub(ValidateWithTestCasesDriver.prototype, "validate").resolves(ok(new Map()));
    const result = await teamsappMgr.validateTeamsApp({
      projectPath: "xxx",
      platform: Platform.CLI,
      "package-file": "xxx",
      "validate-method": "test-cases",
    });
    chai.assert(result.isOk());
  });

  it("CLI - failed", async () => {
    sinon
      .stub(ValidateWithTestCasesDriver.prototype, "validate")
      .resolves(err(new UserCancelError()));
    const result = await teamsappMgr.validateTeamsApp({
      projectPath: "xxx",
      platform: Platform.CLI,
      "package-file": "xxx",
      "validate-method": "test-cases",
    });
    chai.assert(result.isErr());
  });
});
