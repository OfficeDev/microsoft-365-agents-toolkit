// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import { createSandbox } from "sinon";
import axios from "axios";
import "mocha";
import mockedEnv from "mocked-env";
import fs from "fs-extra";
import path from "path";
import AdmZip from "adm-zip";
import { TeamsAppManifest } from "@microsoft/teamsfx-api";
import { InstallAppToChannelDriver } from "../../../../src/component/driver/devChannel/installApp";
import { GraphClient } from "../../../../src/client/graphClient";
import { MockedM365Provider, MockLogProvider } from "../../../core/utils";
import { WrapDriverContext } from "../../../../src/component/driver/util/wrapUtil";
import {
  InvalidActionInputError,
  FileNotFoundError,
  HttpClientError,
} from "../../../../src/error/common";
import { Constants } from "./../../../../src/component/driver/teamsApp/constants";
import { InstallAppArgs } from "../../../../build/component/driver/devChannel/interfaces/InstallAppArgs";

describe("InstallAppToChannelDriver", () => {
  const sandbox = createSandbox();
  const mockTokenProvider = new MockedM365Provider();
  const mockContext: WrapDriverContext = {
    m365TokenProvider: mockTokenProvider,
    logProvider: new MockLogProvider(),
    addSummary: sandbox.stub(),
    summaries: [],
    projectPath: "fake/project/path",
    addTelemetryProperties: sandbox.stub(),
  } as unknown as WrapDriverContext;

  const driver = new InstallAppToChannelDriver();

  const manifest = new TeamsAppManifest();
  manifest.id = "fake-id";
  const zip = new AdmZip();
  zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest)));
  zip.addFile("color.png", new Buffer(""));
  zip.addFile("outlie.png", new Buffer(""));

  const archivedFile = zip.toBuffer();

  beforeEach(() => {});

  afterEach(() => {
    sandbox.restore();
  });

  it("should return error if teamId or channelId is missing", async () => {
    const args = {} as any;
    const outputEnvVarNames = new Map<string, string>();

    const result = await driver.install(args, mockContext, outputEnvVarNames);

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error).to.be.instanceOf(InvalidActionInputError);
      expect(result.error.message).to.include("teamId");
      expect(result.error.message).to.include("channelId");
    }
  });

  it("should return error if app package file does not exist", async () => {
    sandbox.stub(fs, "pathExists").resolves(false);

    const args = {
      appPackagePath: "fake/path/app.zip",
      teamId: "fake-team-id",
      channelId: "fake-channel-id",
    };
    const outputEnvVarNames = new Map<string, string>();

    const result = await driver.install(args, mockContext, outputEnvVarNames);

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error).to.be.instanceOf(FileNotFoundError);
    }
  });

  it("should return error if manifest does not exist", async () => {
    sandbox.stub(fs, "readFile").callsFake(async () => {
      const emptyFile = new AdmZip().toBuffer();
      return emptyFile;
    });
    sandbox.stub(fs, "pathExists").resolves(true);

    const args = {
      appPackagePath: "fake/path/app.zip",
      teamId: "fake-team-id",
      channelId: "fake-channel-id",
    };
    const outputEnvVarNames = new Map<string, string>();

    const result = await driver.install(args, mockContext, outputEnvVarNames);

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error).to.be.instanceOf(FileNotFoundError);
    }
  });

  it("should install app to channel successfully", async () => {
    sandbox.stub(fs, "readFile").callsFake(async () => {
      return archivedFile;
    });
    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(GraphClient.prototype, "InstallAppToChannelAsync").resolves();
    sandbox.stub(GraphClient.prototype, "GetAppInstallationForTeam").resolves([]);

    const args: InstallAppArgs = {
      appPackagePath: "fake/path/app.zip",
      teamId: "fake-team-id",
      channelId: "fake-channel-id",
    };
    const outputEnvVarNames = new Map<string, string>();

    const result = await driver.install(args, mockContext, outputEnvVarNames);

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.size).to.equal(0);
    }
  });

  it("should handle axios error during app installation", async () => {
    sandbox.stub(fs, "readFile").callsFake(async () => {
      return archivedFile;
    });
    sandbox.stub(fs, "pathExists").resolves(true);

    const axiosError = {
      response: {
        data: { error: "installation failed" },
      },
      isAxiosError: true,
    };
    sandbox.stub(GraphClient.prototype, "InstallAppToChannelAsync").throws(axiosError);
    sandbox.stub(GraphClient.prototype, "GetAppInstallationForTeam").resolves([]);
    sandbox.stub(axios, "isAxiosError").returns(true);

    const args: InstallAppArgs = {
      appPackagePath: "fake/path/app.zip",
      teamId: "fake-team-id",
      channelId: "fake-channel-id",
    };
    const outputEnvVarNames = new Map<string, string>();

    const result = await driver.install(args, mockContext, outputEnvVarNames);

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error).to.be.instanceOf(HttpClientError);
      expect(result.error.message).to.include("installation failed");
    }
  });

  it("should handle general error during app installation", async () => {
    sandbox.stub(fs, "readFile").callsFake(async () => {
      return archivedFile;
    });
    sandbox.stub(fs, "pathExists").resolves(true);

    const generalError = new Error("general error");
    sandbox.stub(GraphClient.prototype, "InstallAppToChannelAsync").throws(generalError);
    sandbox.stub(GraphClient.prototype, "GetAppInstallationForTeam").resolves([]);
    sandbox.stub(axios, "isAxiosError").returns(false);

    const args: InstallAppArgs = {
      appPackagePath: "fake/path/app.zip",
      teamId: "fake-team-id",
      channelId: "fake-channel-id",
    };
    const outputEnvVarNames = new Map<string, string>();

    const result = await driver.install(args, mockContext, outputEnvVarNames);

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.message).to.equal("general error");
    }
  });

  it("should delete existing installed app", async () => {
    sandbox.stub(fs, "readFile").callsFake(async () => {
      return archivedFile;
    });
    sandbox.stub(fs, "pathExists").resolves(true);

    sandbox.stub(GraphClient.prototype, "InstallAppToChannelAsync").resolves();
    const deleteStub = sandbox.stub(GraphClient.prototype, "DeleteInstalledApp").resolves();
    sandbox.stub(GraphClient.prototype, "GetAppInstallationForTeam").resolves([
      {
        id: "installation-id",
        teamsApp: {
          externalId: "fake-id",
          id: "fake-id",
          displayName: "test-app",
          distributionMethod: "sideloaded",
        },
      },
    ]);

    const args: InstallAppArgs = {
      appPackagePath: "fake/path/app.zip",
      teamId: "fake-team-id",
      channelId: "fake-channel-id",
    };
    const outputEnvVarNames = new Map<string, string>();

    const result = await driver.execute(args, mockContext, outputEnvVarNames);

    expect(result.result.isOk()).to.be.true;
    expect(deleteStub.calledOnce).to.be.true;
  });
});
