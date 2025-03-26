// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import { createSandbox } from "sinon";
import axios from "axios";
import "mocha";
import mockedEnv from "mocked-env";
import fs from "fs-extra";
import path from "path";
import { InstallAppToChannelDriver } from "../../../../src/component/driver/devChannel/installApp";
import { GraphClient } from "../../../../src/client/graphClient";
import { MockedM365Provider, MockLogProvider } from "../../../core/utils";
import { WrapDriverContext } from "../../../../src/component/driver/util/wrapUtil";
import {
  InvalidActionInputError,
  FileNotFoundError,
  HttpClientError,
} from "../../../../src/error/common";

describe("InstallAppToChannelDriver", () => {
  const sandbox = createSandbox();
  const mockTokenProvider = new MockedM365Provider();
  const mockContext: WrapDriverContext = {
    m365TokenProvider: mockTokenProvider,
    logProvider: new MockLogProvider(),
    addSummary: sandbox.stub(),
    summaries: [],
    projectPath: "fake/project/path",
  } as unknown as WrapDriverContext;

  const driver = new InstallAppToChannelDriver();

  afterEach(() => {
    sandbox.restore();
  });

  it("should return error if teamId or channelId is missing", async () => {
    const restore = mockedEnv({
      TEAM_ID: undefined,
      CHANNEL_ID: undefined,
    });

    const args = { appPackagePath: "fake/path/app.zip" };
    const outputEnvVarNames = new Map<string, string>();

    const result = await driver.install(args, mockContext, outputEnvVarNames);

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error).to.be.instanceOf(InvalidActionInputError);
      expect(result.error.message).to.include("teamId or channelId");
    }
    restore();
  });

  it("should return error if app package file does not exist", async () => {
    const restore = mockedEnv({
      TEAM_ID: "fake-team-id",
      CHANNEL_ID: "fake-channel-id",
    });

    sandbox.stub(fs, "pathExists").resolves(false);

    const args = { appPackagePath: "fake/path/app.zip" };
    const outputEnvVarNames = new Map<string, string>();

    const result = await driver.install(args, mockContext, outputEnvVarNames);

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error).to.be.instanceOf(FileNotFoundError);
    }
    restore();
  });

  it("should install app to channel successfully", async () => {
    const restore = mockedEnv({
      TEAM_ID: "fake-team-id",
      CHANNEL_ID: "fake-channel-id",
    });

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").resolves(Buffer.from("fake-content"));
    sandbox.stub(GraphClient.prototype, "InstallAppToChannelAsync").resolves();

    const args = { appPackagePath: "fake/path/app.zip" };
    const outputEnvVarNames = new Map<string, string>();

    const result = await driver.install(args, mockContext, outputEnvVarNames);

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.size).to.equal(0);
    }
    restore();
  });

  it("should handle axios error during app installation", async () => {
    const restore = mockedEnv({
      TEAM_ID: "fake-team-id",
      CHANNEL_ID: "fake-channel-id",
    });

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").resolves(Buffer.from("fake-content"));

    const axiosError = {
      response: {
        data: { error: "installation failed" },
      },
      isAxiosError: true,
    };
    sandbox.stub(GraphClient.prototype, "InstallAppToChannelAsync").throws(axiosError);
    sandbox.stub(axios, "isAxiosError").returns(true);

    const args = { appPackagePath: "fake/path/app.zip" };
    const outputEnvVarNames = new Map<string, string>();

    const result = await driver.install(args, mockContext, outputEnvVarNames);

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error).to.be.instanceOf(HttpClientError);
      expect(result.error.message).to.include("installation failed");
    }
    restore();
  });

  it("should handle general error during app installation", async () => {
    const restore = mockedEnv({
      TEAM_ID: "fake-team-id",
      CHANNEL_ID: "fake-channel-id",
    });

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "readFile").resolves(Buffer.from("fake-content"));

    const generalError = new Error("general error");
    sandbox.stub(GraphClient.prototype, "InstallAppToChannelAsync").throws(generalError);
    sandbox.stub(axios, "isAxiosError").returns(false);

    const args = { appPackagePath: "fake/path/app.zip" };
    const outputEnvVarNames = new Map<string, string>();

    const result = await driver.install(args, mockContext, outputEnvVarNames);

    expect(result.isErr()).to.be.true;
    if (result.isErr()) {
      expect(result.error.message).to.equal("general error");
    }
    restore();
  });
});
