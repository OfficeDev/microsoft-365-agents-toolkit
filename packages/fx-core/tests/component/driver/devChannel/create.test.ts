// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import { createSandbox } from "sinon";
import axios from "axios";
import "mocha";
import mockedEnv from "mocked-env";
import { CreateDevChannelDriver } from "../../../../src/component/driver/devChannel/create";
import { GraphClient } from "../../../../src/client/graphClient";
import { MockedM365Provider, MockLogProvider } from "../../../core/utils";
import { WrapDriverContext } from "../../../../src/component/driver/util/wrapUtil";

describe("CreateDevChannelDriver", () => {
  const sandbox = createSandbox();
  const mockTokenProvider = new MockedM365Provider();
  const mockContext: WrapDriverContext = {
    m365TokenProvider: mockTokenProvider,
    logProvider: new MockLogProvider(),
    addSummary: sandbox.stub(),
    summaries: [],
  } as unknown as WrapDriverContext;

  const driver = new CreateDevChannelDriver();

  beforeEach(() => {});

  afterEach(() => {
    sandbox.restore();
  });

  it("should skip creation if team and channel already exist", async () => {
    const args = {
      teamName: "Test Team",
      channelName: "Test channel",
      teamDescription: "Test Description",
    };
    const outputEnvVarNames = new Map([
      ["teamId", "TEAM_ID"],
      ["channelId", "CHANNEL_ID"],
      ["channelWebUrl", "CHANNEL_WEB_URL"],
    ]);

    const restore = mockedEnv({
      TEAM_ID: "existing-team-id",
      CHANNEL_ID: "exisitng-channel-id",
    });

    sandbox.stub(GraphClient.prototype, "GetChannelDeeplinkAsync").resolves("fake-deeplink");
    const result = await driver.execute(args, mockContext, outputEnvVarNames);

    expect(result.result.isOk()).to.be.true;
    if (result.result.isOk()) {
      expect(result.result.value.size).to.equal(3);
    }
    restore();
  });

  it("should create team and channel successfully", async () => {
    const args = {
      teamName: "Test Team",
      channelName: "Channel name",
      teamDescription: "Test Description",
    };
    const outputEnvVarNames = new Map([
      ["teamId", "TEAM_ID"],
      ["channelId", "CHANNEL_ID"],
    ]);

    const mockGraphResponse = {
      teamId: "fake-team-id",
      channelId: "fake-channel-id",
    };

    sandbox.stub(GraphClient.prototype, "CreateTeamAndChannelAsync").resolves(mockGraphResponse);
    sandbox.stub(GraphClient.prototype, "GetChannelDeeplinkAsync").resolves("fake-deeplink");

    const result = await driver.create(args, mockContext, outputEnvVarNames);

    expect(result.isOk()).to.be.true;
    if (result.isOk()) {
      expect(result.value.get("TEAM_ID")).to.equal("fake-team-id");
      expect(result.value.get("CHANNEL_ID")).to.equal("fake-channel-id");
    }
  });

  it("should handle error when CreateTeamAndChannelAsync fails", async () => {
    const args = {
      teamName: "Test Team",
      channelName: "Channel name",
      teamDescription: "Test Description",
    };
    const outputEnvVarNames = new Map([
      ["teamId", "TEAM_ID"],
      ["channelId", "CHANNEL_ID"],
    ]);

    const fakeAxiosInstance = axios.create();
    sandbox.stub(axios, "create").returns(fakeAxiosInstance);
    sandbox.stub(axios, "isAxiosError").returns(true);
    const error = {
      response: {
        status: 409,
      },
    };
    sandbox.stub(fakeAxiosInstance, "post").throws(error);

    const result = await driver.create(args, mockContext, outputEnvVarNames);

    expect(result.isErr()).to.be.true;
  });
});
