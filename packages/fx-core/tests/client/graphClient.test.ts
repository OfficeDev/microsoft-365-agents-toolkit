// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import { createSandbox } from "sinon";
import { RetryHandler } from "../../src/client/graphClient";
import { GraphClient } from "../../src/client/graphClient";
import { ok, err, SystemError, SensitivityLabel, signedIn } from "@microsoft/teamsfx-api";
import axios from "axios";
import "mocha";
import { MockedM365Provider, MockTools } from "../core/utils";
import * as globalState from "../../src/common/globalState";
import { setTools, TOOLS } from "../../src/common/globalVars";

describe("GraphAPIClient Test", () => {
  const sandbox = createSandbox();
  const token = "fakeToken";

  beforeEach(() => {
    sandbox.stub(RetryHandler, "RETRIES").value(1);
  });

  afterEach(() => {
    sandbox.restore();
  });

  describe("RetryHandler", () => {
    it("Happy path", async () => {
      const fn = sandbox.stub().resolves("success");
      const result = await RetryHandler.Retry(fn);
      expect(result).to.equal("success");
      expect(fn.calledOnce).to.be.true;
    });

    it("Retry on error and succeed", async () => {
      const fn = sandbox.stub();
      fn.onFirstCall().rejects(new Error("Failed"));
      fn.onSecondCall().resolves("success");

      // Set RETRIES to 2 for this test
      sandbox.stub(RetryHandler, "RETRIES").value(2);

      const result = await RetryHandler.Retry(fn);
      expect(result).to.equal("success");
      expect(fn.calledTwice).to.be.true;
    });

    it("Fail after all retries", async () => {
      const error = new Error("Failed");
      const fn = sandbox.stub().rejects(error);

      try {
        await RetryHandler.Retry(fn);
        expect.fail("Should have thrown error");
      } catch (e) {
        expect(e).to.equal(error);
      }

      expect(fn.calledOnce).to.be.true;
    });
  });

  describe("listSensitivityLabels", () => {
    const tokenProvider = new MockedM365Provider();
    setTools(new MockTools());
    it("Happy path", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      sandbox.stub(tokenProvider, "getStatus").resolves(undefined);
      const response = {
        data: {
          value: [
            {
              id: "label1",
              displayName: "General",
              name: "General Label",
              description: "General Label Description",
            },
            {
              id: "label2",
              displayName: "Confidential",
              name: "Confidential Label",
              description: "Confidential Label Description",
            },
          ],
        },
      };

      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      sandbox.stub(RetryHandler, "Retry").resolves(response);

      const graphAPIClient = new GraphClient(tokenProvider);
      const result = await graphAPIClient.listSensitivityLabels(token);

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.length).to.equal(2);
        expect(result.value[0].id).to.equal("label1");
        expect(result.value[0].displayName).to.equal("General");
        expect(result.value[1].id).to.equal("label2");
        expect(result.value[1].displayName).to.equal("Confidential");
      }
    });

    it("Return error for empty response", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      sandbox.stub(tokenProvider, "getStatus").resolves(undefined);

      const response = {};
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      sandbox.stub(RetryHandler, "Retry").resolves(response);

      const graphAPIClient = new GraphClient(tokenProvider);
      const result = await graphAPIClient.listSensitivityLabels(token);

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("listSensitivityLabelsError");
      }
    });

    it("Return error for empty data", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      sandbox.stub(tokenProvider, "getStatus").resolves(undefined);
      const response = { data: {} };
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      sandbox.stub(RetryHandler, "Retry").resolves(response);

      const graphAPIClient = new GraphClient(tokenProvider);
      const result = await graphAPIClient.listSensitivityLabels(token);

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("listSensitivityLabelsError");
      }
    });

    it("API failure", async () => {
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);
      sandbox.stub(tokenProvider, "getStatus").resolves(undefined);
      const error = new Error("API failed");
      sandbox.stub(fakeAxiosInstance, "get").rejects(error);
      sandbox.stub(RetryHandler, "Retry").rejects(error);

      const graphAPIClient = new GraphClient(tokenProvider);
      const result = await graphAPIClient.listSensitivityLabels(token);

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("listSensitivityLabelsError");
        expect(result.error.message).to.include("API failed");
      }
    });

    it("Should use cache when useCache is true and cache is valid", async () => {
      sandbox.stub(tokenProvider, "getStatus").resolves(
        ok({
          accountInfo: { unique_name: "name", tid: "123" },
          status: signedIn,
          token: "token",
        } as any)
      );
      const graphAPIClient = new GraphClient(tokenProvider);
      const labels = [
        {
          id: "label1",
          displayName: "General",
          name: "General Label",
          description: "General Label Description",
        },
      ];
      const cacheValue = {
        labels: labels,
        unixTimestamp: Date.now() - 1000 * 60 * 60, // 1 hour ago
      };

      sandbox.stub(globalState, "globalStateGet").resolves(cacheValue);
      const result = await graphAPIClient.listSensitivityLabels(token, true);

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.deep.equal(labels);
      }
    });

    it("Should not use cache when cache is expired", async () => {
      sandbox.stub(tokenProvider, "getStatus").resolves(
        ok({
          accountInfo: { unique_name: "name", tid: "123" },
          status: signedIn,
          token: "token",
        } as any)
      );
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: {
          value: [
            {
              id: "newLabel",
              displayName: "New Label",
              name: "New Label",
              description: "New Label Description",
            },
          ],
        },
      };

      const oldCache = {
        labels: [{ id: "oldLabel" }],
        unixTimestamp: Date.now() - 1000 * 60 * 60 * 25, // 25 hours ago
      };

      sandbox.stub(globalState, "globalStateGet").resolves(oldCache);
      sandbox.stub(globalState, "globalStateUpdate").resolves();
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      sandbox.stub(RetryHandler, "Retry").resolves(response);

      const graphAPIClient = new GraphClient(tokenProvider);
      const result = await graphAPIClient.listSensitivityLabels(token, true);

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.deep.equal(response.data.value);
      }
    });

    it("Should update cache after API call", async () => {
      sandbox.stub(tokenProvider, "getStatus").resolves(
        ok({
          accountInfo: { unique_name: "name", tid: "123" },
          status: signedIn,
          token: "token",
        } as any)
      );
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: {
          value: [
            {
              id: "label1",
              displayName: "General",
              name: "General Label",
              description: "General Label Description",
            },
          ],
        },
      };

      let updatedCache: any;
      sandbox.stub(globalState, "globalStateUpdate").callsFake(async (key: string, value: any) => {
        updatedCache = value;
      });

      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      sandbox.stub(RetryHandler, "Retry").resolves(response);

      const graphAPIClient = new GraphClient(tokenProvider);
      const result = await graphAPIClient.listSensitivityLabels(token, false);

      expect(result.isOk()).to.be.true;
      expect(updatedCache).to.not.be.undefined;
      expect(updatedCache.labels).to.deep.equal(response.data.value);
      expect(updatedCache.unixTimestamp).to.be.closeTo(Date.now(), 1000);
    });

    it("Should not use cache when useCache is false", async () => {
      sandbox.stub(tokenProvider, "getStatus").resolves(
        ok({
          accountInfo: { unique_name: "name", tid: "123" },
          status: signedIn,
          token: "token",
        } as any)
      );
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: {
          value: [
            {
              id: "newLabel",
              displayName: "New Label",
              name: "New Label",
              description: "New Label Description",
            },
          ],
        },
      };

      const cache = {
        labels: [{ id: "oldLabel" }],
        unixTimestamp: Date.now(),
      };

      sandbox.stub(globalState, "globalStateGet").resolves(cache);
      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      sandbox.stub(RetryHandler, "Retry").resolves(response);

      const graphAPIClient = new GraphClient(tokenProvider);
      const result = await graphAPIClient.listSensitivityLabels(token, false);

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value).to.deep.equal(response.data.value);
      }
    });

    it("Should handle response with undefined or missing label properties", async () => {
      sandbox.stub(tokenProvider, "getStatus").resolves(
        ok({
          accountInfo: { unique_name: "name", tid: "123" },
          status: signedIn,
          token: "token",
        } as any)
      );
      sandbox.stub(globalState, "globalStateGet").resolves(undefined);
      const fakeAxiosInstance = axios.create();
      sandbox.stub(axios, "create").returns(fakeAxiosInstance);

      const response = {
        data: {
          value: [
            {
              // No properties defined
            },
            {
              id: undefined,
              name: undefined,
              description: undefined,
              displayName: undefined,
            },
            {
              id: "label1",
              // Missing some properties
              displayName: "Test Label",
            },
            undefined,
          ],
        },
      };

      sandbox.stub(fakeAxiosInstance, "get").resolves(response);
      sandbox.stub(RetryHandler, "Retry").resolves(response);

      const graphAPIClient = new GraphClient(tokenProvider);
      const result = await graphAPIClient.listSensitivityLabels(token, true);

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.length).to.equal(4);
        expect(result.value[0].id).to.be.undefined;
        expect(result.value[0].name).to.be.undefined;
        expect(result.value[1].id).to.be.undefined;
        expect(result.value[1].displayName).to.be.undefined;
        expect(result.value[2].id).to.equal("label1");
        expect(result.value[2].displayName).to.equal("Test Label");
        expect(result.value[2].name).to.be.undefined;
      }
    });
  });

  describe("getGeneralSentivityLabel", () => {
    const tokenProvider = new MockedM365Provider();
    it("Happy path", async () => {
      const graphAPIClient = new GraphClient(tokenProvider);

      const labels: SensitivityLabel[] = [
        {
          id: "general-id",
          displayName: "General",
          name: "General Label",
          description: "General Label Description",
        },
        {
          id: "confidential-id",
          displayName: "Confidential",
          name: "Confidential Label",
          description: "Confidential Label Description",
        },
      ];

      sandbox.stub(graphAPIClient, "listSensitivityLabels").resolves(ok(labels));

      const result = await graphAPIClient.getGeneralSentivityLabel(token);

      expect(result.isOk()).to.be.true;
      if (result.isOk()) {
        expect(result.value.id).to.equal("general-id");
      }
    });

    it("No General label found", async () => {
      const graphAPIClient = new GraphClient(tokenProvider);

      const labels: SensitivityLabel[] = [
        {
          id: "confidential-id",
          displayName: "Confidential",
          name: "Confidential Label",
          description: "Confidential Label Description",
        },
      ];

      sandbox.stub(graphAPIClient, "listSensitivityLabels").resolves(ok(labels));

      const result = await graphAPIClient.getGeneralSentivityLabel(token);

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("getGeneralSentivityLabelError");
      }
    });

    it("General label has no ID", async () => {
      const graphAPIClient = new GraphClient(tokenProvider);

      const labels: SensitivityLabel[] = [
        {
          displayName: "General",
          name: "General Label",
          description: "General Label Description",
        },
        {
          id: "confidential-id",
          displayName: "Confidential",
          name: "Confidential Label",
          description: "Confidential Label Description",
        },
      ];

      sandbox.stub(graphAPIClient, "listSensitivityLabels").resolves(ok(labels));

      const result = await graphAPIClient.getGeneralSentivityLabel(token);

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error.name).to.equal("getGeneralSentivityLabelError");
      }
    });

    it("listSensitivityLabels returns error", async () => {
      const graphAPIClient = new GraphClient(tokenProvider);

      const fakeError = {
        name: "listSensitivityLabelsError",
        message: "API failed",
        source: "GraphAPI",
      };

      sandbox.stub(graphAPIClient, "listSensitivityLabels").resolves({
        isErr: () => true,
        isOk: () => false,
        error: fakeError,
        value: undefined,
      } as any);

      const result = await graphAPIClient.getGeneralSentivityLabel(token);

      expect(result.isErr()).to.be.true;
      if (result.isErr()) {
        expect(result.error).to.equal(fakeError);
      }
    });
  });

  describe("getCurrentUserInfo", () => {
    const tokenProvider = new MockedM365Provider();
    const graphClient = new GraphClient(tokenProvider);

    beforeEach(() => {
      sandbox.restore();
    });

    it("Should return empty strings when user is not logged in", async () => {
      sandbox.stub(tokenProvider, "getStatus").resolves(undefined);
      
      const result = await graphClient.getCurrentUserInfo();
      
      expect(result).to.deep.equal(["", ""]);
    });

    it("Should return empty strings when login status is error", async () => {
      sandbox.stub(tokenProvider, "getStatus").resolves(err(new SystemError("source", "name", "Failed to get status")));
      
      const result = await graphClient.getCurrentUserInfo();
      
      expect(result).to.deep.equal(["", ""]);
    });

    it("Should return empty strings when user is not signed in", async () => {
      sandbox.stub(tokenProvider, "getStatus").resolves(
        ok({
          status: "SignedOut",
          token: "token",
        } as any)
      );
      
      const result = await graphClient.getCurrentUserInfo();
      
      expect(result).to.deep.equal(["", ""]);
    });

    it("Should return empty strings when token is not available", async () => {
      sandbox.stub(tokenProvider, "getStatus").resolves(
        ok({
          status: signedIn,
          token: undefined,
          accountInfo: {
            unique_name: "test@example.com",
            tid: "test-tenant-id"
          }
        } as any)
      );
      
      const result = await graphClient.getCurrentUserInfo();
      
      expect(result).to.deep.equal(["", ""]);
    });

    it("Should return empty strings when accountInfo values are not strings", async () => {
      sandbox.stub(tokenProvider, "getStatus").resolves(
        ok({
          status: signedIn,
          token: "token",
          accountInfo: {
            unique_name: 123,
            tid: true
          }
        } as any)
      );
      
      const result = await graphClient.getCurrentUserInfo();
      
      expect(result).to.deep.equal(["", ""]);
    });

    it("Should return values when all required info is available", async () => {
      sandbox.stub(tokenProvider, "getStatus").resolves(
        ok({
          status: signedIn,
          token: "token",
          accountInfo: {
            unique_name: "test@example.com",
            tid: "test-tenant-id"
          }
        } as any)
      );
      
      const result = await graphClient.getCurrentUserInfo();
      
      expect(result).to.deep.equal(["test@example.com", "test-tenant-id"]);
    });

    it("Should return empty strings when accountInfo is missing", async () => {
      sandbox.stub(tokenProvider, "getStatus").resolves(
        ok({
          status: signedIn,
          token: "token"
        } as any)
      );
      
      const result = await graphClient.getCurrentUserInfo();
      
      expect(result).to.deep.equal(["", ""]);
    });

    it("Should handle undefined values in accountInfo", async () => {
      sandbox.stub(tokenProvider, "getStatus").resolves(
        ok({
          status: signedIn,
          token: "token",
          accountInfo: {
            unique_name: undefined,
            tid: undefined
          }
        } as any)
      );
      
      const result = await graphClient.getCurrentUserInfo();
      
      expect(result).to.deep.equal(["", ""]);
    });
  });
});

describe("Sandbox related APIs", () => {
  const tokenProvider = new MockedM365Provider();
  const graphClient = new GraphClient(tokenProvider);
  const sandbox = createSandbox();
  const fakeAxiosInstance = axios.create();

  beforeEach(() => {
    sandbox.stub(axios, "create").returns(fakeAxiosInstance);
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("GetJoinedSandboxedTeamsAsync should return joined sandboxed teams", async () => {
    const mockResponse = {
      data: {
        value: [
          { id: "team1", displayName: "Team 1", description: "Description 1" },
          { id: "team2", displayName: "Team 2", description: "Description 2" },
        ],
      },
    };
    sandbox.stub(fakeAxiosInstance, "get").resolves(mockResponse);

    const result = await graphClient.GetJoinedSandboxedTeamsAsync();
    expect(result).equal(mockResponse.data.value);
  });

  it("GetChannelDeeplinkAsync should return channel deeplink", async () => {
    const teamId = "fake-team-id";
    const channelId = "fake-channel-id";
    const mockResponse = {
      data: {
        webUrl: "https://teams.microsoft.com/l/channel/fake-channel",
      },
    };
    sandbox.stub(fakeAxiosInstance, "get").resolves(mockResponse);

    const result = await graphClient.GetChannelDeeplinkAsync(teamId, channelId);
    expect(result).to.equal("https://teams.microsoft.com/l/channel/fake-channel");
  });

  it("InstallAppToChannelAsync should install app successfully", async () => {
    const teamId = "fake-team-id";
    const channelId = "fake-channel-id";
    const file = Buffer.from("fake-file-content");
    sandbox.stub(fakeAxiosInstance, "post").resolves({ status: 200 });

    let error: any = undefined;
    try {
      await graphClient.InstallAppToChannelAsync(teamId, channelId, file);
    } catch (e) {
      error = e;
    }
    expect(error).to.be.undefined;
  });

  it("CreateTeamAndChannelAsync should create team and channel successfully", async () => {
    const teamName = "Test Team";
    const description = "Test Description";
    const defaultChannelName = "General";
    const locationHeader =
      "/teams('dbd8de4f-5d47-48da-87f1-594bed003375')/operations('3a6fdce1-c261-48bc-89de-1cfef658c0d5')";
    const teamId = "dbd8de4f-5d47-48da-87f1-594bed003375";

    sandbox.stub(fakeAxiosInstance, "post").resolves({ headers: { location: locationHeader } });

    const statusStub = sandbox.stub(fakeAxiosInstance, "get");
    statusStub.onFirstCall().resolves({ data: { status: "inProgress" } });
    statusStub.onSecondCall().resolves({ data: { status: "succeeded" } });

    const channelsResponse = {
      data: {
        value: [{ id: "fake-channel-id", displayName: defaultChannelName }],
      },
    };
    statusStub.onThirdCall().resolves(channelsResponse);

    const result = await graphClient.CreateTeamAndChannelAsync(
      teamName,
      description,
      defaultChannelName
    );

    expect(result).to.deep.equal({
      teamId: teamId,
      channelId: "fake-channel-id",
    });
  });

  it("CreateChannelAsync should create a channel successfully", async () => {
    const teamId = "fake-team-id";
    const channelName = "Test Channel";
    const description = "Test Channel Description";

    const mockResponse = {
      data: {
        id: "fake-channel-id",
        webUrl: "https://teams.microsoft.com/l/channel/fake-channel-id",
      },
    };
    sandbox.stub(fakeAxiosInstance, "post").resolves(mockResponse);

    const result = await graphClient.CreateChannelAsync(teamId, channelName, description);

    expect(result).to.deep.equal({
      id: "fake-channel-id",
      webUrl: "https://teams.microsoft.com/l/channel/fake-channel-id",
    });
  });

  it("GetChannelsInTeamAsync should return channels in a team", async () => {
    const teamId = "fake-team-id";
    const mockResponse = {
      data: {
        value: [
          { id: "channel1", webUrl: "https://teams.microsoft.com/l/channel/channel1" },
          { id: "channel2", webUrl: "https://teams.microsoft.com/l/channel/channel2" },
        ],
      },
    };
    sandbox.stub(fakeAxiosInstance, "get").resolves(mockResponse);

    const result = await graphClient.GetChannelsInTeamAsync(teamId);

    expect(result).to.deep.equal(mockResponse.data.value);
  });

  it("GetTeamsAppSettingsAsync should return teams app settings", async () => {
    const mockResponse = {
      data: {
        sandboxingConfiguration: {
          sensitivityLabelUsedToIdentifySandboxedContainers: "0fcfd0ff-1cda-407e-bc2b-a350307bd1d5",
        },
      },
    };
    sandbox.stub(fakeAxiosInstance, "get").resolves(mockResponse);

    const result = await graphClient.GetTeamsAppSettingsAsync();

    expect(result).to.deep.equal(mockResponse.data);
  });

  it("GetAppInstallationForTeam should return installed apps successfully", async () => {
    const teamId = "fake-team-id";
    const mockResponse = {
      data: {
        value: [
          {
            id: "installation-id-1",
            teamsApp: {
              externalId: "app-external-id-1",
              displayName: "App 1",
            },
          },
          {
            id: "installation-id-2",
            teamsApp: {
              externalId: "app-external-id-2",
              displayName: "App 2",
            },
          },
        ],
      },
    };

    sandbox.stub(fakeAxiosInstance, "get").resolves(mockResponse);

    const result = await graphClient.GetAppInstallationForTeam(teamId);

    expect(result).to.deep.equal(mockResponse.data.value);
  });

  it("DeleteInstalledApp should delete app installation successfully", async () => {
    const teamId = "fake-team-id";
    const installationId = "fake-installation-id";

    sandbox.stub(fakeAxiosInstance, "delete").resolves({ status: 204 });

    let error: any = undefined;
    try {
      await graphClient.DeleteInstalledApp(teamId, installationId);
    } catch (e) {
      error = e;
    }

    expect(error).to.be.undefined;
  });
});

describe("Sandbox related APIs - failed token", () => {
  const tokenProvider = new MockedM365Provider();
  const graphClient = new GraphClient(tokenProvider);
  const sandbox = createSandbox();

  beforeEach(() => {
    sandbox
      .stub(tokenProvider, "getAccessToken")
      .resolves(err(new SystemError("GraphClient", "TokenError", "Failed to get access token")));
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("GetJoinedSandboxedTeamsAsync failed to get access token", async () => {
    let error: any = undefined;
    try {
      await graphClient.GetJoinedSandboxedTeamsAsync();
    } catch (e) {
      error = e;
    }
    expect(error).to.not.be.undefined;
    expect(error.name).to.equal("TokenError");
    expect(error.message).to.equal("Failed to get access token");
  });

  it("GetChannelDeeplinkAsync failed to get access token", async () => {
    let error: any = undefined;
    try {
      await graphClient.GetChannelDeeplinkAsync("fake-team-id", "fake-channel-id");
    } catch (e) {
      error = e;
    }
    expect(error).to.not.be.undefined;
    expect(error.name).to.equal("TokenError");
  });

  it("InstallAppToChannelAsync failed to get access token", async () => {
    let error: any = undefined;
    try {
      await graphClient.InstallAppToChannelAsync(
        "fake-team-id",
        "fake-channel-id",
        Buffer.from("fake-content")
      );
    } catch (e) {
      error = e;
    }
    expect(error).to.not.be.undefined;
    expect(error.name).to.equal("TokenError");
  });

  it("CreateTeamAndChannelAsync failed to get access token", async () => {
    let error: any = undefined;
    try {
      await graphClient.CreateTeamAndChannelAsync("Test Team", "Test Description", "General");
    } catch (e) {
      error = e;
    }
    expect(error).to.not.be.undefined;
    expect(error.name).to.equal("TokenError");
  });

  it("CreateChannelAsync failed to get access token", async () => {
    let error: any = undefined;
    try {
      await graphClient.CreateChannelAsync("fake-team-id", "Test Channel", "Test Description");
    } catch (e) {
      error = e;
    }
    expect(error).to.not.be.undefined;
    expect(error.name).to.equal("TokenError");
  });

  it("GetTeamsAppSettingsAsync failed to get access token", async () => {
    let error: any = undefined;
    try {
      await graphClient.GetTeamsAppSettingsAsync();
    } catch (e) {
      error = e;
    }
    expect(error).to.not.be.undefined;
    expect(error.name).to.equal("TokenError");
  });

  it("GetChannelsInTeamAsync failed to get access token", async () => {
    let error: any = undefined;
    try {
      await graphClient.GetChannelsInTeamAsync("fake-team-id");
    } catch (e) {
      error = e;
    }
    expect(error).to.not.be.undefined;
    expect(error.name).to.equal("TokenError");
  });

  it("GetAppInstallationForTeam failed to get access token", async () => {
    let error: any = undefined;
    try {
      await graphClient.GetAppInstallationForTeam("fake-team-id");
    } catch (e) {
      error = e;
    }
    expect(error).to.not.be.undefined;
    expect(error.name).to.equal("TokenError");
  });

  it("DeleteInstalledApp failed to get access token", async () => {
    let error: any = undefined;
    try {
      await graphClient.DeleteInstalledApp("fake-team-id", "installation-id");
    } catch (e) {
      error = e;
    }
    expect(error).to.not.be.undefined;
    expect(error.name).to.equal("TokenError");
  });
});
