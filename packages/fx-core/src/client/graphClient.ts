// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";
import { M365TokenProvider } from "@microsoft/teamsfx-api";
import { AxiosInstance } from "axios";
import { ErrorContextMW } from "../common/globalVars";
import { GetTeamsAppSettingsResponse } from "./interfaces/GetTeamsAppSettingsResponse";
import {
  GraphTeamsAppSettingsReadScopes,
  GraphTeamsChannelCreateScopes,
  GraphTeamsChannelReadScopes,
  GraphTeamsInstallAppScopes,
  GraphTeamsTeamCreateScopes,
  GraphTeamsTeamReadScopes,
} from "../common/constants";
import { GetJoinedTeamsResponse } from "./interfaces/GetJoinedTeamsResponse";
import { GetChannelResponse } from "./interfaces/GetChannelResponse";
import { WrappedAxiosClient } from "../common/wrappedAxiosClient";
import { CreateChannelResponse } from "./interfaces/CreateChannelResponse";
import { CreateTeamAndChannelResponse } from "./interfaces/CreateTeamAndChannelResponse";

export class GraphClient {
  private readonly baseUrl: string = "https://graph.microsoft.com/beta";
  private readonly tokenProvider: M365TokenProvider;

  constructor(tokenProvider: M365TokenProvider) {
    this.tokenProvider = tokenProvider;
  }

  private createRequesterWithToken(token: string): AxiosInstance {
    const instance = WrappedAxiosClient.create({
      baseURL: this.baseUrl,
    });
    instance.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    return instance;
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "GraphClient" })])
  public async GetTeamsAppSettingsAsync(): Promise<GetTeamsAppSettingsResponse> {
    const tokenResponse = await this.tokenProvider.getAccessToken({
      scopes: GraphTeamsAppSettingsReadScopes,
    });
    if (tokenResponse.isErr()) {
      throw tokenResponse.error;
    }
    const requester = this.createRequesterWithToken(tokenResponse.value);

    const response = await requester.get(
      `/teamwork/teamsAppSettings?$select=sandboxingConfiguration`
    );
    return <GetTeamsAppSettingsResponse>response.data.value;
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "GraphClient" })])
  public async GetJoinedSandboxedTeamsAsync(): Promise<GetJoinedTeamsResponse> {
    const tokenResponse = await this.tokenProvider.getAccessToken({
      scopes: GraphTeamsTeamReadScopes,
    });
    if (tokenResponse.isErr()) {
      throw tokenResponse.error;
    }
    const requester = this.createRequesterWithToken(tokenResponse.value);

    const response = await requester.get(`/me/joinedTeams?isSandboxedTeam=true`);
    return <GetJoinedTeamsResponse>response.data.value;
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "GraphClient" })])
  public async GetChannelDeeplinkAsync(teamId: string, channelId: string): Promise<string> {
    const tokenResponse = await this.tokenProvider.getAccessToken({
      scopes: GraphTeamsChannelReadScopes,
    });
    if (tokenResponse.isErr()) {
      throw tokenResponse.error;
    }
    const requester = this.createRequesterWithToken(tokenResponse.value);

    const response = await requester.get(`/teams/${teamId}/channels/${channelId}`);
    const data = <GetChannelResponse>response.data.value;
    return data.webUrl;
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "GraphClient" })])
  public async InstallAppToChannelAsync(
    teamId: string,
    channelId: string,
    file: Buffer
  ): Promise<void> {
    const tokenResponse = await this.tokenProvider.getAccessToken({
      scopes: GraphTeamsInstallAppScopes,
    });
    if (tokenResponse.isErr()) {
      throw tokenResponse.error;
    }
    const requester = this.createRequesterWithToken(tokenResponse.value);

    await requester.post(`/teams/${teamId}/installApps?targetChannelId=${channelId}`, file, {
      headers: { "Content-Type": "application/zip" },
    });
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "GraphClient" })])
  public async CreateTeamAndChannelAsync(
    teamName: string,
    description: string,
    defaultChannelName: string
  ): Promise<CreateTeamAndChannelResponse> {
    const tokenResponse = await this.tokenProvider.getAccessToken({
      scopes: GraphTeamsTeamCreateScopes,
    });
    if (tokenResponse.isErr()) {
      throw tokenResponse.error;
    }
    const requester = this.createRequesterWithToken(tokenResponse.value);

    const teamData = {
      "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
      displayName: teamName,
      description: description,
      firstChannelName: defaultChannelName,
    };

    const response = await requester.post(`/teams?isSandboxedTeam=true`, teamData);
    const location = void response.headers.Location;
    if (location) {
      // TODO: Check the status of opearation completion, logging the status
      // await status = requester.get(location);
      return {
        teamId: location,
        channelId: location,
      };
    } else {
      throw new Error("Failed to create team and channel.");
    }
  }

  @hooks([ErrorContextMW({ source: "Teams", component: "GraphClient" })])
  public async CreateChannelAsync(
    teamId: string,
    channelName: string,
    description: string
  ): Promise<CreateChannelResponse> {
    const tokenResponse = await this.tokenProvider.getAccessToken({
      scopes: GraphTeamsChannelCreateScopes,
    });
    if (tokenResponse.isErr()) {
      throw tokenResponse.error;
    }
    const requester = this.createRequesterWithToken(tokenResponse.value);

    const channelData = {
      displayName: channelName,
      description: description,
      membershipType: "standard",
    };

    const response = await requester.post(`/teams/${teamId}/channels`, channelData);
    return <CreateChannelResponse>response.data;
  }
}
