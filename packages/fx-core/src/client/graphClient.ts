// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";
import { M365TokenProvider, LogProvider } from "@microsoft/teamsfx-api";
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
import { waitSeconds } from "../common/utils";
import { getLocalizedString } from "../common/localizeUtils";
import { GetAppInstallationResponse } from "./interfaces/GetAppInstallationResponse";

export class GraphClient {
  private readonly baseUrl: string =
    process.env.GRAPH_ENDPOINT ?? "https://graph.microsoft.com/beta";
  private readonly tokenProvider: M365TokenProvider;
  private readonly logProvider: LogProvider | undefined;

  constructor(tokenProvider: M365TokenProvider, logProvider?: LogProvider) {
    this.tokenProvider = tokenProvider;
    this.logProvider = logProvider;
  }

  private createRequesterWithToken(token: string): AxiosInstance {
    const instance = WrappedAxiosClient.create({
      baseURL: this.baseUrl,
    });
    instance.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    return instance;
  }

  /**
   * Get sandboxing configuration of team app settings.
   * @returns
   */
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
    return <GetTeamsAppSettingsResponse>response.data;
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

  /**
   * Get weburl of a channel.
   * @param teamId
   * @param channelId
   * @returns
   */
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
    const data = <GetChannelResponse>response.data;
    return data.webUrl;
  }

  /**
   * Install Teams app package into a channel.
   * @param teamId
   * @param channelId
   * @param file Teams app package zip file
   */
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

    await requester.post(`/teams/${teamId}/installedApps?targetChannelId=${channelId}`, file, {
      headers: { "Content-Type": "application/zip" },
    });
  }

  /**
   * Get installed apps in a team.
   * @param teamId
   * @returns An array of installed apps, the externalId is the Teams app id.
   */
  public async GetAppInstallationForTeam(teamId: string): Promise<GetAppInstallationResponse[]> {
    const tokenResponse = await this.tokenProvider.getAccessToken({
      scopes: GraphTeamsInstallAppScopes,
    });
    if (tokenResponse.isErr()) {
      throw tokenResponse.error;
    }
    const requester = this.createRequesterWithToken(tokenResponse.value);

    const response = await requester.get(`/teams/${teamId}/installedApps?$expand=teamsapp`);
    return <GetAppInstallationResponse[]>response.data.value;
  }

  public async DeleteInstalledApp(teamId: string, installationId: string): Promise<void> {
    const tokenResponse = await this.tokenProvider.getAccessToken({
      scopes: GraphTeamsInstallAppScopes,
    });
    if (tokenResponse.isErr()) {
      throw tokenResponse.error;
    }
    const requester = this.createRequesterWithToken(tokenResponse.value);

    await requester.delete(`/teams/${teamId}/installedApps/${installationId}`);
  }

  /**
   * Create a sandboxed team and a channel.
   * @param teamName Team name
   * @param description Team description
   * @param defaultChannelName Channel name
   * @returns
   */
  @hooks([ErrorContextMW({ source: "Teams", component: "GraphClient" })])
  public async CreateTeamAndChannelAsync(
    teamName: string,
    description: string,
    defaultChannelName: string
  ): Promise<CreateTeamAndChannelResponse> {
    const LocationRegex = /teams\('([0-9a-fA-F-]{36})'\)\/operations\('([0-9a-fA-F-]{36})'/;
    const tokenResponse = await this.tokenProvider.getAccessToken({
      scopes: [
        ...GraphTeamsTeamCreateScopes,
        ...GraphTeamsTeamReadScopes,
        ...GraphTeamsChannelReadScopes,
      ],
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
    const location = response.headers.location;

    if (location) {
      // this.logProvider?.info(`Location header: ${location}`);
      const match = location.match(LocationRegex);
      if (match) {
        const teamId = match[1];
        let status = await requester.get(location);

        // Query team creation status, until it's succeeded
        while (status.data.status !== "succeeded") {
          await waitSeconds(5);
          const message = getLocalizedString("driver.devChannel.status", status.data.status);
          this.logProvider?.info(message);
          status = await requester.get(location);
        }

        // Get Channel ID
        const channels = await this.GetChannelsInTeamAsync(teamId);
        const channel = channels.find((channel) => channel.displayName === defaultChannelName);
        if (channel) {
          const channelId = channel.id;
          return {
            teamId: teamId,
            channelId: channelId,
          };
        } else {
          throw new Error(`Failed to find channel with name: ${defaultChannelName}`);
        }
      } else {
        throw new Error("Failed to parse location header.");
      }
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

  /**
   * List channels in a team
   * @param teamId Team ID
   * @returns A list of channels, with id and webUrl
   */
  @hooks([ErrorContextMW({ source: "Teams", component: "GraphClient" })])
  public async GetChannelsInTeamAsync(teamId: string): Promise<GetChannelResponse[]> {
    const tokenResponse = await this.tokenProvider.getAccessToken({
      scopes: GraphTeamsChannelReadScopes,
    });
    if (tokenResponse.isErr()) {
      throw tokenResponse.error;
    }
    const requester = this.createRequesterWithToken(tokenResponse.value);

    const response = await requester.get(`/teams/${teamId}/channels`);
    return <GetChannelResponse[]>response.data.value;
  }
}
