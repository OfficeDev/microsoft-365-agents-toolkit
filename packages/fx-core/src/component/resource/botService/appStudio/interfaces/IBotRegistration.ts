// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Qianhao Dong <qidon@microsoft.com>
 */
export enum BotChannelType {
  MicrosoftTeams = "msteams",
  M365Extensions = "m365extensions",
}

export interface IBotRegistration {
  botId?: string;
  name: string;
  description: string;
  iconUrl: string;
  messagingEndpoint: string;
  callingEndpoint: string;
  configuredChannels?: BotChannelType[];
  isSingleTenant?: boolean;
}
