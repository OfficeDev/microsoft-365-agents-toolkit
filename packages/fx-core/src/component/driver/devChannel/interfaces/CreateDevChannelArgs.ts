// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface CreateDevChannelArgs {
  /**
   * The name of the team to create.
   */
  teamName: string;
  /**
   * The name of the channel to create.
   */
  channelName: string;

  /**
   * The description of the team to create.
   */
  teamDescription: string;
}
