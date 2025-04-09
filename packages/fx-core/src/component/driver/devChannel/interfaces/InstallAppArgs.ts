// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface InstallAppArgs {
  /**
   * Zipped app package path
   */
  appPackagePath: string;

  teamId: string;
  channelId: string;
}
