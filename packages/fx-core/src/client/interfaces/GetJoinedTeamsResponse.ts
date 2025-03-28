// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export type GetJoinedTeamsResponse = TeamsInfo[];

export interface TeamsInfo {
  id: string;
  displayName: string;
  description: string;
}
