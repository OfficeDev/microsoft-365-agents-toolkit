// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface GetAppInstallationResponse {
  id: string; // Installation ID
  teamsApp: {
    id: string;
    externalId: string;
    displayName: string;
    distributionMethod: string;
  };
}
