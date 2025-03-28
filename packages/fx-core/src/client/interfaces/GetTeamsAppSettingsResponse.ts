// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface GetTeamsAppSettingsResponse {
  sandboxingConfiguration: {
    isSideloadingEnabled: boolean;
    sensitivityLabelUsedToIdentifySandboxedContainers: string;
  };
}
