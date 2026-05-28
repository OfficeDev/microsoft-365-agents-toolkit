// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as core from "@microsoft/teamsfx-core";
import { AzureAccountManager } from "../../../commonlib/azureLogin";
import { signedIn } from "../../../commonlib/common/constant";
import { M365Login } from "../../../commonlib/m365Login";

export async function checkCredential(): Promise<{
  m365LoggedIn: boolean;
  azureLoggedIn: boolean;
}> {
  const m365Status = await M365Login.getInstance().getStatus({ scopes: core.AppStudioScopes });
  const azureStatus = await AzureAccountManager.getInstance().getStatus();
  return {
    m365LoggedIn: m365Status.isOk() && m365Status.value.status === signedIn,
    azureLoggedIn: azureStatus.status === signedIn,
  };
}

export function getProjectMetadata(rootPath: string | undefined) {
  return core.getProjectMetadata(rootPath);
}

export function globalStateGet(key: string, defaultValue?: any) {
  return core.globalStateGet(key, defaultValue);
}

export function globalStateUpdate(key: string, value: any) {
  return core.globalStateUpdate(key, value);
}
