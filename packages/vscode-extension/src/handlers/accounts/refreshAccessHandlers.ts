// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, FxError, ok } from "@microsoft/teamsfx-api";
import { AppStudioScopes } from "@microsoft/teamsfx-core";
import accountTreeViewProviderInstance from "../../treeview/account/accountTreeViewProvider";
import M365TokenInstance from "../../commonlib/m365Login";

export async function refreshSideloadingCallback(args?: any[]): Promise<Result<null, FxError>> {
  const status = await M365TokenInstance.getStatus({ scopes: AppStudioScopes });
  if (status.isOk() && status.value.token !== undefined) {
    accountTreeViewProviderInstance.m365AccountNode.updateChecks(status.value.token, true, false);
  }

  return ok(null);
}

export async function refreshCopilotCallback(args?: any[]): Promise<Result<null, FxError>> {
  const status = await M365TokenInstance.getStatus({ scopes: AppStudioScopes });
  if (status.isOk() && status.value.token !== undefined) {
    accountTreeViewProviderInstance.m365AccountNode.updateChecks(status.value.token, false, true);
  }

  return ok(null);
}
