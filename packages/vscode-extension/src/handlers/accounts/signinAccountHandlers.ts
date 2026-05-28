// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result, FxError, ok, err } from "@microsoft/teamsfx-api";
import { AppStudioScopes, isUserCancelError } from "@microsoft/teamsfx-core";
import { tools } from "../../globalVariables";
import { ExtTelemetry } from "../../telemetry/extTelemetry";
import { AccountType, TelemetryEvent, TelemetryProperty } from "../../telemetry/extTelemetryEvents";
import { AzureAccountNode } from "../../treeview/account/azureNode";
import { AccountItemStatus } from "../../treeview/account/common";
import { M365AccountNode } from "../../treeview/account/m365Node";
import { getTriggerFromProperty } from "../../utils/telemetryUtils";
import envTreeProviderInstance from "../../treeview/environmentTreeViewProvider";
import azureAccountManager from "../../commonlib/azureLogin";

export async function signinM365Callback(...args: unknown[]): Promise<Result<null, FxError>> {
  let node: M365AccountNode | undefined;
  if (args && args.length > 1) {
    node = args[1] as M365AccountNode;
    if (node && node.status === AccountItemStatus.SignedIn) {
      return ok(null);
    }
  }

  const triggerFrom = getTriggerFromProperty(args);
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.LoginClick, {
    [TelemetryProperty.AccountType]: AccountType.M365,
    ...triggerFrom,
  });

  const tokenRes = await tools.tokenProvider.m365TokenProvider.getJsonObject({
    scopes: AppStudioScopes,
    showDialog: true,
  });
  const token = tokenRes.isOk() ? tokenRes.value : undefined;
  if (token !== undefined && node) {
    await node.setSignedIn(
      (token as any).upn ?? (token as any).unique_name ?? "",
      (token as any).tid ?? ""
    );
  }

  await envTreeProviderInstance.reloadEnvironments();
  return ok(null);
}

export async function signinAzureCallback(...args: unknown[]): Promise<Result<null, FxError>> {
  let node: AzureAccountNode | undefined;
  if (args && args.length > 1) {
    node = args[1] as AzureAccountNode;
    if (node && node.status === AccountItemStatus.SignedIn) {
      return ok(null);
    }
  }

  if (azureAccountManager.getAccountInfo() === undefined) {
    // make sure user has not logged in
    const triggerFrom = getTriggerFromProperty(args);
    ExtTelemetry.sendTelemetryEvent(TelemetryEvent.LoginClick, {
      [TelemetryProperty.AccountType]: AccountType.Azure,
      ...triggerFrom,
    });
  }
  try {
    await azureAccountManager.getIdentityCredentialAsync(true);
  } catch (error) {
    if (!isUserCancelError(error)) {
      return err(error);
    }
  }
  return ok(null);
}
