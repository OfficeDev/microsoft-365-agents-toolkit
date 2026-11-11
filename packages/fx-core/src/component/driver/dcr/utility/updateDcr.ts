// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { URL } from "url";
import { InvalidActionInputError } from "../../../../error/common";
import {
  mapPersonalMicrosoftAccountsOption,
  validateSupportedAccountTypesCloud,
} from "../../oauth/utility/supportedAccountTypes";
import { DcrRegistration, DcrRegistrationUpdate } from "../../teamsApp/interfaces/DcrRegistration";
import {
  OauthRegistrationAppType,
  OauthRegistrationTargetAudience,
} from "../../teamsApp/interfaces/OauthRegistration";
import { UpdateDcrArgs } from "../interface/updateDcrArgs";

const actionName = "dcr/update";
const stringFields: readonly ("appId" | "resource")[] = ["appId", "resource"];
const allowedFields: ReadonlySet<string> = new Set([
  "configurationId",
  ...stringFields,
  "applicableToApps",
  "targetAudience",
  "includePersonalMicrosoftAccounts",
]);

export interface DcrUpdatePlan {
  changes: DcrRegistrationUpdate;
  // Binding removal is intent; its provisional wire representation belongs to the driver.
  clearAppBinding: boolean;
}

function isRecord(input: unknown): input is Record<string, unknown> {
  return typeof input === "object" && input !== null && !Array.isArray(input);
}

export function validateUpdateDcrArgs(input: unknown): Result<UpdateDcrArgs, FxError> {
  if (!isRecord(input)) {
    return err(new InvalidActionInputError(actionName, ["with"]));
  }

  const unexpected = Object.keys(input).filter((field) => !allowedFields.has(field));
  if (unexpected.length > 0) {
    return err(new InvalidActionInputError(actionName, unexpected));
  }

  if (typeof input.configurationId !== "string" || input.configurationId.trim().length === 0) {
    return err(new InvalidActionInputError(actionName, ["configurationId"]));
  }
  const args: UpdateDcrArgs = { configurationId: input.configurationId };

  for (const field of stringFields) {
    const value = input[field];
    if (value !== undefined) {
      if (typeof value !== "string" || value.length === 0) {
        return err(new InvalidActionInputError(actionName, [field]));
      }
      args[field] = value;
    }
  }

  if (
    args.resource !== undefined &&
    (!URL.canParse(args.resource) || args.resource.includes("#"))
  ) {
    return err(new InvalidActionInputError(actionName, ["resource"]));
  }

  const appMode = input.applicableToApps;
  if (appMode !== undefined) {
    if (
      appMode !== OauthRegistrationAppType.AnyApp &&
      appMode !== OauthRegistrationAppType.SpecificApp
    ) {
      return err(new InvalidActionInputError(actionName, ["applicableToApps"]));
    }
    args.applicableToApps = appMode;
  }

  const tenantScope = input.targetAudience;
  if (tenantScope !== undefined) {
    if (
      tenantScope !== OauthRegistrationTargetAudience.HomeTenant &&
      tenantScope !== OauthRegistrationTargetAudience.AnyTenant
    ) {
      return err(new InvalidActionInputError(actionName, ["targetAudience"]));
    }
    args.targetAudience = tenantScope;
  }

  const accountTypes = mapPersonalMicrosoftAccountsOption(
    input.includePersonalMicrosoftAccounts,
    actionName
  );
  if (accountTypes.isErr()) {
    return err(accountTypes.error);
  }
  const cloudValidation = validateSupportedAccountTypesCloud(accountTypes.value, actionName);
  if (cloudValidation.isErr()) {
    return err(cloudValidation.error);
  }
  if (typeof input.includePersonalMicrosoftAccounts === "boolean") {
    args.includePersonalMicrosoftAccounts = input.includePersonalMicrosoftAccounts;
  }

  return ok(args);
}

export function createDcrUpdatePlan(
  args: UpdateDcrArgs,
  current: Readonly<DcrRegistration>
): Result<DcrUpdatePlan, FxError> {
  const accountTypes = mapPersonalMicrosoftAccountsOption(
    args.includePersonalMicrosoftAccounts,
    actionName
  );
  if (accountTypes.isErr()) {
    return err(accountTypes.error);
  }

  const changes: DcrRegistrationUpdate = {};
  let clearAppBinding = false;

  if (args.resource !== undefined && args.resource !== current.resource) {
    changes.resource = args.resource;
  }

  if (args.appId !== undefined || args.applicableToApps !== undefined) {
    const effectiveMode = args.applicableToApps ?? current.applicableToApps;
    if (effectiveMode === OauthRegistrationAppType.SpecificApp) {
      const appId = args.appId ?? current.m365AppId;
      if (!appId) {
        return err(new InvalidActionInputError(actionName, ["appId"]));
      }
      if (appId !== current.m365AppId) {
        changes.m365AppId = appId;
      }
    } else if (args.applicableToApps === OauthRegistrationAppType.AnyApp) {
      clearAppBinding = current.applicableToApps === OauthRegistrationAppType.SpecificApp;
    }
    if (args.applicableToApps !== undefined && args.applicableToApps !== current.applicableToApps) {
      changes.applicableToApps = args.applicableToApps;
    }
  }

  if (args.targetAudience !== undefined && args.targetAudience !== current.targetAudience) {
    changes.targetAudience = args.targetAudience;
  }
  if (accountTypes.value !== undefined && accountTypes.value !== current.supportedAccountTypes) {
    changes.supportedAccountTypes = accountTypes.value;
  }

  return ok({ changes, clearAppBinding });
}
