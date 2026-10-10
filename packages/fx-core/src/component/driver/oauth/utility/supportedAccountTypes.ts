// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import {
  getSovereignCloudEnvironment,
  SovereignCloudEnvironment,
} from "../../../../common/accountUtils";
import { getLocalizedString } from "../../../../common/localizeUtils";
import { InvalidActionInputError } from "../../../../error/common";
import { OauthRegistrationSupportedAccountTypes } from "../../teamsApp/interfaces/OauthRegistration";
import { OauthSupportedAccountTypesUnsupportedCloudError } from "../error/oauthSupportedAccountTypesUnsupportedCloud";

const ENTERPRISE = "Enterprise";
const CONSUMER = "Consumer";

export function normalizeSupportedAccountTypes(
  input: unknown,
  actionName: string
): Result<OauthRegistrationSupportedAccountTypes | undefined, FxError> {
  if (input === undefined) {
    return ok(undefined);
  }
  if (typeof input !== "string") {
    return err(new InvalidActionInputError(actionName, ["supportedAccountTypes"]));
  }

  const values = input.split(",").map((value) => value.trim());
  const uniqueValues = new Set(values);
  if (
    values.some((value) => value.length === 0) ||
    uniqueValues.size !== values.length ||
    !uniqueValues.has(ENTERPRISE) ||
    values.some((value) => value !== ENTERPRISE && value !== CONSUMER)
  ) {
    return err(new InvalidActionInputError(actionName, ["supportedAccountTypes"]));
  }

  return ok(uniqueValues.has(CONSUMER) ? "Enterprise, Consumer" : "Enterprise");
}

export function validateSupportedAccountTypesCloud(
  supportedAccountTypes: OauthRegistrationSupportedAccountTypes | undefined,
  actionName: string
): Result<void, FxError> {
  if (
    supportedAccountTypes !== undefined &&
    getSovereignCloudEnvironment() !== SovereignCloudEnvironment.Public
  ) {
    return err(
      new OauthSupportedAccountTypesUnsupportedCloudError(
        actionName,
        getLocalizedString("driver.oauth.error.supportedAccountTypesUnsupportedCloud")
      )
    );
  }
  return ok(undefined);
}
