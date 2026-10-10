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

export function mapPersonalMicrosoftAccountsOption(
  input: unknown,
  actionName: string
): Result<OauthRegistrationSupportedAccountTypes | undefined, FxError> {
  if (input === undefined) {
    return ok(undefined);
  }
  if (typeof input !== "boolean") {
    return err(new InvalidActionInputError(actionName, ["includePersonalMicrosoftAccounts"]));
  }

  return ok(input ? "Enterprise, Consumer" : "Enterprise");
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
