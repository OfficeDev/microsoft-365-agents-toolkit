// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DcrRegistration } from "../../teamsApp/interfaces/DcrRegistration";
import {
  OauthRegistrationAppType,
  OauthRegistrationSupportedAccountType,
  OauthRegistrationTargetAudience,
} from "../../teamsApp/interfaces/OauthRegistration";

export function isDcrSnapshot(input: unknown, configurationId: string): input is DcrRegistration {
  if (typeof input !== "object" || input === null || Array.isArray(input)) {
    return false;
  }
  if (
    !("clientName" in input) ||
    typeof input.clientName !== "string" ||
    !input.clientName ||
    !("m365AppId" in input) ||
    typeof input.m365AppId !== "string" ||
    !("applicableToApps" in input) ||
    (input.applicableToApps !== OauthRegistrationAppType.AnyApp &&
      input.applicableToApps !== OauthRegistrationAppType.SpecificApp) ||
    !("targetAudience" in input) ||
    (input.targetAudience !== OauthRegistrationTargetAudience.HomeTenant &&
      input.targetAudience !== OauthRegistrationTargetAudience.AnyTenant) ||
    !("targetUrlsShouldStartWith" in input) ||
    !Array.isArray(input.targetUrlsShouldStartWith) ||
    !input.targetUrlsShouldStartWith.every((value: unknown) => typeof value === "string")
  ) {
    return false;
  }
  if (!("identityProvider" in input) || input.identityProvider !== "Dynamic") {
    return false;
  }
  if ("oAuthConfigId" in input && input.oAuthConfigId !== configurationId) {
    return false;
  }
  for (const field of ["mcpResourceUrl", "wellKnownAuthorizationServer", "resource"]) {
    if (field in input && typeof Reflect.get(input, field) !== "string") {
      return false;
    }
  }
  return (
    !("supportedAccountTypes" in input) ||
    input.supportedAccountTypes === OauthRegistrationSupportedAccountType.Enterprise ||
    input.supportedAccountTypes === OauthRegistrationSupportedAccountType.EnterpriseAndConsumer
  );
}
