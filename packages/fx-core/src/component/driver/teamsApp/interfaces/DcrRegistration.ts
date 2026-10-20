// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  OauthRegistrationAppType,
  OauthRegistrationSupportedAccountTypes,
  OauthRegistrationTargetAudience,
} from "./OauthRegistration";

export interface DcrRegistration {
  m365AppId: string;
  clientName: string;
  applicableToApps: OauthRegistrationAppType;
  targetAudience: OauthRegistrationTargetAudience;
  targetUrlsShouldStartWith: string[];
  mcpResourceUrl?: string;
  wellKnownAuthorizationServer?: string;
  resource?: string;
  supportedAccountTypes?: OauthRegistrationSupportedAccountTypes;
  // TODO: add this part back after TDP update
  // manageableByUsers: [
  //   {
  //     userId: string;
  //     accessType: OauthRegistrationUserAccessType;
  //   },
  // ],
}
