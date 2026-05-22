// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { OauthRegistrationAppType, OauthRegistrationTargetAudience } from "./OauthRegistration";

export interface DcrRegistration {
  m365AppId: string;
  clientName: string;
  applicableToApps: OauthRegistrationAppType;
  targetAudience: OauthRegistrationTargetAudience;
  targetUrlsShouldStartWith: string[];
  wellKnownAuthorizationServer: string;
  // TODO: add this part back after TDP update
  // manageableByUsers: [
  //   {
  //     userId: string;
  //     accessType: OauthRegistrationUserAccessType;
  //   },
  // ],
}
