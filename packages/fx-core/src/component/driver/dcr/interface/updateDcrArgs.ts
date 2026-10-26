// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  OauthRegistrationAppType,
  OauthRegistrationTargetAudience,
} from "../../teamsApp/interfaces/OauthRegistration";

export interface UpdateDcrArgs {
  configurationId: string;
  name?: string;
  appId?: string;
  mcpResourceUrl?: string;
  wellKnownAuthorizationServer?: string;
  resource?: string;
  targetUrlsShouldStartWith?: string[];
  applicableToApps?: OauthRegistrationAppType;
  targetAudience?: OauthRegistrationTargetAudience;
  includePersonalMicrosoftAccounts?: boolean;
}
