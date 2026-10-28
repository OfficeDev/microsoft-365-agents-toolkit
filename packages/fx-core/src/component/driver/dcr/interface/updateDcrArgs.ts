// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  OauthRegistrationAppType,
  OauthRegistrationTargetAudience,
} from "../../teamsApp/interfaces/OauthRegistration";

export interface UpdateDcrArgs {
  configurationId: string;
  appId?: string;
  resource?: string;
  applicableToApps?: OauthRegistrationAppType;
  targetAudience?: OauthRegistrationTargetAudience;
  includePersonalMicrosoftAccounts?: boolean;
}
