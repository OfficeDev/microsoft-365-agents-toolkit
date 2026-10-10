// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";

export class OauthSupportedAccountTypesUnsupportedCloudError extends UserError {
  constructor(actionName: string, message: string) {
    super(actionName, "OauthSupportedAccountTypesUnsupportedCloud", message, message);
  }
}
