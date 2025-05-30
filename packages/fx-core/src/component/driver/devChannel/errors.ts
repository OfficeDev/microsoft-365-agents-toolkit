// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";

export class InstallAppOutsideSandboxError extends UserError {
  constructor(actionName: string, helplink?: string) {
    super({
      source: actionName,
      // eslint-disable-next-line no-secrets/no-secrets
      name: "InstallAppOutsideSandbox",
      message: getDefaultString("error.installApp.outsideSandbox"),
      displayMessage: getLocalizedString("error.installApp.outsideSandbox"),
      helpLink: helplink,
    });
  }
}
