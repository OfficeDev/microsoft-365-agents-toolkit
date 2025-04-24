// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";

export class InstallAppOutsideSandboxError extends UserError {
  constructor(actionName: string) {
    super({
      source: actionName,
      name: "InstallAppOutsideSandbox",
      message: getDefaultString("error.installApp.outsideSandbox"),
      displayMessage: getLocalizedString("error.installApp.outsideSandbox"),
    });
  }
}
