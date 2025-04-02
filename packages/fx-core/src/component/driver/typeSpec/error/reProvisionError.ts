// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../../../../common/localizeUtils";

const errorCode = "ReProvisionError";
const messageKey = "driver.typeSpec.compile.reprovision";

export class ReProvisionError extends UserError {
  constructor(actionName: string, configFile: string) {
    super({
      source: actionName,
      name: errorCode,
      message: getDefaultString(messageKey, configFile),
      displayMessage: getLocalizedString(messageKey, configFile),
    });
  }
}
