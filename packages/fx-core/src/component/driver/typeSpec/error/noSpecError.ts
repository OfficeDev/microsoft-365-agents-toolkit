// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../../../../common/localizeUtils";

const errorCode = "NoSpecError";
const messageKey = "driver.typespec.error.noSpecError";

export class NoSpecError extends UserError {
  constructor(actionName: string) {
    super({
      source: actionName,
      name: errorCode,
      message: getDefaultString(messageKey),
      displayMessage: getLocalizedString(messageKey),
    });
  }
}
