// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../../../../common/localizeUtils";
import { constants } from "../utility/constants";

const errorCode = "ClientSecretNotAllowed";
const messageKey = "driver.aadApp.error.credentialTypeNotAllowedAsPerAppPolicy";

export class ClientSecretNotAllowedError extends UserError {
  constructor(actionName: string, isMicrosoftUser = false) {
    super({
      source: actionName,
      name: errorCode,
      message: getDefaultString(messageKey),
      displayMessage: getLocalizedString(messageKey),
      helpLink: isMicrosoftUser ? constants.sniHelpLink : constants.defaultHelpLink,
    });
  }
}
