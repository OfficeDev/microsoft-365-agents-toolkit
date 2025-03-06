// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../../../../common/localizeUtils";
import { constants } from "../utility/constants";

const errorCode = "SignInAudienceNotAllowed";
const messageKey = "driver.aadApp.error.signInAudienceNotAllowedAsPerAppPolicy";

export class SignInAudienceNotAllowedError extends UserError {
  constructor(source: string, errorDetail?: string, isMicrosoftUser = false) {
    const helpLink = isMicrosoftUser ? constants.sniHelpLink : constants.defaultHelpLink;
    super({
      source: source,
      name: errorCode,
      message: getDefaultString(messageKey, errorDetail),
      displayMessage: getLocalizedString(messageKey, errorDetail),
      helpLink: helpLink,
    });
  }
}
