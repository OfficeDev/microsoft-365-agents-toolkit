// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../../../../../common/localizeUtils";
import { MetadataV4 } from "../../../../../common/versionMetadata";
import { Constants } from "../utility/constants";

const errorCode = "NoValidAppCatelog";
const messageKey = "driver.spfx.error.noValidAppCatelog";

export class NoValidAppCatelog extends UserError {
  constructor() {
    super({
      source: Constants.DeployDriverName,
      name: errorCode,
      message: getDefaultString(messageKey, MetadataV4.configFile),
      displayMessage: getLocalizedString(messageKey, MetadataV4.configFile),
    });
  }
}
