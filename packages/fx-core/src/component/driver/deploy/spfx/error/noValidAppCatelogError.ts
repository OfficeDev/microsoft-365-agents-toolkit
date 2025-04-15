// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../../../../../common/localizeUtils";
import { MetadataV3, MetadataV4 } from "../../../../../common/versionMetadata";
import { Constants } from "../utility/constants";

const errorCode = "NoValidAppCatelog";
const messageKey = "driver.spfx.error.noValidAppCatelog";

export class NoValidAppCatelog extends UserError {
  constructor() {
    const ymlFile = `${MetadataV4.configFile} or ${MetadataV3.configFile}`;
    super({
      source: Constants.DeployDriverName,
      name: errorCode,
      message: getDefaultString(messageKey, ymlFile),
      displayMessage: getLocalizedString(messageKey, ymlFile),
    });
  }
}
