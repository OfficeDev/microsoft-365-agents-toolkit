// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SystemError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../common/localizeUtils";

export class KiotaGeneratePluginError extends SystemError {
  constructor(errorMsg: string) {
    super({
      source: "kiota",
      message: getDefaultString("error.kiota.KiotaGeneratePluginError", errorMsg),
      displayMessage: getLocalizedString("error.kiota.KiotaGeneratePluginError", errorMsg),
    });
  }
}
