// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SystemError, SystemErrorOptions } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../common/localizeUtils";
import { ErrorCategory } from "./types";
import { matchDnsError } from "./common";

export class TeamsGraphAPIFailedSystemError extends SystemError {
  constructor(e: any, correlationId: string, apiName: string, extraData: string) {
    const displayMessage = matchDnsError(e.message);
    const errorOptions: SystemErrorOptions = {
      source: "TeamsGraphClient",
      error: e,
      message: `Teams Graph API failed: [${e.name}] ${e.message as string} (api: ${apiName}, correlationId: ${correlationId}) ${extraData}`,
      displayMessage: displayMessage || getLocalizedString("error.teamsGraph.apiFailed"),
      categories: [ErrorCategory.Unhandled, apiName],
    };
    super(errorOptions);
  }
}
