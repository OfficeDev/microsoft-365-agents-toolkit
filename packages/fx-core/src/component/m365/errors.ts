// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";

import { getDefaultString, getLocalizedString } from "../../common/localizeUtils";
import { M365HelpLink } from "./constants";

export class NotExtendedToM365Error extends UserError {
  constructor(source: string) {
    super({
      source: source,
      name: "NotExtendedToM365Error",
      message: getDefaultString("error.m365.NotExtendedToM365Error"),
      displayMessage: getLocalizedString("error.m365.NotExtendedToM365Error"),
      helpLink: M365HelpLink,
    });
  }
}

export interface PackageValidationFailureReason {
  errorCode: string;
  errorDetail: string;
}

/**
 * A sanitized MOS3 request failure of a prepared publication. It never carries the underlying
 * HTTP client error, which retains authorization headers and response bodies; only the HTTP
 * status is exposed. `definitive` is true only when MOS3 refused the request and did not accept it.
 */
export class PreparedPackageRequestError extends Error {
  public readonly response?: { status: number };

  constructor(
    public readonly code: string,
    public readonly definitive: boolean,
    status?: number
  ) {
    super(code);
    this.name = "PreparedPackageRequestError";
    if (status !== undefined) {
      this.response = { status };
    }
  }
}

export class PackageValidationFailedError extends UserError {
  constructor(source: string, failureReasons: PackageValidationFailureReason[]) {
    const reasons = failureReasons
      .map(({ errorCode, errorDetail }) => `${errorCode}: ${errorDetail}`)
      .join("\n");
    super({
      source,
      name: "PackageValidationFailed",
      message: getDefaultString("error.m365.packageService.validationFailed", reasons),
      displayMessage: getLocalizedString("error.m365.packageService.validationFailed", reasons),
      helpLink: M365HelpLink,
    });
  }
}
