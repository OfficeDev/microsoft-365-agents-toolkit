// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError, UserErrorOptions } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../common/localizeUtils";
import { maskSecret } from "../common/stringUtils";
import { ErrorCategory } from "./types";

/**
 * Script execution timeout
 */
export class ScriptTimeoutError extends UserError {
  constructor(error: Error, script: string) {
    const maskedScript = maskSecret(script, { replace: "***" });
    const errorOptions: UserErrorOptions = {
      source: "script",
      name: "ScriptTimeoutError",
      message: getDefaultString("error.script.ScriptTimeoutError", maskedScript),
      displayMessage: getLocalizedString("error.script.ScriptTimeoutError.Notification"),
      error: error,
      categories: [ErrorCategory.External],
    };
    super(errorOptions);
  }
}

/**
 * Script execution error
 */
export class ScriptExecutionError extends UserError {
  constructor(error: Error, script: string) {
    const maskedScript = maskSecret(script, { replace: "***" });
    const maskedError = maskSecret(error.message || "", { replace: "***" });
    const maskedUserData = maskSecret(JSON.stringify(error, Object.getOwnPropertyNames(error)), {
      replace: "***",
    });
    const errorOptions: UserErrorOptions = {
      source: "script",
      name: "ScriptExecutionError",
      message: getDefaultString("error.script.ScriptExecutionError", maskedScript, maskedError),
      displayMessage: getLocalizedString(
        "error.script.ScriptExecutionError.Notification",
        maskedError
      ),
      error: error,
      categories: [ErrorCategory.External],
      userData: maskedUserData,
    };
    super(errorOptions);
  }
}
