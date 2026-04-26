// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../../../../common/localizeUtils";

const errorCode = "TypeSpecCompileError";
const messageKey = "driver.typespec.error.compileError";

export class TypeSpecCompileError extends UserError {
  constructor(actionName: string, output: string) {
    super({
      source: actionName,
      name: errorCode,
      message: getDefaultString(messageKey, output),
      displayMessage: getLocalizedString(messageKey, output),
    });
  }
}
