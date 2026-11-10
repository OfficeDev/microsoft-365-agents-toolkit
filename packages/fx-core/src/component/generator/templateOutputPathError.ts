// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { BaseComponentInnerError } from "../error/componentError";
import { errorSource } from "./constant";

export class TemplateOutputPathError extends BaseComponentInnerError {
  constructor(entryName: string) {
    super(
      errorSource,
      "SystemError",
      "TemplateOutputPathError",
      "error.generator.TemplateOutputPathError",
      [entryName]
    );
  }
}
