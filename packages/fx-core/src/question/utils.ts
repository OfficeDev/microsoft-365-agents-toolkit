// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs } from "@microsoft/teamsfx-api";

export function ensureInputs(inputs?: Inputs): Inputs {
  if (!inputs) {
    throw new Error("inputs is undefined"); // should never happen
  }
  return inputs;
}
