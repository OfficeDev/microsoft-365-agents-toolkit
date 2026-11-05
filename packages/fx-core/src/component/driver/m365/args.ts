// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { AppScope } from "../../m365/packageService";
import { createArgsContract } from "../util/args";

export const acquireActionName = "teamsApp/extendToM365";

export interface AcquireArgs {
  appPackagePath: string;
  scope?: `${AppScope}` | Lowercase<AppScope>;
}

export const acquireArgs = createArgsContract<AcquireArgs>({
  type: "object",
  additionalProperties: false,
  required: ["appPackagePath"],
  properties: {
    appPackagePath: { type: "string", minLength: 1 },
    scope: {
      type: "string",
      nullable: true,
      enum: ["Personal", "Shared", "Tenant", "personal", "shared", "tenant"],
    },
  },
});
