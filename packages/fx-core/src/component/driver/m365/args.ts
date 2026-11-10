// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { AppScope } from "../../m365/packageService";
import { getLocalizedString } from "../../../common/localizeUtils";
import { createArgsContract } from "../util/args";

export const acquireActionName = "teamsApp/extendToM365";

export interface AcquireArgs {
  appPackagePath: string;
  scope?: `${AppScope}` | Lowercase<AppScope>;
}

export const acquireArgs = createArgsContract<AcquireArgs>({
  type: "object",
  description: `${getLocalizedString("driver.m365.acquire.description")} ${getLocalizedString("core.actions.writesWarning")}`,
  additionalProperties: false,
  required: ["appPackagePath"],
  properties: {
    appPackagePath: {
      type: "string",
      minLength: 1,
      description: getLocalizedString("core.actions.acquire.appPackagePath"),
    },
    scope: {
      type: "string",
      description: getLocalizedString("core.actions.acquire.scope"),
      nullable: true,
      enum: ["Personal", "Shared", "Tenant", "personal", "shared", "tenant"],
    },
  },
});
