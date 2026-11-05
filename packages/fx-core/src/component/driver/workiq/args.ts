// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { PublishPolicyArgs } from "../../workiq/policyPublication";
import { getLocalizedString } from "../../../common/localizeUtils";
import { createArgsContract } from "../util/args";

export type { PublishPolicyArgs } from "../../workiq/policyPublication";
export { publishPolicyActionName } from "../../workiq/policyPublicationReport";

export const publishPolicyArgs = createArgsContract<PublishPolicyArgs>({
  type: "object",
  description: `${getLocalizedString("driver.workiq.description")} ${getLocalizedString("core.actions.writesWarning")}`,
  additionalProperties: false,
  required: ["titleId", "policyPath", "callerId"],
  properties: {
    titleId: {
      type: "string",
      minLength: 1,
      pattern: "\\S",
      description: getLocalizedString("core.actions.publishPolicy.titleId"),
    },
    policyPath: {
      type: "string",
      minLength: 1,
      pattern: "\\S",
      description: getLocalizedString("core.actions.publishPolicy.policyPath"),
    },
    callerId: {
      type: "string",
      minLength: 1,
      pattern: "\\S",
      description: getLocalizedString("core.actions.publishPolicy.callerId"),
    },
  },
});
