// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { PublishPolicyArgs } from "../../workiq/policyPublication";
import { createArgsContract } from "../util/args";

export type { PublishPolicyArgs } from "../../workiq/policyPublication";
export { publishPolicyActionName } from "../../workiq/policyPublicationReport";

export const publishPolicyArgs = createArgsContract<PublishPolicyArgs>({
  type: "object",
  additionalProperties: false,
  required: ["titleId", "policyPath", "callerId"],
  properties: {
    titleId: { type: "string", minLength: 1, pattern: "\\S" },
    policyPath: { type: "string", minLength: 1, pattern: "\\S" },
    callerId: { type: "string", minLength: 1, pattern: "\\S" },
  },
});
