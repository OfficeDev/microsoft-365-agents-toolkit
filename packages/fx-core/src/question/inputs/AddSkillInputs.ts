// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs } from "@microsoft/teamsfx-api";

export interface AddSkillInputs extends Inputs {
  "skill-name"?: string;
  "skill-description"?: string;
  "skill-expose-to-copilot"?: string;
  "skill-from"?: string;
  "manifest-path"?: string;
}
