// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs } from "@microsoft/teamsfx-api";

export interface AddSkillInputs extends Inputs {
  "skill-name"?: string;
  "skill-description"?: string;
  "skill-from"?: string;
  "skill-source-type"?: string;
  "skill-from-zip-file"?: string;
  "manifest-path"?: string;
}
