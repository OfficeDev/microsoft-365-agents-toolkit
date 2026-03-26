// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CLICommandOption, CLICommandArgument } from "@microsoft/teamsfx-api";

export const AddSkillOptions: CLICommandOption[] = [
  {
    name: "skill-name",
    type: "string",
    description: "Name of the skill (lowercase, hyphens only).",
    required: true,
  },
  {
    name: "skill-description",
    type: "string",
    description: "Description of what the skill does.",
    required: true,
  },
  {
    name: "skill-expose-to-copilot",
    type: "string",
    description: "Whether to expose the skill to mainline M365 Copilot (yes/no).",
    required: false,
    default: "no",
  },
  {
    name: "skill-from",
    type: "string",
    description: "Path to an existing skill directory within appPackage.",
    required: false,
  },
  {
    name: "manifest-file",
    questionName: "manifest-path",
    type: "string",
    shortName: "t",
    description: "Specifies the app manifest file path.",
    required: true,
    default: "./appPackage/manifest.json",
  },
];
export const AddSkillArguments: CLICommandArgument[] = [];
