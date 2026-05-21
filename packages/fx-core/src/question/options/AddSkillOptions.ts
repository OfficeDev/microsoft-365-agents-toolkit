// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CLICommandOption, CLICommandArgument } from "@microsoft/teamsfx-api";

export const AddSkillOptions: CLICommandOption[] = [
  {
    name: "name",
    questionName: "skill-name",
    type: "string",
    description:
      "Name of the skill (letters, numbers, and hyphens only). Required when not using --from.",
    required: false,
  },
  {
    name: "description",
    questionName: "skill-description",
    type: "string",
    description: "Description of what the skill does. Required when not using --from.",
    required: false,
  },
  {
    name: "from",
    questionName: "skill-from",
    type: "string",
    description:
      "Path to an existing skill directory within appPackage, or path to a .zip file containing a skill.",
    required: false,
  },
  {
    name: "expose-to-copilot",
    questionName: "expose-to-copilot",
    type: "boolean",
    description:
      "Expose this skill to all Copilot surfaces via the Teams manifest agentSkills property.",
    required: false,
    default: false,
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
