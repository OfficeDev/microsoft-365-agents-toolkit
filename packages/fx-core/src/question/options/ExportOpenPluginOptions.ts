// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CLICommandArgument, CLICommandOption } from "@microsoft/teamsfx-api";

export const ExportOpenPluginOptions: CLICommandOption[] = [
  {
    name: "path",
    type: "string",
    shortName: "p",
    description: "Path to the ATK project folder (containing appPackage/manifest.json) to export.",
    required: true,
  },
  {
    name: "output",
    type: "string",
    shortName: "o",
    description:
      "Destination Open Plugin directory. Defaults to ./<plugin-name>-openplugin in the current working directory.",
  },
  {
    name: "manifest-kind",
    type: "string",
    description:
      "Where to write plugin.json: 'open-plugin' (.plugin/plugin.json, default), 'claude-plugin' (.claude-plugin/plugin.json), or 'cursor-plugin' (.cursor-plugin/plugin.json).",
    default: "open-plugin",
    choices: ["open-plugin", "claude-plugin", "cursor-plugin"],
  },
];

export const ExportOpenPluginArguments: CLICommandArgument[] = [];
