// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs } from "@microsoft/teamsfx-api";

export interface ExportOpenPluginInputs extends Inputs {
  /** Path to the ATK project folder (containing appPackage/manifest.json). */
  path: string;
  /** Destination Open Plugin directory. */
  output?: string;
  /** Manifest kind to emit. Defaults to 'open-plugin'. */
  "manifest-kind"?: "open-plugin" | "claude-plugin" | "cursor-plugin";
}
