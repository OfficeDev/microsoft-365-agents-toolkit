// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CLICommandOption } from "@microsoft/teamsfx-api";
import { commands } from "../resource";

export const ProjectFolderOption: CLICommandOption = {
  name: "folder",
  questionName: "projectPath",
  shortName: "f",
  description: "Project folder.",
  type: "string",
  required: true,
  default: "./",
};
export const TeamsAppManifestFileOption: CLICommandOption = {
  name: "manifest-file",
  type: "string",
  description: "Specifies the Microsoft 365 App manifest file path.",
  default: "./appPackage/manifest.json",
};
export const EntraAppManifestFileOption: CLICommandOption = {
  name: "manifest-file",
  questionName: "manifest-file-path",
  type: "string",
  description: "Specifies the Microsoft Entra app manifest file path.",
  default: "./aad.manifest.json",
};
export const TeamsAppPackageOption: CLICommandOption = {
  name: "package-file",
  type: "string",
  description: "Specifies the zipped Microsoft 365 App package file path.",
};
export const TeamsAppOuputPackageOption: CLICommandOption = {
  name: "output-package-file",
  type: "string",
  description: "Specifies the output zipped Microsoft 365 App package file path.",
  default: "./appPackage/build/appPackage.${env}.zip",
};
export const TeamsAppOutputFolderOption: CLICommandOption = {
  name: "output-folder",
  type: "string",
  description: "Specifies the output folder containing the manifest(s).",
  default: "./appPackage/build",
};
export const EnvOption: CLICommandOption = {
  name: "env",
  type: "string",
  description:
    "Specifies the environment name for the project scaffolded by Microsoft 365 Agents Toolkit.",
};
export const IgnoreLoadEnvOption: CLICommandOption = {
  name: "ignore-env-file",
  type: "boolean",
  description: "Whether to skip loading .env file when --env is not specified.",
};
export const EnvFileOption: CLICommandOption = {
  name: "env-file",
  type: "string",
  description:
    "Specifies the .env file that defines the variables to replace in the Microsoft 365 App manifest template file.",
};
export const IgnoreKeysOption: CLICommandOption = {
  name: "ignore-keys",
  type: "array",
  description: "Specifies the keys to ignore in the .env file.",
};

export const ListFormatOption: CLICommandOption = {
  name: "format",
  shortName: "f",
  description: commands["list.templates"].options.format,
  type: "string",
  choices: ["table", "json"],
  default: "table",
  required: true,
};

export const ShowDescriptionOption: CLICommandOption = {
  name: "description",
  shortName: "d",
  description: "Whether to show description in the result.",
  type: "boolean",
  default: false,
  required: true,
};

export const ConfigFilePathOption: CLICommandOption = {
  type: "string",
  name: "config-file-path",
  shortName: "c",
  description: "Specifies the path of the configuration yaml file.",
};

export const ValidateMethodOption: CLICommandOption = {
  type: "string",
  name: "validate-method",
  shortName: "m",
  choices: ["validation-rules", "test-cases"],
  description: "Specifies validation method",
};
