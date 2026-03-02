// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CLICommandOption } from "@microsoft/teamsfx-api";

export const playgroundOption: CLICommandOption = {
  name: "playground",
  questionName: "include-playground",
  description: "include playground configuration files.",
  type: "boolean",
  required: true,
  default: true,
};

export const localDebugOption: CLICommandOption = {
  name: "local",
  questionName: "include-local",
  description: "include local debug configuration files.",
  type: "boolean",
  required: true,
  default: true,
};

export const remoteDeployOption: CLICommandOption = {
  name: "remote",
  questionName: "include-remote",
  description: "include remote deploy configuration files.",
  type: "boolean",
  required: true,
  default: false,
};

export const programmingLanguageOption: CLICommandOption = {
  name: "language",
  questionName: "programming-language",
  description: "specify the programming language.",
  type: "string",
  required: true,
  default: "typescript",
  choices: ["typescript", "python"],
};
