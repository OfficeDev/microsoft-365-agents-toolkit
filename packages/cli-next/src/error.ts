// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError, SystemError } from "@microsoft/teamsfx-core-next";

const source = "CLI";

export class MissingRequiredOptionError extends UserError {
  constructor(commandName: string, optionName: string) {
    super({
      source,
      name: "MissingRequiredOptionError",
      message: `Missing required option '--${optionName}' for command '${commandName}'.`,
    });
  }
}

export class MissingRequiredArgumentError extends UserError {
  constructor(commandName: string, argName: string) {
    super({
      source,
      name: "MissingRequiredArgumentError",
      message: `Missing required argument '<${argName}>' for command '${commandName}'.`,
    });
  }
}

export class InvalidChoiceError extends UserError {
  constructor(optionName: string, value: string, choices: string[]) {
    super({
      source,
      name: "InvalidChoiceError",
      message: `Invalid value '${value}' for option '--${optionName}'. Valid choices: ${choices.join(", ")}`,
    });
  }
}

export class UnknownCommandError extends UserError {
  constructor(commandName: string, suggestion?: string) {
    const msg = suggestion
      ? `Unknown command '${commandName}'. Did you mean '${suggestion}'?`
      : `Unknown command '${commandName}'.`;
    super({
      source,
      name: "UnknownCommandError",
      message: msg,
    });
  }
}

export class CLISystemError extends SystemError {
  constructor(message: string, innerError?: Error) {
    super({
      source,
      name: "CLISystemError",
      message,
      error: innerError,
    });
  }
}
