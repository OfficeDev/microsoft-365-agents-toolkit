// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  CLIArrayOption,
  CLICommandArgument,
  CLICommandOption,
  CLIStringOption,
  UserError,
} from "@microsoft/teamsfx-api";
import * as constants from "./constants";
import { errors as strings } from "./resource";
import * as util from "util";
import { helper } from "./commands/helper";

export class MissingRequiredOptionError extends UserError {
  constructor(command: string, option: string | CLICommandOption) {
    super({
      source: constants.cliSource,
      message: util.format(
        strings["error.MissingRequiredOptionError"],
        command,
        typeof option === "string" ? option : option.name,
        typeof option === "string" ? option : helper.formatOptionName(option, false)
      ),
    });
  }
}
export class MissingRequiredArgumentError extends UserError {
  constructor(command: string, argument: string | CLICommandArgument | CLICommandOption) {
    super({
      source: constants.cliSource,
      message: util.format(
        strings["error.MissingRequiredArgumentError"],
        command,
        typeof argument === "string" ? argument : argument.name
      ),
    });
  }
}

export class InvalidChoiceError extends UserError {
  constructor(command: string, value: string, option: CLIStringOption | CLIArrayOption) {
    super({
      source: constants.cliSource,
      message: util.format(
        strings["error.InvalidChoiceError"],
        command,
        value,
        helper.formatOptionName(option, false),
        option.choices!.join(", ")
      ),
    });
  }
}

export class ArgumentConflictError extends UserError {
  constructor(command: string, name1: string, name2: string) {
    super({
      source: constants.cliSource,
      message: util.format(strings["error.ArgumentConflictError"], command, name1, name2),
    });
  }
}

export class UnknownOptionError extends UserError {
  constructor(command: string, name: string) {
    super({
      source: constants.cliSource,
      message: util.format(strings["error.UnknownOptionError"], command, name),
    });
  }
}

export class UnknownArgumentError extends UserError {
  constructor(command: string, name: string) {
    super({
      source: constants.cliSource,
      message: util.format(strings["error.UnknownArgumentError"], command, name),
    });
  }
}

export class UnknownCommandError extends UserError {
  constructor(token: string, fullName: string, mostSimilar?: string) {
    super({
      source: constants.cliSource,
      message: `'${token}' is misspelled or not recognized by Microsoft 365 Toolkit CLI.${
        mostSimilar ? " Did you mean '" + fullName + " " + mostSimilar + "'?" : ""
      } Use '${fullName} -h' for more command information.`,
    });
  }
}
