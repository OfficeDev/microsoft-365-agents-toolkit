// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result } from "neverthrow";
import { FxError } from "./error";

export type OptionValue = string | boolean | string[] | undefined;
export type CLIOptionType = "boolean" | "string" | "array";

export interface CLICommand {
  name: string;
  aliases?: string[];
  fullName?: string;
  version?: string;
  description: string;
  arguments?: CLICommandArgument[];
  options?: CLICommandOption[];
  sortOptions?: boolean;
  commands?: CLICommand[];
  sortCommands?: boolean;
  examples?: CLIExample[];
  handler?: (ctx: CLIContext) => Promise<Result<undefined, FxError>> | Result<undefined, FxError>;
  telemetry?: { event: string };
  header?: string;
  footer?: string;
  hidden?: boolean;
  defaultInteractiveOption?: boolean;
  reservedOptionNamesInInteractiveMode?: string[];
}

export interface CLIFoundCommand extends CLICommand {
  fullName: string;
}

export interface CLIContext {
  command: CLIFoundCommand;
  optionValues: Record<string, OptionValue>;
  globalOptionValues: Record<string, OptionValue>;
  argumentValues: OptionValue[];
  telemetryProperties: Record<string, string>;
}

interface CLICommandOptionBase {
  name: string;
  questionName?: string;
  description: string;
  shortName?: string;
  type: CLIOptionType;
  required?: boolean;
  hidden?: boolean;
}

export interface CLIBooleanOption extends CLICommandOptionBase {
  type: "boolean";
  default?: boolean;
  value?: boolean;
}

export interface CLIStringOption extends CLICommandOptionBase {
  type: "string";
  default?: string;
  value?: string;
  choices?: string[];
  skipValidation?: boolean;
  choiceListCommand?: string;
}

export interface CLIArrayOption extends CLICommandOptionBase {
  type: "array";
  default?: string[];
  value?: string[];
  choices?: string[];
  skipValidation?: boolean;
  choiceListCommand?: string;
}

export type CLICommandOption = CLIBooleanOption | CLIStringOption | CLIArrayOption;

export type CLICommandArgument = CLICommandOption;

export interface CLIExample {
  command: string;
  description: string;
}
