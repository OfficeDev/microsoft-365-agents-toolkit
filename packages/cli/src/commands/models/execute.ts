// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  CLICommand,
  CLICommandOption,
  FxError,
  Result,
  SystemError,
  UserError,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import { maskSecret, maskSecretFromEnv } from "@microsoft/teamsfx-core";
import {
  ActionName,
  ActionNames,
  getActionDefinition,
  validateActionArgs,
} from "@microsoft/teamsfx-core/actions";
import path from "path";
import { format } from "util";
import { getFxCore } from "../../activate";
import { logger } from "../../commonlib/logger";
import { cliSource } from "../../constants";
import { MissingRequiredOptionError } from "../../error";
import { commands, errors } from "../../resource";
import { ProjectFolderOptionWithoutValidation } from "../common";

const reservedOptionNames = [
  "folder",
  "f",
  "help",
  "h",
  "version",
  "v",
  "interactive",
  "i",
  "debug",
  "verbose",
  "telemetry",
];

const hiddenInteractiveOption: CLICommandOption = {
  name: "interactive",
  shortName: "i",
  type: "boolean",
  hidden: true,
  description: commands.root.options.interactive,
};

const publicResultFields = new Set([
  "titleId",
  "appId",
  "callerId",
  "operationId",
  "policySha256",
  "etag",
  "revision",
]);

function redactResultValue(key: string, value: unknown): unknown {
  if (value === undefined || typeof value === "function") return value;
  if (
    /(secret|password|token|key|authorization|credential|connection[-_]?string|cookie)/i.test(key)
  ) {
    return "<REDACTED:secret>";
  }
  if (value instanceof Map) return Object.fromEntries(value);
  if (typeof value !== "string") return value;
  // Business metadata skips heuristic classification, not explicit registered-secret masking.
  return publicResultFields.has(key) ? maskSecretFromEnv(value) : maskSecret(value);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((item) => typeof item === "string");
}

function actionOptions(schema: Record<string, unknown>): Result<CLICommandOption[], FxError> {
  const unsupported = (): Result<CLICommandOption[], FxError> =>
    err(
      new SystemError({
        source: cliSource,
        name: "UnsupportedActionSchema",
        message: errors["error.UnsupportedActionSchema"],
      })
    );
  const required = schema.required ?? [];
  const properties = schema.properties;
  if (schema.type !== "object" || !isRecord(properties) || !isStringArray(required)) {
    return unsupported();
  }
  const names = new Set(reservedOptionNames);
  const options: CLICommandOption[] = [];
  for (const [propertyName, property] of Object.entries(properties)) {
    const name = propertyName.replace(/[A-Z]/g, (letter) => `-${letter.toLowerCase()}`);
    if (!name || names.has(name) || !isRecord(property) || property.type !== "string") {
      return unsupported();
    }
    const choices = property.enum;
    if (choices !== undefined && !isStringArray(choices)) return unsupported();
    names.add(name);
    options.push({
      name,
      questionName: propertyName,
      type: "string",
      description: typeof property.description === "string" ? property.description : propertyName,
      required: required.includes(propertyName),
      choices,
    });
  }
  if (required.some((name) => !Object.hasOwn(properties, name))) return unsupported();
  return ok(options);
}

function invalidValueOption(commandOptions: CLICommandOption[]): string | undefined {
  // The shared parser drops missing values, truncates inline '=', and can requeue flag values.
  const args = process.argv.slice(2);
  const names = new Set(reservedOptionNames);
  const options = new Map<string, CLICommandOption>();
  for (const option of commandOptions) {
    for (const name of [option.name, option.shortName]) {
      if (name) {
        names.add(name);
        options.set(name, option);
      }
    }
  }
  const isFlag = (value: string | undefined) =>
    value !== undefined &&
    value.startsWith("-") &&
    names.has(value.replace(/^--?/, "").split("=")[0]);
  for (const [index, arg] of args.entries()) {
    if (!arg.startsWith("-")) continue;
    const name = arg.replace(/^--?/, "").split("=")[0];
    if (!names.has(name)) continue;
    const parts = arg.split("=");
    const inline = parts.length > 1;
    const value = inline ? parts[1] : args[index + 1];
    if (
      (inline && isFlag(value)) ||
      (options.get(name)?.type === "string" &&
        (parts.length > 2 || value === undefined || value === "" || isFlag(value)))
    ) {
      return name;
    }
  }
  return undefined;
}

function actionCommand<K extends ActionName>(name: K): CLICommand {
  const definition = getActionDefinition(name);
  const options = definition.andThen(({ argsSchema }) => actionOptions(argsSchema));
  if (options.isErr()) {
    const error = options.error;
    return {
      name,
      description: error.message,
      defaultInteractiveOption: false,
      options: [hiddenInteractiveOption],
      handler: () => err(error),
    };
  }
  const schema = definition.isOk() ? definition.value.argsSchema : undefined;
  const description = typeof schema?.description === "string" ? schema.description : name;
  return {
    name,
    description: [
      description,
      commands["execute.action"].argumentMode,
      ...(name === ActionNames.PublishPolicy
        ? [commands["execute.action"].policyPathResolution]
        : []),
    ].join("\n\n"),
    defaultInteractiveOption: false,
    options: [
      ...options.value.map((option) =>
        name === ActionNames.PublishPolicy && option.questionName === "policyPath"
          ? { ...option, description: commands["execute.action"].policyPathResolution }
          : option
      ),
      { ...ProjectFolderOptionWithoutValidation, hidden: true },
      hiddenInteractiveOption,
    ],
    handler: async (context) => {
      if (
        context.optionValues.interactive === true ||
        context.globalOptionValues.interactive === true
      ) {
        return err(
          new UserError({
            source: cliSource,
            name: "ActionInteractiveUnsupported",
            message: errors["error.ActionInteractiveUnsupported"],
          })
        );
      }
      const invalid = invalidValueOption(context.command.options ?? []);
      if (invalid) {
        return err(
          new UserError({
            source: cliSource,
            name: "InvalidActionOptionValue",
            message: format(errors["error.InvalidActionOptionValue"], `--${invalid}`),
          })
        );
      }
      let projectPath = context.optionValues.projectPath;
      const input = Object.fromEntries(
        options.value
          .map((option) => option.questionName ?? option.name)
          .filter((key) => context.optionValues[key] !== undefined)
          .map((key) => [key, context.optionValues[key]])
      );
      if (
        name === ActionNames.PublishPolicy &&
        typeof input.policyPath === "string" &&
        input.policyPath.trim()
      ) {
        input.policyPath = path.resolve(process.cwd(), input.policyPath);
        projectPath = path.dirname(input.policyPath);
      }
      const args = validateActionArgs(name, input);
      if (args.isErr()) return err(args.error);
      if (typeof projectPath !== "string") {
        return err(new MissingRequiredOptionError(context.command.fullName, "--folder"));
      }
      const execution = await getFxCore().executeAction(name, args.value, { projectPath });
      const { result, ...details } = execution;
      const output = result.isOk() ? { ...details, outputs: result.value } : details;
      logger.info(JSON.stringify(output, redactResultValue, 2));
      return result.isErr() ? err(result.error) : ok(undefined);
    },
  };
}

export function getExecuteCommand(): CLICommand {
  return {
    name: "execute",
    description: `${commands.execute.description} ${commands["execute.action"].argumentMode}`,
    hidden: true,
    defaultInteractiveOption: false,
    options: [hiddenInteractiveOption],
    commands: [
      {
        name: "action",
        description: `${commands["execute.action"].description} ${commands["execute.action"].argumentMode}`,
        defaultInteractiveOption: false,
        options: [hiddenInteractiveOption],
        commands: Object.values(ActionNames).map(actionCommand),
      },
    ],
  };
}
