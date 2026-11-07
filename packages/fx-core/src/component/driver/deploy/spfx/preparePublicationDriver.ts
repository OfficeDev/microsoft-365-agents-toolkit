// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  FxError,
  SystemError,
  SystemErrorOptions,
  UserError,
  UserErrorOptions,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import { getLocalizedString } from "../../../../common/localizeUtils";
import { maskSecret } from "../../../../common/stringUtils";
import { ErrorCategory } from "../../../../error/types";
import { DriverContext } from "../../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../../interface/stepDriver";
import { addStartAndEndTelemetry } from "../../middleware/addStartAndEndTelemetry";
import {
  PreparedPublicationError,
  PreparedPublicationOutput,
  preparePublication,
} from "./utility/preparedPublication";
import { SharePointPackageInspectionError } from "./utility/packageFile";

export const preparePublicationActionName = "spfx/preparePublication";

type OutputKey = keyof PreparedPublicationOutput;

const outputKeys: OutputKey[] = [
  "publicationId",
  "preparedPublicationPath",
  "sourcePackageSha256",
  "sourceAppPackageSha256",
  "pluginSha256",
  "sharePointPackagePath",
  "sharePointPackageSha256",
  "appPackagePath",
  "appPackageSha256",
];

function isOutputKey(value: string): value is OutputKey {
  return outputKeys.some((key) => key === value);
}

interface PreparePublicationArgs {
  packagePath: string;
  appPackagePath: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function validateArgs(value: unknown): PreparePublicationArgs | undefined {
  if (
    !isRecord(value) ||
    typeof value["packagePath"] !== "string" ||
    value["packagePath"].length === 0 ||
    typeof value["appPackagePath"] !== "string" ||
    value["appPackagePath"].length === 0
  ) {
    return undefined;
  }
  const unexpected = Object.keys(value).filter(
    (key) => key !== "packagePath" && key !== "appPackagePath"
  );
  return unexpected.length === 0
    ? { packagePath: value["packagePath"], appPackagePath: value["appPackagePath"] }
    : undefined;
}

function toFxError(error: PreparedPublicationError | SharePointPackageInspectionError): FxError {
  const message = getLocalizedString("driver.spfx.prepared.error", error.message);
  if (
    error instanceof PreparedPublicationError &&
    (error.code.startsWith("Journal") ||
      error.code.startsWith("PublicationWrite") ||
      error.code.startsWith("PublicationPrepare") ||
      error.code.startsWith("ProjectPath"))
  ) {
    const options: SystemErrorOptions = {
      source: preparePublicationActionName,
      error,
      message: error.code,
      displayMessage: message,
      categories: [ErrorCategory.External],
    };
    return new SystemError(options);
  }
  const options: UserErrorOptions = {
    source: preparePublicationActionName,
    error,
    message: error.code,
    displayMessage: message,
    categories: [ErrorCategory.External],
  };
  return new UserError(options);
}

@Service(preparePublicationActionName)
export class PreparePublicationDriver implements StepDriver {
  public readonly description = getLocalizedString("driver.spfx.prepared.description");
  public readonly progressTitle = getLocalizedString("driver.spfx.prepared.progress");

  @hooks([addStartAndEndTelemetry(preparePublicationActionName, preparePublicationActionName)])
  public async execute(
    args: unknown,
    context: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const validatedArgs = validateArgs(args);
    if (!validatedArgs) {
      return {
        result: err(
          new UserError({
            source: preparePublicationActionName,
            message: "InvalidActionInput",
            displayMessage: getLocalizedString(
              "driver.spfx.prepared.invalidInput",
              "packagePath, appPackagePath"
            ),
            categories: [ErrorCategory.External],
          })
        ),
        summaries: [],
      };
    }

    if (
      outputEnvVarNames === undefined ||
      outputEnvVarNames.size !== outputKeys.length ||
      outputKeys.some((key) => !outputEnvVarNames.has(key)) ||
      new Set(outputEnvVarNames.values()).size !== outputEnvVarNames.size ||
      [...outputEnvVarNames].some(
        ([key, variable]) => !isOutputKey(key) || !/^[A-Za-z_][A-Za-z0-9_]*$/.test(variable)
      )
    ) {
      return {
        result: err(
          new UserError({
            source: preparePublicationActionName,
            message: "InvalidActionOutput",
            displayMessage: getLocalizedString(
              "driver.spfx.prepared.invalidOutput",
              "writeToEnvironmentFile"
            ),
            categories: [ErrorCategory.External],
          })
        ),
        summaries: [],
      };
    }

    const result = await preparePublication({
      projectPath: context.projectPath,
      packagePath: validatedArgs.packagePath,
      appPackagePath: validatedArgs.appPackagePath,
    });
    if (result.isErr()) {
      const error = toFxError(result.error);
      context.logProvider.error(maskSecret(error.message));
      return { result: err(error), summaries: [] };
    }

    const values = result.value;
    const outputs = new Map<string, string>();
    for (const key of outputKeys) {
      const value = values[key];
      const variable = outputEnvVarNames.get(key);
      if (variable !== undefined) {
        outputs.set(variable, value);
      }
    }
    return {
      result: ok(outputs),
      summaries: [getLocalizedString("driver.spfx.prepared.summary")],
    };
  }
}
