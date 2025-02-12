// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yuqzho@microsoft.com
 */

import {
  AdaptiveCardGenerator,
  ErrorResult as ApiSpecErrorResult,
  ErrorType as ApiSpecErrorType,
  ConstantString,
  ErrorType,
  InvalidAPIInfo,
  ListAPIResult,
  ParseOptions,
  ProjectType,
  SpecParser,
  SpecParserError,
  Utils,
  ValidationStatus,
  WarningResult,
  WarningType,
} from "@microsoft/m365-spec-parser";
import { AuthType, ListAPIInfo } from "@microsoft/m365-spec-parser/dist/src/interfaces";
import {
  ApiOperation,
  AppPackageFolderName,
  Context,
  FxError,
  Inputs,
  ManifestTemplateFileName,
  ManifestUtil,
  Platform,
  Result,
  SystemError,
  TeamsAppManifest,
  UserError,
  Warning,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { OpenAPIV3 } from "openapi-types";
import { EOL } from "os";
import path from "path";
import { FeatureFlags, featureFlagManager } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import { assembleError, MissingRequiredInputError } from "../../../error";
import {
  apiPluginApiSpecOptionId,
  CustomCopilotRagOptions,
  ProgrammingLanguage,
  QuestionNames,
} from "../../../question/constants";
import {
  APIKeyAuthType,
  MicrosoftEntraAuthType,
  OAuthAuthType,
  SummaryConstant,
} from "../../configManager/constant";
import { manifestUtils } from "../../driver/teamsApp/utils/ManifestUtils";
import { pluginManifestUtils } from "../../driver/teamsApp/utils/PluginManifestUtils";
import {
  ApiSpecTelemetryPropertis,
  sendTelemetryErrorEvent,
  TelemetryProperty,
} from "../../../common/telemetry";
import * as util from "util";
import { SpecParserSource } from "../../../common/constants";
import { MetadataV3 } from "../../../common/versionMetadata";
import { ActionInjector, AuthActionInjectResult } from "../../configManager/actionInjector";

const enum telemetryProperties {
  validationStatus = "validation-status",
  validationErrors = "validation-errors",
  validationWarnings = "validation-warnings",
  validApisCount = "valid-apis-count",
  allApisCount = "all-apis-count",
  specHash = "spec-hash",
  bearerTokenAuthCount = "bearer-token-auth-count",
  oauth2AuthCount = "oauth2-auth-count",
  otherAuthCount = "other-auth-count",
  isFromAddingApi = "is-from-adding-api",
  failedReason = "failed-reason",
  generateType = "generate-type",
  projectType = "project-type",
}

const enum telemetryEvents {
  validateApiSpec = "validate-api-spec",
  listApis = "spec-parser-list-apis-result",
  failedToGetGenerateWarning = "failed-to-get-generate-warning",
}

export function getParserOptions(
  type: ProjectType,
  isDeclarativeCopilot?: boolean,
  platform?: string
): ParseOptions {
  return type === ProjectType.Copilot
    ? {
        isGptPlugin: isDeclarativeCopilot,
        allowAPIKeyAuth: false,
        allowBearerTokenAuth: !!platform && platform === Platform.VS ? false : true,
        allowMultipleParameters: true,
        allowOauth2: !!platform && platform === Platform.VS ? false : true,
        projectType: ProjectType.Copilot,
        allowMissingId: true,
        allowSwagger: true,
        allowMethods: [
          "get",
          "post",
          "put",
          "delete",
          "patch",
          "head",
          "connect",
          "options",
          "trace",
        ],
        allowResponseSemantics: true,
        allowConversationStarters: false, // Conversation starters in the plugin file are no longer used; they are now sourced from the declarativeAgent file.
        allowConfirmation: false, // confirmation is not stable for public preview in Sydney, so it's temporarily set to false
      }
    : type === ProjectType.TeamsAi
    ? {
        allowAPIKeyAuth: true,
        allowBearerTokenAuth: true,
        allowMultipleParameters: true,
        allowOauth2: true,
        projectType: ProjectType.TeamsAi,
        allowMethods: [
          "get",
          "post",
          "put",
          "delete",
          "patch",
          "head",
          "connect",
          "options",
          "trace",
        ],
      }
    : {
        projectType: type,
        allowBearerTokenAuth: !!platform && platform === Platform.VS ? false : true, // Currently, API key auth support is actually bearer token auth
        allowMultipleParameters: true,
        allowOauth2: featureFlagManager.getBooleanValue(FeatureFlags.SMEOAuth),
      };
}

export const specParserGenerateResultTelemetryEvent = "spec-parser-generate-result";
export const specParserGenerateResultAllSuccessTelemetryProperty = "all-success";
export const specParserGenerateResultWarningsTelemetryProperty = "warnings";

export const invalidApiSpecErrorName = "invalid-api-spec";
const apiSpecNotUsedInPlugin = "api-spec-not-used-in-plugin";

export interface ErrorResult {
  /**
   * The type of error.
   */
  type: ApiSpecErrorType;

  /**
   * The content of the error.
   */
  content: string;

  data?: any;
}

export async function listOperations(
  context: Context,
  apiSpecUrl: string | undefined,
  inputs: Inputs,
  includeExistingAPIs = true,
  shouldLogWarning = true,
  existingCorrelationId?: string
): Promise<Result<ApiOperation[], ErrorResult[]>> {
  const isPlugin = inputs[QuestionNames.ApiPluginType] === apiPluginApiSpecOptionId;
  const isCustomApi =
    inputs[QuestionNames.CustomCopilotRag] === CustomCopilotRagOptions.customApi().id;
  const projectType = isPlugin
    ? ProjectType.Copilot
    : isCustomApi
    ? ProjectType.TeamsAi
    : ProjectType.SME;

  try {
    const specParser = new SpecParser(
      apiSpecUrl as string,
      getParserOptions(projectType, undefined, inputs.platform)
    );
    const validationRes = await specParser.validate();
    validationRes.errors = formatValidationErrors(validationRes.errors, inputs);

    logValidationResults(
      projectType,
      validationRes.errors,
      validationRes.warnings,
      context,
      shouldLogWarning,
      false,
      validationRes.specHash,
      existingCorrelationId
    );
    if (validationRes.status === ValidationStatus.Error) {
      return err(validationRes.errors);
    }

    const listResult: ListAPIResult = await specParser.list();

    const invalidAPIs = listResult.APIs.filter((value) => !value.isValid);
    for (const invalidAPI of invalidAPIs) {
      context.logProvider.warning(
        `${invalidAPI.api} ${getLocalizedString(
          "core.copilotPlugin.list.unsupportedBecause"
        )} ${invalidAPI.reason.map(mapInvalidReasonToMessage).join(", ")}`
      );
    }

    const bearerTokenAuthAPIs = listResult.APIs.filter(
      (api) => api.auth && Utils.isBearerTokenAuth(api.auth.authScheme)
    );

    const oauth2AuthAPIs = listResult.APIs.filter(
      (api) => api.auth && Utils.isOAuthWithAuthCodeFlow(api.auth.authScheme)
    );

    const otherAuthAPIs = listResult.APIs.filter(
      (api) =>
        api.auth &&
        !Utils.isOAuthWithAuthCodeFlow(api.auth.authScheme) &&
        !Utils.isBearerTokenAuth(api.auth.authScheme)
    );

    let operations = listResult.APIs.filter((value) => value.isValid);
    context.telemetryReporter.sendTelemetryEvent(telemetryEvents.listApis, {
      [telemetryProperties.generateType]: projectType.toString(),
      [telemetryProperties.validApisCount]: listResult.validAPICount.toString(),
      [telemetryProperties.allApisCount]: listResult.allAPICount.toString(),
      [telemetryProperties.isFromAddingApi]: (!includeExistingAPIs).toString(),
      [telemetryProperties.bearerTokenAuthCount]: bearerTokenAuthAPIs.length.toString(),
      [telemetryProperties.oauth2AuthCount]: oauth2AuthAPIs.length.toString(),
      [telemetryProperties.otherAuthCount]: otherAuthAPIs.length.toString(),
      [telemetryProperties.specHash]: validationRes.specHash!,
    });

    // Filter out exsiting APIs
    if (!includeExistingAPIs) {
      const teamsManifestPath = inputs[QuestionNames.ManifestPath];
      if (!teamsManifestPath) {
        throw new MissingRequiredInputError("teamsManifestPath", "inputs");
      }
      const manifest = await manifestUtils._readAppManifest(teamsManifestPath);
      let existingOperations: string[] = [];
      if (manifest.isOk()) {
        let isOriginalSpec;

        if (isPlugin) {
          existingOperations = await listPluginExistingOperations(
            manifest.value,
            teamsManifestPath,
            inputs[QuestionNames.DestinationApiSpecFilePath]
          );

          const operationAPIs = operations.map((operation) => operation.api);
          isOriginalSpec = existingOperations.every((operation) =>
            operationAPIs.includes(operation)
          );
        } else {
          const existingOperationIds = manifestUtils.getOperationIds(manifest.value);
          existingOperations = operations
            .filter((operation) => existingOperationIds.includes(operation.operationId))
            .map((operation) => operation.api);

          isOriginalSpec = existingOperations.length === existingOperationIds.length;
        }

        if (!isOriginalSpec) {
          const errors = formatValidationErrors(
            [
              {
                type: ApiSpecErrorType.AddedAPINotInOriginalSpec,
                content: "",
              },
            ],
            inputs
          );

          logValidationResults(
            projectType,
            errors,
            [],
            context,
            true,
            false,
            validationRes.specHash,
            existingCorrelationId
          );
          return err(errors);
        }

        operations = operations.filter(
          (operation: ListAPIInfo) => !existingOperations.includes(operation.api)
        );
        // No extra API can be added
        if (operations.length == 0) {
          const errors = formatValidationErrors(
            [
              {
                type: ApiSpecErrorType.NoExtraAPICanBeAdded,
                content: "",
              },
            ],
            inputs
          );
          logValidationResults(
            projectType,
            errors,
            [],
            context,
            true,
            false,
            validationRes.specHash,
            existingCorrelationId
          );
          return err(errors);
        }
      } else {
        throw manifest.error;
      }
    }

    const sortedOperations = sortOperations(operations);
    return ok(sortedOperations);
  } catch (e) {
    if (e instanceof SpecParserError) {
      throw convertSpecParserErrorToFxError(e);
    } else {
      throw e;
    }
  }
}

function sortOperations(operations: ListAPIInfo[]): ApiOperation[] {
  const operationsWithSeparator: ApiOperation[] = [];
  for (const operation of operations) {
    const arr = operation.api.toUpperCase().split(" ");
    const result: ApiOperation = {
      id: operation.api,
      label: operation.api,
      groupName: arr[0],
      detail: !operation.auth
        ? getLocalizedString("core.copilotPlugin.api.noAuth")
        : Utils.isBearerTokenAuth(operation.auth.authScheme)
        ? getLocalizedString("core.copilotPlugin.api.apiKeyAuth")
        : Utils.isOAuthWithAuthCodeFlow(operation.auth.authScheme)
        ? getLocalizedString("core.copilotPlugin.api.oauth")
        : Utils.isAPIKeyAuthButNotInCookie(operation.auth.authScheme)
        ? getLocalizedString("core.copilotPlugin.api.apiKeyWithHeaderOrQuery")
        : getLocalizedString("core.copilotPlugin.api.notSupportedAuth"),
      data: {
        serverUrl: operation.server,
      },
    };

    if (operation.auth) {
      if (
        Utils.isBearerTokenAuth(operation.auth.authScheme) ||
        Utils.isAPIKeyAuthButNotInCookie(operation.auth.authScheme)
      ) {
        result.data.authType = "apiKey";
        result.data.authName = operation.auth.name;
      } else if (Utils.isOAuthWithAuthCodeFlow(operation.auth.authScheme)) {
        result.data.authType = "oauth2";
        result.data.authName = operation.auth.name;
      }
    }

    operationsWithSeparator.push(result);
  }

  return operationsWithSeparator.sort((operation1: ApiOperation, operation2: ApiOperation) => {
    const arr1 = operation1.id.toLowerCase().split(" ");
    const arr2 = operation2.id.toLowerCase().split(" ");
    return arr1[0] < arr2[0] ? -1 : arr1[0] > arr2[0] ? 1 : arr1[1].localeCompare(arr2[1]);
  });
}

function formatTelemetryValidationProperty(result: ErrorResult | WarningResult): string {
  return result.type.toString();
}

export async function listPluginExistingOperations(
  manifest: TeamsAppManifest,
  teamsManifestPath: string,
  destinationApiSpecFilePath: string
): Promise<string[]> {
  const getApiSPecFileRes = await pluginManifestUtils.getApiSpecFilePathFromTeamsManifest(
    manifest,
    teamsManifestPath
  );
  if (getApiSPecFileRes.isErr()) {
    throw getApiSPecFileRes.error;
  }

  let apiSpecFilePath;
  const apiSpecFiles = getApiSPecFileRes.value;
  for (const file of apiSpecFiles) {
    if (path.resolve(file) === path.resolve(destinationApiSpecFilePath)) {
      apiSpecFilePath = file;
      break;
    }
  }
  if (!apiSpecFilePath) {
    throw new UserError(
      "listPluginExistingOperations",
      apiSpecNotUsedInPlugin,
      getLocalizedString("error.copilotPlugin.apiSpecNotUsedInPlugin", destinationApiSpecFilePath),
      getLocalizedString("error.copilotPlugin.apiSpecNotUsedInPlugin", destinationApiSpecFilePath)
    );
  }

  const specParser = new SpecParser(apiSpecFilePath, getParserOptions(ProjectType.Copilot));
  const listResult = await specParser.list();
  return listResult.APIs.map((o) => o.api);
}

interface SpecParserOutputFilePath {
  destinationApiSpecFilePath: string;
  pluginManifestFilePath?: string;
  responseTemplateFolder?: string;
}

interface SpecParserGenerateResult {
  warnings: WarningResult[];
}
export async function generateFromApiSpec(
  specParser: SpecParser,
  teamsManifestPath: string,
  inputs: Inputs,
  context: Context,
  sourceComponent: string,
  projectType: ProjectType,
  outputFilePath: SpecParserOutputFilePath
): Promise<Result<SpecParserGenerateResult, FxError>> {
  const operations = inputs[QuestionNames.ApiOperation] as string[];
  const validationRes = await specParser.validate();
  const warnings = validationRes.warnings;
  const operationIdWarning = warnings.find((w) => w.type === WarningType.OperationIdMissing);

  if (operationIdWarning && operationIdWarning.data) {
    const apisMissingOperationId = (operationIdWarning.data as string[]).filter((api) =>
      operations.includes(api)
    );
    if (apisMissingOperationId.length > 0) {
      operationIdWarning.content = util.format(
        getLocalizedString("core.common.MissingOperationId"),
        apisMissingOperationId.join(", ")
      );
      delete operationIdWarning.data;
    } else {
      warnings.splice(warnings.indexOf(operationIdWarning), 1);
    }
  }

  const specVersionWarning = warnings.find((w) => w.type === WarningType.ConvertSwaggerToOpenAPI);
  if (specVersionWarning) {
    specVersionWarning.content = ""; // We don't care content of this warning
  }

  if (validationRes.status === ValidationStatus.Error) {
    logValidationResults(
      projectType,
      validationRes.errors,
      warnings,
      context,
      false,
      true,
      validationRes.specHash
    );
    const errorMessage =
      inputs.platform === Platform.VSCode
        ? getLocalizedString(
            "core.createProjectQuestion.apiSpec.multipleValidationErrors.vscode.message"
          )
        : getLocalizedString("core.createProjectQuestion.apiSpec.multipleValidationErrors.message");
    return err(new UserError(sourceComponent, invalidApiSpecErrorName, errorMessage, errorMessage));
  }

  try {
    const generateResult =
      projectType === ProjectType.Copilot
        ? await specParser.generateForCopilot(
            teamsManifestPath,
            operations,
            outputFilePath.destinationApiSpecFilePath,
            outputFilePath.pluginManifestFilePath!
          )
        : await specParser.generate(
            teamsManifestPath,
            operations,
            outputFilePath.destinationApiSpecFilePath,
            projectType === ProjectType.TeamsAi ? undefined : outputFilePath.responseTemplateFolder
          );

    // Send SpecParser.generate() warnings
    context.telemetryReporter.sendTelemetryEvent(specParserGenerateResultTelemetryEvent, {
      [telemetryProperties.generateType]: projectType.toString(),
      [specParserGenerateResultAllSuccessTelemetryProperty]: generateResult.allSuccess.toString(),
      [specParserGenerateResultWarningsTelemetryProperty]: generateResult.warnings
        .map((w) => w.type.toString() + ": " + w.content)
        .join(";"),
      [TelemetryProperty.Component]: sourceComponent,
    });

    if (generateResult.warnings && generateResult.warnings.length > 0) {
      generateResult.warnings.find((o) => {
        if (o.type === WarningType.OperationOnlyContainsOptionalParam) {
          o.content = ""; // We don't care content of this warning
        }
      });
      warnings.push(...generateResult.warnings);
    }

    return ok({ warnings });
  } catch (e) {
    let error: FxError;
    if (e instanceof SpecParserError) {
      error = convertSpecParserErrorToFxError(e);
    } else {
      error = assembleError(e, sourceComponent);
    }
    return err(error);
  }
}

export function logValidationResults(
  projectType: ProjectType,
  errors: ErrorResult[],
  warnings: WarningResult[],
  context: Context,
  shouldLogWarning: boolean,
  shouldSkipTelemetry: boolean,
  specHash?: string,
  existingCorrelationId?: string
): void {
  if (!shouldSkipTelemetry) {
    const properties: { [key: string]: string } = {
      [telemetryProperties.validationStatus]:
        errors.length !== 0 ? "error" : warnings.length !== 0 ? "warning" : "success",
      [telemetryProperties.validationErrors]: errors
        .map((error: ErrorResult) => formatTelemetryValidationProperty(error))
        .join(";"),
      [telemetryProperties.validationWarnings]: warnings
        .map((warn: WarningResult) => formatTelemetryValidationProperty(warn))
        .join(";"),
      [telemetryProperties.projectType]: projectType.toString(),
    };

    if (specHash) {
      properties[telemetryProperties.specHash] = specHash;
    }

    const specNotValidError = errors.find((error) => error.type === ErrorType.SpecNotValid);
    if (specNotValidError) {
      properties[ApiSpecTelemetryPropertis.SpecNotValidDetails] = specNotValidError.content;
    }

    if (existingCorrelationId) {
      properties["correlation-id"] = existingCorrelationId;
    }
    context.telemetryReporter.sendTelemetryEvent(telemetryEvents.validateApiSpec, properties);
  }

  if (errors.length === 0 && (warnings.length === 0 || !shouldLogWarning)) {
    return;
  }

  // errors > 0 || (warnings > 0 && shouldLogWarning)
  const errorMessage = errors
    .map((error) => {
      return `${SummaryConstant.Failed} ${error.content}`;
    })
    .join(EOL);
  const warningMessage = shouldLogWarning
    ? warnings
        .map((warning) => {
          return `${SummaryConstant.NotExecuted} ${warning.content}`;
        })
        .join(EOL)
    : "";

  const failed = errors.length;
  const warns = warnings.length;
  const summaryStr = [];

  if (failed > 0) {
    summaryStr.push(
      getLocalizedString("core.copilotPlugin.validate.summary.validate.failed", failed)
    );
  }
  if (warns > 0 && shouldLogWarning) {
    summaryStr.push(
      getLocalizedString("core.copilotPlugin.validate.summary.validate.warning", warns)
    );
  }

  const outputMessage =
    EOL +
    getLocalizedString(
      "core.copilotPlugin.validate.apiSpec.summary",
      summaryStr.join(", "),
      errorMessage,
      warningMessage
    );

  void context.logProvider.info(outputMessage);
}

export async function injectAuthAction(
  projectPath: string,
  authName: string,
  authScheme: AuthType | undefined,
  outputApiSpecPath: string,
  forceToAddNew: boolean,
  authType?: string
): Promise<AuthActionInjectResult | undefined> {
  const ymlPath = path.join(projectPath, MetadataV3.configFile);
  const localYamlPath = path.join(projectPath, MetadataV3.localConfigFile);

  const relativeSpecPath = "./" + path.relative(projectPath, outputApiSpecPath).replace(/\\/g, "/");

  if ((!!authScheme && Utils.isBearerTokenAuth(authScheme)) || authType === APIKeyAuthType) {
    const res = await ActionInjector.injectCreateAPIKeyAction(
      ymlPath,
      authName,
      relativeSpecPath,
      forceToAddNew
    );

    if (await fs.pathExists(localYamlPath)) {
      await ActionInjector.injectCreateAPIKeyAction(
        localYamlPath,
        authName,
        relativeSpecPath,
        forceToAddNew
      );
    }
    return res;
  } else if (
    (!!authScheme && Utils.isOAuthWithAuthCodeFlow(authScheme)) ||
    authType === OAuthAuthType ||
    authType === MicrosoftEntraAuthType
  ) {
    const res = await ActionInjector.injectCreateOAuthAction(
      ymlPath,
      authName,
      relativeSpecPath,
      forceToAddNew,
      authType === MicrosoftEntraAuthType
    );

    if (await fs.pathExists(localYamlPath)) {
      await ActionInjector.injectCreateOAuthAction(
        localYamlPath,
        authName,
        relativeSpecPath,
        forceToAddNew,
        authType === MicrosoftEntraAuthType
      );
    }
    return res;
  }
}

/**
 * Generate scaffolding warning summary.
 * @param warnings warnings returned from spec-parser.
 * @param teamsManifest Teams manifest.
 * @param apiSpecFilePath API spec path relative of project path.
 * @param pluginManifestPath Plugin manifest path relative of project path.
 * @param projectPath Project path.
 * @returns Warning message.
 */
export async function generateScaffoldingSummary(
  warnings: Warning[],
  teamsManifest: TeamsAppManifest,
  apiSpecFilePath: string,
  pluginManifestPath: string | undefined,
  projectPath: string
): Promise<string> {
  const apiSpecWarningMessage = formatApiSpecValidationWarningMessage(
    warnings,
    apiSpecFilePath,
    teamsManifest
  );
  const manifestWarningResult = validateTeamsManifestLength(teamsManifest, warnings);
  const manifestWarningMessage = manifestWarningResult.map((warn) => {
    return `${SummaryConstant.NotExecuted} ${warn}`;
  });

  let pluginWarningMessage: string[] = [];
  if (pluginManifestPath) {
    const pluginManifestWarningResult = await validatePluginManifestLength(
      pluginManifestPath,
      projectPath,
      warnings
    );
    pluginWarningMessage = pluginManifestWarningResult.map((warn) => {
      return `${SummaryConstant.NotExecuted} ${warn}`;
    });
  }

  if (
    apiSpecWarningMessage.length ||
    manifestWarningMessage.length ||
    pluginWarningMessage.length
  ) {
    let details = "";
    if (apiSpecWarningMessage.length) {
      details += EOL + apiSpecWarningMessage.join(EOL);
    }

    if (manifestWarningMessage.length) {
      details += EOL + manifestWarningMessage.join(EOL);
    }

    if (pluginWarningMessage.length) {
      details += EOL + pluginWarningMessage.join(EOL);
    }

    return getLocalizedString("core.copilotPlugin.scaffold.summary", details);
  } else {
    return "";
  }
}

function formatApiSpecValidationWarningMessage(
  specWarnings: Warning[],
  apiSpecFileName: string,
  teamsManifest: TeamsAppManifest
): string[] {
  const resultWarnings = [];
  const operationIdWarning = specWarnings.find((w) => w.type === WarningType.OperationIdMissing);

  if (operationIdWarning) {
    const isApiMe = ManifestUtil.parseCommonProperties(teamsManifest).isApiME;
    resultWarnings.push(
      getLocalizedString(
        "core.copilotPlugin.scaffold.summary.warning.operationId",
        `${SummaryConstant.NotExecuted} ${operationIdWarning.content}`,
        isApiMe ? ManifestTemplateFileName : apiSpecFileName
      )
    );
  }

  const swaggerWarning = specWarnings.find((w) => w.type === WarningType.ConvertSwaggerToOpenAPI);

  if (swaggerWarning) {
    resultWarnings.push(
      `${SummaryConstant.NotExecuted} ` +
        getLocalizedString(
          "core.copilotPlugin.scaffold.summary.warning.swaggerVersion",
          apiSpecFileName
        )
    );
  }

  const specialCharactersWarnings = specWarnings.filter(
    (w) => w.type === WarningType.OperationIdContainsSpecialCharacters
  );

  specialCharactersWarnings.forEach((warning) => {
    resultWarnings.push(
      getLocalizedString(
        "core.copilotPlugin.scaffold.summary.warning.operationIdContainsSpecialCharacters",
        warning.data.operationId,
        warning.data.operationId.replace(/[^a-zA-Z0-9]/g, "_")
      )
    );
  });

  return resultWarnings;
}

function validateTeamsManifestLength(
  teamsManifest: TeamsAppManifest,
  warnings: Warning[]
): string[] {
  const nameShortLimit = 30;
  const nameFullLimit = 100;
  const descriptionShortLimit = 80;
  const descriptionFullLimit = 4000;
  const appnameSuffixPlaceholder = "${{APP_NAME_SUFFIX}}";
  const devEnv = "dev";
  const resultWarnings = [];

  // validate name
  const shortNameLength = teamsManifest.name.short.includes(appnameSuffixPlaceholder)
    ? teamsManifest.name.short.length - appnameSuffixPlaceholder.length + devEnv.length
    : teamsManifest.name.short.length;
  if (shortNameLength > nameShortLimit) {
    resultWarnings.push(formatLengthExceedingErrorMessage("/name/short", nameShortLimit));
  }

  if (!!teamsManifest.name.full && teamsManifest.name.full?.length > nameFullLimit) {
    resultWarnings.push(formatLengthExceedingErrorMessage("/name/full", nameFullLimit));
  }

  // validate description
  if (teamsManifest.description.short.length > descriptionShortLimit) {
    resultWarnings.push(
      formatLengthExceedingErrorMessage("/description/short", descriptionShortLimit)
    );
  }
  if (!teamsManifest.description.full?.length) {
    resultWarnings.push(
      getLocalizedString(
        "core.copilotPlugin.scaffold.summary.warning.teamsManifest.missingFullDescription"
      ) +
        getLocalizedString(
          "core.copilotPlugin.scaffold.summary.warning.teamsManifest.mitigation",
          "full/description",
          path.join(AppPackageFolderName, ManifestTemplateFileName)
        )
    );
  }
  if (teamsManifest.description.full!.length > descriptionFullLimit) {
    resultWarnings.push(
      formatLengthExceedingErrorMessage("/description/full", descriptionFullLimit)
    );
  }

  // validate command
  if (ManifestUtil.parseCommonProperties(teamsManifest).isApiME) {
    const optionalParamsOnlyWarnings = warnings.filter(
      (o) => o.type === WarningType.OperationOnlyContainsOptionalParam
    );

    if (optionalParamsOnlyWarnings) {
      for (const optionalParamsOnlyWarning of optionalParamsOnlyWarnings) {
        resultWarnings.push(
          getLocalizedString(
            "core.copilotPlugin.scaffold.summary.warning.api.optionalParametersOnly",
            optionalParamsOnlyWarning.data.commandId,
            optionalParamsOnlyWarning.data.commandId
          ) +
            getLocalizedString(
              "core.copilotPlugin.scaffold.summary.warning.api.optionalParametersOnly.mitigation",
              optionalParamsOnlyWarning.data.parameterName,
              optionalParamsOnlyWarning.data.commandId,
              path.join(AppPackageFolderName, ManifestTemplateFileName),
              path.join(
                AppPackageFolderName,
                teamsManifest.composeExtensions![0].apiSpecificationFile ?? ""
              )
            )
        );
      }
    }

    const commands = teamsManifest.composeExtensions![0].commands;

    for (const command of commands) {
      if (command.type === "query") {
        if (!command.apiResponseRenderingTemplateFile) {
          const errorDetail = warnings.find(
            (w) => w.type === WarningType.GenerateCardFailed && w.data === command.id
          )?.content;
          resultWarnings.push(
            getLocalizedString(
              "core.copilotPlugin.scaffold.summary.warning.teamsManifest.missingCardTemlate",
              "apiResponseRenderingTemplateFile",
              command.id
            ) +
              getLocalizedString(
                "core.copilotPlugin.scaffold.summary.warning.teamsManifest.missingCardTemlate.mitigation",
                AppPackageFolderName,
                `composeExtensions/commands/${command.id}/apiResponseRenderingTemplateFile`,
                path.join(AppPackageFolderName, ManifestTemplateFileName)
              ) +
              (errorDetail ? EOL + errorDetail : "")
          );
        }
      }
    }
  }

  return resultWarnings;
}

async function validatePluginManifestLength(
  pluginManifestPath: string,
  projectPath: string,
  warnings: Warning[]
): Promise<string[]> {
  const functionDescriptionLimit = 100;
  const resultWarnings: string[] = [];

  const manifestRes = await pluginManifestUtils.readPluginManifestFile(
    path.join(projectPath, pluginManifestPath)
  );
  if (manifestRes.isErr()) {
    sendTelemetryErrorEvent(
      "spec-generator",
      telemetryEvents.failedToGetGenerateWarning,
      manifestRes.error
    );
    return [];
  }

  // validate function description
  const functions = manifestRes.value.functions;
  const functionDescriptionWarnings = warnings
    .filter((w) => w.type === WarningType.FuncDescriptionTooLong)
    .map((w) => w.data);
  if (functions) {
    functions.forEach((func) => {
      if (!func.description) {
        resultWarnings.push(
          getLocalizedString(
            "core.copilotPlugin.scaffold.summary.warning.pluginManifest.missingFunctionDescription",
            func.name
          ) +
            getLocalizedString(
              "core.copilotPlugin.scaffold.summary.warning.pluginManifest.missingFunctionDescription.mitigation",
              func.name,
              pluginManifestPath
            )
        );
      } else if (functionDescriptionWarnings.includes(func.name)) {
        resultWarnings.push(
          getLocalizedString(
            "core.copilotPlugin.scaffold.summary.warning.pluginManifest.functionDescription.lengthExceeding",
            func.name,
            functionDescriptionLimit
          ) +
            getLocalizedString(
              "core.copilotPlugin.scaffold.summary.warning.pluginManifest.functionDescription.lengthExceeding.mitigation",
              func.name,
              pluginManifestPath
            )
        );
      }
    });
  }
  return resultWarnings;
}

function formatLengthExceedingErrorMessage(field: string, limit: number): string {
  return (
    getLocalizedString(
      "core.copilotPlugin.scaffold.summary.warning.teamsManifest.lengthExceeding",
      field,
      limit.toString()
    ) +
    getLocalizedString(
      "core.copilotPlugin.scaffold.summary.warning.teamsManifest.mitigation",
      field,
      path.join(AppPackageFolderName, ManifestTemplateFileName)
    )
  );
}

export function convertSpecParserErrorToFxError(error: SpecParserError): FxError {
  return new SystemError(
    SpecParserSource,
    error.errorType.toString(),
    error.message,
    error.message
  );
}

export function formatValidationErrors(
  errors: ApiSpecErrorResult[],
  inputs: Inputs
): ApiSpecErrorResult[] {
  return errors.map((error) => {
    return {
      type: error.type,
      content: formatValidationErrorContent(error, inputs),
      data: error.data,
    };
  });
}

function mapInvalidReasonToMessage(reason: ErrorType): string {
  switch (reason) {
    case ErrorType.AuthTypeIsNotSupported:
      return getLocalizedString("core.common.invalidReason.AuthTypeIsNotSupported");
    case ErrorType.MissingOperationId:
      return getLocalizedString("core.common.invalidReason.MissingOperationId");
    case ErrorType.PostBodyContainMultipleMediaTypes:
      return getLocalizedString("core.common.invalidReason.PostBodyContainMultipleMediaTypes");
    case ErrorType.ResponseContainMultipleMediaTypes:
      return getLocalizedString("core.common.invalidReason.ResponseContainMultipleMediaTypes");
    case ErrorType.ResponseJsonIsEmpty:
      return getLocalizedString("core.common.invalidReason.ResponseJsonIsEmpty");
    case ErrorType.PostBodyContainsRequiredUnsupportedSchema:
      return getLocalizedString(
        "core.common.invalidReason.PostBodyContainsRequiredUnsupportedSchema"
      );
    case ErrorType.ParamsContainRequiredUnsupportedSchema:
      return getLocalizedString("core.common.invalidReason.ParamsContainRequiredUnsupportedSchema");
    case ErrorType.ExceededRequiredParamsLimit:
      return getLocalizedString("core.common.invalidReason.ExceededRequiredParamsLimit");
    case ErrorType.NoParameter:
      return getLocalizedString("core.common.invalidReason.NoParameter");
    case ErrorType.NoAPIInfo:
      return getLocalizedString("core.common.invalidReason.NoAPIInfo");
    case ErrorType.MethodNotAllowed:
      return getLocalizedString("core.common.invalidReason.MethodNotAllowed");
    case ErrorType.UrlPathNotExist:
      return getLocalizedString("core.common.invalidReason.UrlPathNotExist");
    case ErrorType.CircularReferenceNotSupported:
      return getLocalizedString("core.common.invalidReason.CircularReference");
    default:
      return reason.toString();
  }
}

function formatValidationErrorContent(error: ApiSpecErrorResult, inputs: Inputs): string {
  const isPlugin = inputs[QuestionNames.ApiPluginType] === apiPluginApiSpecOptionId;
  try {
    switch (error.type) {
      case ErrorType.SpecNotValid: {
        let content: string = error.content;
        if (error.content.startsWith("ResolverError: Error downloading")) {
          content = error.content
            .split("\n")
            .map((o) => o.trim())
            .join(". ");
          content = `${content}. ${getLocalizedString("core.common.ErrorFetchApiSpec")}`;
        }
        return content;
      }

      case ErrorType.RemoteRefNotSupported:
        return getLocalizedString("core.common.RemoteRefNotSupported", error.data.join(", "));
      case ErrorType.NoServerInformation:
        return getLocalizedString("core.common.NoServerInformation");
      case ErrorType.UrlProtocolNotSupported:
        return getLocalizedString("core.common.UrlProtocolNotSupported", error.data);
      case ErrorType.RelativeServerUrlNotSupported:
        return getLocalizedString("core.common.RelativeServerUrlNotSupported");
      case ErrorType.NoSupportedApi:
        const messages = [];
        const invalidAPIInfo = error.data as InvalidAPIInfo[];
        for (const info of invalidAPIInfo) {
          // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
          const mes = `${info.api}: ${info.reason.map(mapInvalidReasonToMessage).join(", ")}`;
          messages.push(mes);
        }

        if (messages.length === 0) {
          messages.push(getLocalizedString("core.common.invalidReason.NoAPIs"));
        }
        return isPlugin
          ? getLocalizedString("core.common.NoSupportedApiCopilot", messages.join("\n"))
          : getLocalizedString("core.common.NoSupportedApi", messages.join("\n"));
      case ErrorType.NoExtraAPICanBeAdded:
        return isPlugin
          ? getLocalizedString("error.copilot.noExtraAPICanBeAdded")
          : getLocalizedString("error.apime.noExtraAPICanBeAdded");
      case ErrorType.ResolveServerUrlFailed:
        return error.content;
      case ErrorType.Cancelled:
        return getLocalizedString("core.common.CancelledMessage");
      case ErrorType.SwaggerNotSupported:
        return getLocalizedString("core.common.SwaggerNotSupported");
      case ErrorType.SpecVersionNotSupported:
        return getLocalizedString("core.common.SpecVersionNotSupported", error.data);
      case ErrorType.AddedAPINotInOriginalSpec:
        return getLocalizedString("core.common.AddedAPINotInOriginalSpec");

      default:
        return error.content;
    }
  } catch (e) {
    return error.content;
  }
}

interface SpecObject {
  pathUrl: string;
  method: string;
  item: OpenAPIV3.OperationObject;
  auth: boolean;
}

function parseSpec(spec: OpenAPIV3.Document): [SpecObject[], boolean] {
  const res: SpecObject[] = [];
  let needAuth = false;

  const paths = spec.paths;
  if (paths) {
    for (const pathUrl in paths) {
      const pathItem = paths[pathUrl];
      if (pathItem) {
        const operations = pathItem;
        for (const method in operations) {
          if (ConstantString.AllOperationMethods.includes(method)) {
            const operationItem = (operations as any)[method] as OpenAPIV3.OperationObject;
            if (operationItem) {
              const authResult = Utils.getAuthArray(operationItem.security, spec);
              const hasAuth = authResult.length != 0;
              if (hasAuth) {
                needAuth = true;
              }
              res.push({
                item: operationItem,
                method: method,
                pathUrl: pathUrl,
                auth: hasAuth,
              });
            }
          }
        }
      }
    }
  }

  return [res, needAuth];
}

const commonLanguages = [
  ProgrammingLanguage.TS,
  ProgrammingLanguage.JS,
  ProgrammingLanguage.PY,
  ProgrammingLanguage.CSharp,
];

async function updatePromptForCustomApi(
  spec: OpenAPIV3.Document,
  language: string,
  chatFolder: string
): Promise<void> {
  if (commonLanguages.includes(language as ProgrammingLanguage)) {
    const object = `{ "path": null, "body": null, "query": null }`;
    const cSharpObject = `{ "path": {}, "body": {}, "query": {} }`;
    const promptFilePath = path.join(chatFolder, "skprompt.txt");
    const prompt = `The following is a conversation with an AI assistant.\nThe assistant can help to call APIs for the open api spec file${
      spec.info.description ? ". " + spec.info.description : "."
    }\nIf the API doesn't require parameters, invoke it with default JSON object ${
      (language as ProgrammingLanguage) === ProgrammingLanguage.CSharp ? cSharpObject : object
    }.\n\ncontext:\nAvailable actions: {{getAction}}.`;
    await fs.writeFile(promptFilePath, prompt, { encoding: "utf-8", flag: "w" });
  }
}

async function updateAdaptiveCardForCustomApi(
  specItems: SpecObject[],
  language: string,
  destinationPath: string
): Promise<WarningResult[]> {
  const warnings: WarningResult[] = [];
  if (commonLanguages.includes(language as ProgrammingLanguage)) {
    let adaptiveCardsFolderPath = path.join(destinationPath, "src", "adaptiveCards");
    if (language === ProgrammingLanguage.CSharp) {
      adaptiveCardsFolderPath = path.join(destinationPath, "adaptiveCards");
    }
    await fs.ensureDir(adaptiveCardsFolderPath);

    for (const item of specItems) {
      const name = item.item.operationId!.replace(/[^a-zA-Z0-9]/g, "_");
      try {
        const [card, jsonPath, jsonData, generateWarnings] =
          AdaptiveCardGenerator.generateAdaptiveCard(item.item, true, 5);
        if (jsonPath !== "$" && card.body && card.body[0] && (card.body[0] as any).$data) {
          (card.body[0] as any).$data = `\${${jsonPath}}`;
        }
        const cardFilePath = path.join(adaptiveCardsFolderPath, `${name}.json`);
        const jsonDataPath = path.join(adaptiveCardsFolderPath, `${name}.data.json`);
        await fs.writeFile(cardFilePath, JSON.stringify(card, null, 2));
        await fs.writeFile(jsonDataPath, JSON.stringify(jsonData, null, 2));

        generateWarnings.forEach((w) => {
          warnings.push({
            type: WarningType.GenerateJsonDataFailed,
            // TODO: move message to package.nls.json file
            content: util.format(
              "Failed to create the adaptive card mock data for API '%s': %s. Mitigation: Not required but you can manually add it to the adaptiveCards folder.",
              item.item.operationId,
              w.content
            ),
            data: item.item.operationId,
          });
        });
      } catch (err) {
        warnings.push({
          type: WarningType.GenerateCardFailed,
          content: getLocalizedString(
            "core.copilotPlugin.scaffold.summary.warning.generate.ac.failed",
            item.item.operationId,
            err.message
          ),
          data: item.item.operationId,
        });
      }
    }
  }
  return warnings;
}

function filterSchema(schema: OpenAPIV3.SchemaObject): OpenAPIV3.SchemaObject {
  const filteredSchema: any = { type: schema.type };

  if (schema.description) {
    filteredSchema.description = schema.description;
  }

  if (schema.type === "object" && schema.properties) {
    filteredSchema.properties = {};
    filteredSchema.required = schema.required;
    for (const key in schema.properties) {
      const property = schema.properties[key] as OpenAPIV3.SchemaObject;
      if (property.type === "object") {
        filteredSchema.properties[key] = filterSchema(property as OpenAPIV3.SchemaObject);
        filteredSchema.required = schema.required;
      } else if (property.type === "array") {
        filteredSchema.properties[key] = {
          type: "array",
          items: filterSchema(property.items as OpenAPIV3.SchemaObject),
          description: property.description,
        };
      } else {
        filteredSchema.properties[key] = {
          type: property.type,
          description: property.description,
        };
      }
    }
  } else if (schema.type === "array" && schema.items) {
    filteredSchema.items = filterSchema(schema.items as OpenAPIV3.SchemaObject);
  }

  return filteredSchema;
}

async function updateActionForCustomApi(
  specItems: SpecObject[],
  language: string,
  chatFolder: string
): Promise<void> {
  if (commonLanguages.includes(language as ProgrammingLanguage)) {
    const actionsFilePath = path.join(chatFolder, "actions.json");
    const actions = [];

    for (const item of specItems) {
      const parameters: any = {
        type: "object",
        properties: {} as OpenAPIV3.SchemaObject,
        required: [],
      };

      const paramObject = item.item.parameters as OpenAPIV3.ParameterObject[];
      if (paramObject) {
        for (let i = 0; i < paramObject.length; i++) {
          const param = paramObject[i];
          const schema = param.schema as OpenAPIV3.SchemaObject;
          const paramType = param.in;

          if (!parameters.properties[paramType]) {
            parameters.properties[paramType] = {
              type: "object",
              properties: {},
              required: [],
            };
          }
          parameters.properties[paramType].properties[param.name] = filterSchema(schema);
          parameters.properties[paramType].properties[param.name].description =
            param.description ?? "";
          if (param.required) {
            parameters.properties[paramType].required.push(param.name);
            if (!parameters.required.includes(paramType)) {
              parameters.required.push(paramType);
            }
          }
        }
      }

      const requestBody = item.item.requestBody as OpenAPIV3.RequestBodyObject;
      if (requestBody) {
        const content = requestBody.content;
        if (content) {
          const contentSchema = content["application/json"].schema as OpenAPIV3.SchemaObject;
          if (Object.keys(contentSchema).length !== 0) {
            parameters.properties["body"] = filterSchema(contentSchema);
            parameters.properties["body"].description = requestBody.description ?? "";
            if (requestBody.required) {
              parameters.required.push("body");
            }
          }
        }
      }

      actions.push({
        name: item.item.operationId,
        description: item.item.description ?? item.item.summary,
        parameters: parameters,
      });
    }

    await fs.writeFile(actionsFilePath, JSON.stringify(actions, null, 2));
  }
}

const ActionCode = {
  javascript: `
app.ai.action("{{operationId}}", async (context, state, parameter) => {
  const client = await api.getClient();
  // Add authentication configuration for the client
  const apiPath = client.paths["{{pathUrl}}"];
  if (apiPath && apiPath.{{method}}) {
    const result = await apiPath.{{method}}(parameter.path, parameter.body, {
      params: parameter.query,
    });
    if (!result || !result.data) {
      throw new Error("Get empty result from api call.");
    }
    const cardName = "{{operationId}}".replace(/[^a-zA-Z0-9]/g, "_");
    const cardTemplatePath = path.join(__dirname, '../adaptiveCards', cardName + '.json');
    const isTeamsChannel = context.activity.channelId === Channels.Msteams;
    if (await fs.exists(cardTemplatePath)){
      const card = generateAdaptiveCard(cardTemplatePath, result);
      await context.sendActivity({
        attachments: [card],
        ...(isTeamsChannel ? { channelData: { feedbackLoopEnabled: true }} : {}),
        entities: [
          {
            type: "https://schema.org/Message",
            "@type": "Message",
            "@context": "https://schema.org",
            additionalType: ["AIGeneratedContent"], // AI Generated label
          },
        ]
      });
    }
    else {
      await context.sendActivity(JSON.stringify(result.data));
    }
  } else {
    await context.sendActivity("no result");
  }
  return "result";
});
  `,
  typescript: `
app.ai.action("{{operationId}}", async (context: TurnContext, state: ApplicationTurnState, parameter: any) => {
  const client = await api.getClient();
  // Add authentication configuration for the client
  const apiPath = client.paths["{{pathUrl}}"];
  if (apiPath && apiPath.{{method}}) {
    const result = await apiPath.{{method}}(parameter.path, parameter.body, {
      params: parameter.query,
    });
    if (!result || !result.data) {
      throw new Error("Get empty result from api call.");
    }
    const cardName = "{{operationId}}".replace(/[^a-zA-Z0-9]/g, "_");
    const cardTemplatePath = path.join(__dirname, '../adaptiveCards', cardName + '.json');
    const isTeamsChannel = context.activity.channelId === Channels.Msteams;
    if (await fs.exists(cardTemplatePath)){
      const card = generateAdaptiveCard(cardTemplatePath, result);
      await context.sendActivity({
        attachments: [card],
        ...(isTeamsChannel ? { channelData: { feedbackLoopEnabled: true }} : {}),
        entities: [
          {
            type: "https://schema.org/Message",
            "@type": "Message",
            "@context": "https://schema.org",
            additionalType: ["AIGeneratedContent"], // AI Generated label
          },
        ]
      });
    }
    else {
      await context.sendActivity(JSON.stringify(result.data));
    }
  } else {
    await context.sendActivity("no result");
  }
  return "result";
});
  `,
  python: `
@bot_app.ai.action("{{operationId}}")
async def {{operationId}}(
  context: ActionTurnContext[Dict[str, Any]],
  state: AppTurnState,
):
  parameters = context.data
  path = parameters.get("path", {})
  body = parameters.get("body", None)
  query = parameters.get("query", {})
  resp = client.{{operationId}}(**path, json=body, _headers={}, _params=query, _cookies={})

  if resp.status_code != 200:
    await context.send_activity(resp.reason)
  else:
    card_template_path = os.path.join(current_dir, 'adaptiveCards/{{operationId}}.json')
    if not os.path.exists(card_template_path):
      json_resoponse_str = resp.text
      await context.send_activity(json_resoponse_str)
    else:
      with open(card_template_path) as card_template_file:
        adaptive_card_template = card_template_file.read()

      renderer = AdaptiveCardRenderer(adaptive_card_template)

      json_resoponse_str = resp.text
      rendered_card_str = renderer.render(json_resoponse_str)
      rendered_card_json = json.loads(rendered_card_str)
      card = CardFactory.adaptive_card(rendered_card_json)
      isTeamsChannel = context.activity.channel_id == "msteams"
      message = MessageFactory.attachment(card)
      message.entities = [
        {
          "type": "https://schema.org/Message",
          "@type": "Message",
          "@context": "https://schema.org",
          "additionalType": ["AIGeneratedContent"],
        },
      ]
      message.channel_data = {
        "feedbackLoopEnabled": isTeamsChannel
      }
      
      await context.send_activity(message)
  return "success"
  `,
  cs: `
        [Action("{{operationId}}")]
        public async Task<string> {{functionName}}Async([ActionTurnContext] ITurnContext turnContext, [ActionTurnState] TurnState turnState, [ActionParameters] Dictionary<string, object> args)
        {
            try
            {
                RequestParams requestParam = ParseRequestParams(args);

                var response = await Client.CallAsync("{{apiPath}}", Method.{{apiMethod}}, requestParam);
                var data = response.Content;

                var cardTemplatePath = "./adaptiveCards/{{operationId}}.json";
                if (File.Exists(cardTemplatePath)) {
                    var message = RenderCardToMessage(cardTemplatePath, data);
                    await turnContext.SendActivityAsync(message);
                }
                else
                {
                    await turnContext.SendActivityAsync(data);
                }
            }
            catch (Exception ex) {
                await turnContext.SendActivityAsync("Failed to call API with error:  " + ex.Message);
            }

            return "complete";
        }`,
};

const AuthCode = {
  javascript: {
    actionCode: `addAuthConfig(client);`,
    actionPlaceholder: `// Add authentication configuration for the client`,
  },
  typescript: {
    actionCode: `addAuthConfig(client);`,
    actionPlaceholder: `// Add authentication configuration for the client`,
  },
};

async function updateCodeForCustomApi(
  specItems: SpecObject[],
  language: string,
  destinationPath: string,
  openapiSpecFileName: string,
  needAuth: boolean
): Promise<void> {
  if (language === ProgrammingLanguage.JS || language === ProgrammingLanguage.TS) {
    const codeTemplate =
      ActionCode[language === ProgrammingLanguage.JS ? "javascript" : "typescript"];
    const appFolderPath = path.join(destinationPath, "src", "app");

    const actionsCode = [];
    const authCodeTemplate =
      AuthCode[language === ProgrammingLanguage.JS ? "javascript" : "typescript"];
    for (const item of specItems) {
      const auth = item.auth;
      const code = codeTemplate
        .replace(authCodeTemplate.actionPlaceholder, auth ? authCodeTemplate.actionCode : "")
        .replace(/{{operationId}}/g, item.item.operationId!)
        .replace(/{{pathUrl}}/g, item.pathUrl)
        .replace(/{{method}}/g, item.method);
      actionsCode.push(code);
    }

    // Update code in app file
    const indexFilePath = path.join(
      appFolderPath,
      language === ProgrammingLanguage.JS ? "app.js" : "app.ts"
    );
    const indexFileContent = (await fs.readFile(indexFilePath)).toString();
    const updateIndexFileContent = indexFileContent
      .replace("{{OPENAPI_SPEC_PATH}}", openapiSpecFileName)
      .replace("// Replace with action code", actionsCode.join("\n"));
    await fs.writeFile(indexFilePath, updateIndexFileContent);
  } else if (language === ProgrammingLanguage.PY) {
    // Update code in bot.py
    const actionsCode = [];
    const codeTemplate = ActionCode["python"];
    for (const item of specItems) {
      const code = codeTemplate
        .replace(/{{operationId}}/g, item.item.operationId!)
        .replace(/{{pathUrl}}/g, item.pathUrl)
        .replace(/{{method}}/g, item.method);
      actionsCode.push(code);
    }

    const botFilePath = path.join(destinationPath, "src", "bot.py");
    const botFileContent = (await fs.readFile(botFilePath)).toString();
    const updateBotFileContent = botFileContent
      .replace("{{OPENAPI_SPEC_PATH}}", openapiSpecFileName)
      .replace("# Replace with action code", actionsCode.join("\n"));
    await fs.writeFile(botFilePath, updateBotFileContent);
  } else if (language === ProgrammingLanguage.CSharp) {
    const actionsCode = [];
    const codeTemplate = ActionCode["cs"];
    for (const item of specItems) {
      const code = codeTemplate
        .replace(/{{operationId}}/g, item.item.operationId!)
        .replace(/{{apiPath}}/g, item.pathUrl)
        .replace(/{{apiMethod}}/g, Utils.updateFirstLetter(item.method))
        .replace(/{{functionName}}/g, Utils.updateFirstLetter(item.item.operationId!));
      actionsCode.push(code);
    }

    const apiActionCsFilePath = path.join(destinationPath, "APIActions.cs");
    const apiActionCsFileContent = (await fs.readFile(apiActionCsFilePath)).toString();
    const updateApiActionCsFileContent = apiActionCsFileContent
      .replace("{{OPENAPI_SPEC_PATH}}", "apiSpecificationFile/" + openapiSpecFileName)
      .replace("// Replace with action code", actionsCode.join("\n"));
    await fs.writeFile(apiActionCsFilePath, updateApiActionCsFileContent);

    const files = await fs.readdir(destinationPath);
    const projectFileName = files.find((file) => file.endsWith(".csproj"));
    const projectFilePath = path.join(destinationPath, projectFileName!);
    const projectFileContent = (await fs.readFile(projectFilePath)).toString();
    const updateProjectFileContent = projectFileContent.replace(
      /{{OPENAPI_SPEC_PATH}}/g,
      openapiSpecFileName
    );
    await fs.writeFile(projectFilePath, updateProjectFileContent);
  }
}

export async function updateForCustomApi(
  spec: OpenAPIV3.Document,
  language: string,
  destinationPath: string,
  openapiSpecFileName: string
): Promise<WarningResult[]> {
  const warnings: WarningResult[] = [];
  let chatFolder = path.join(destinationPath, "src", "prompts", "chat");
  if (language === ProgrammingLanguage.CSharp) {
    chatFolder = path.join(destinationPath, "prompts", "Chat");
  }
  await fs.ensureDir(chatFolder);

  // 1. update prompt folder
  await updatePromptForCustomApi(spec, language, chatFolder);

  const [specItems, needAuth] = parseSpec(spec);

  // 2. update adaptive card folder
  const generateWarnings = await updateAdaptiveCardForCustomApi(
    specItems,
    language,
    destinationPath
  );

  warnings.push(...generateWarnings);

  // 3. update actions file
  await updateActionForCustomApi(specItems, language, chatFolder);

  // 4. update code
  await updateCodeForCustomApi(specItems, language, destinationPath, openapiSpecFileName, needAuth);

  // 5. add prompt suggestions
  const manifestPath = path.join(destinationPath, AppPackageFolderName, ManifestTemplateFileName);
  await updatePromptSuggestions(specItems, manifestPath);

  return warnings;
}

async function updatePromptSuggestions(specItems: SpecObject[], manifestPath: string) {
  const descriptions: string[] = specItems
    .map((item) => item.item.summary ?? item.item.description)
    .filter((item): item is string => item !== undefined)
    .slice(0, 10);

  const manifestRes = await manifestUtils._readAppManifest(manifestPath);
  if (manifestRes.isOk()) {
    const manifest = manifestRes.value;
    manifest.bots![0].commandLists = [
      {
        scopes: ["personal"],
        commands: [
          {
            title: "Hello, how can you help me?",
            description: "How can you help me?",
          },
          {
            title: "How to build apps with TTK?",
            description: "How can I develop apps with Teams Toolkit?",
          },
          ...descriptions.map((des) => {
            return {
              title: des.slice(0, 32),
              description: des.slice(0, 128),
            };
          }),
        ],
      },
    ];

    await manifestUtils._writeAppManifest(manifest, manifestPath);
  } else {
    throw manifestRes.error;
  }
}

export function getEnvName(authName: string): string {
  return Utils.getSafeRegistrationIdEnvName(`${authName}_${ConstantString.RegistrationIdPostfix}`);
}

export async function copyKiotaFolder(specPath: string, projectPath: string): Promise<void> {
  const originKiotaFolder = path.join(path.dirname(specPath), "..", ".kiota");
  if (!(await fs.pathExists(originKiotaFolder))) {
    return;
  }

  const destinationKiotaFolder = path.join(projectPath, ".kiota");
  if (await fs.pathExists(destinationKiotaFolder)) {
    return;
  }

  await fs.ensureDir(destinationKiotaFolder);
  await fs.copy(originKiotaFolder, destinationKiotaFolder, { recursive: true });
  return;
}
