// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import { Context, FxError, Result, ok } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { merge } from "lodash";
import { TelemetryEvent, TelemetryProperty } from "../../common/telemetry";
import { BaseComponentInnerError } from "../error/componentError";
import { LogMessages, ProgressMessages, ProgressTitles } from "../messages";
import { ActionContext, ActionExecutionMW } from "../middleware/actionExecutionMW";
import {
  commonTemplateName,
  componentName,
  errorSource,
  sampleDefaultTimeoutInMs,
} from "./constant";
import {
  DownloadSampleApiLimitError,
  DownloadSampleNetworkError,
  FetchSampleInfoError,
  ScaffoldLocalTemplateError,
} from "./error";
import {
  SampleActionSeq,
  GeneratorAction,
  GeneratorActionName,
  GeneratorContext,
  TemplateActionSeq,
} from "./generatorAction";
import {
  convertToUrl,
  isApiLimitError,
  renderTemplateFileData,
  renderTemplateFileName,
} from "./utils";
import { featureFlagManager, FeatureFlags } from "../../common/featureFlags";
import { Utils } from "@microsoft/m365-spec-parser";
import { convertToAlphanumericOnly } from "../../common/stringUtils";

export class Generator {
  public static getDefaultVariables(
    appName: string,
    safeProjectNameFromVS?: string,
    targetFramework?: string,
    placeProjectFileInSolutionDir?: boolean,
    authData?: {
      authName: string;
      openapiSpecPath: string;
      registrationIdEnvName: string;
      authType?: string;
    }[],
    llmServiceData?: {
      llmService?: string;
      openAIKey?: string;
      azureOpenAIKey?: string;
      azureOpenAIEndpoint?: string;
      azureOpenAIDeploymentName?: string;
    }
  ): { [key: string]: any } {
    const safeProjectName = safeProjectNameFromVS ?? convertToAlphanumericOnly(appName);
    const apiKeyActionData = [];
    const oauthActionData = [];
    for (const auth of authData ?? []) {
      const safeRegistrationIdEnvName = Utils.getSafeRegistrationIdEnvName(
        auth.registrationIdEnvName
      );

      const actionData = {
        ApiSpecAuthName: auth.authName,
        ApiSpecPath: auth.openapiSpecPath,
        ApiSpecAuthRegistrationIdEnvName: safeRegistrationIdEnvName,
      };

      if (auth.authType === "apiKey") {
        apiKeyActionData.push(actionData);
      } else if (auth.authType === "oauth2") {
        oauthActionData.push(actionData);
      }
    }

    return {
      appName: appName,
      ProjectName: appName,
      TargetFramework: targetFramework ?? "net8.0",
      PlaceProjectFileInSolutionDir: placeProjectFileInSolutionDir ? "true" : "",
      SafeProjectName: safeProjectName,
      SafeProjectNameLowerCase: safeProjectName.toLocaleLowerCase(),
      ApiKey: apiKeyActionData,
      OAuth: oauthActionData,
      enableTestToolByDefault: featureFlagManager.getBooleanValue(FeatureFlags.TestTool)
        ? "true"
        : "",
      enableMETestToolByDefault: featureFlagManager.getBooleanValue(FeatureFlags.METestTool)
        ? "true"
        : "",
      useOpenAI: llmServiceData?.llmService === "llm-service-openai" ? "true" : "",
      useAzureOpenAI: llmServiceData?.llmService === "llm-service-azure-openai" ? "true" : "",
      openAIKey: llmServiceData?.openAIKey ?? "",
      azureOpenAIKey: llmServiceData?.azureOpenAIKey ?? "",
      azureOpenAIEndpoint: llmServiceData?.azureOpenAIEndpoint ?? "",
      azureOpenAIDeploymentName: llmServiceData?.azureOpenAIDeploymentName ?? "",
      isNewProjectTypeEnabled: featureFlagManager.getBooleanValue(FeatureFlags.NewProjectType)
        ? "true"
        : "",
      NewProjectTypeName: process.env.TEAMSFX_NEW_PROJECT_TYPE_NAME ?? "TeamsApp",
      NewProjectTypeExt: process.env.TEAMSFX_NEW_PROJECT_TYPE_EXTENSION ?? "ttkproj",
      CEAEnabled: featureFlagManager.getBooleanValue(FeatureFlags.CEAEnabled) ? "true" : "",
    };
  }
  @hooks([
    ActionExecutionMW({
      enableProgressBar: true,
      progressTitle: ProgressTitles.create,
      progressSteps: 1,
      componentName: componentName,
      errorSource: errorSource,
      enableTelemetry: true,
      telemetryEventName: TelemetryEvent.GenerateTemplate,
    }),
  ])
  public static async generateTemplate(
    ctx: Context,
    destinationPath: string,
    scenario: string,
    language?: string,
    actionContext?: ActionContext
  ): Promise<Result<undefined, FxError>> {
    const replaceMap = ctx.templateVariables ?? {};
    const lang = language ?? commonTemplateName;
    const generatorContext: GeneratorContext = {
      name: scenario,
      language: lang,
      destination: destinationPath,
      logProvider: ctx.logProvider,
      fileNameReplaceFn: (fileName, fileData) =>
        renderTemplateFileName(fileName, fileData, replaceMap)
          .replace(/\\/g, "/")
          .replace(`${scenario}/`, ""),
      fileDataReplaceFn: (fileName, fileData) =>
        renderTemplateFileData(fileName, fileData, replaceMap),
      filterFn: (fileName) => fileName.replace(/\\/g, "/").startsWith(`${scenario}/`),
      onActionError: templateDefaultOnActionError,
    };
    const templateName = `${scenario}-${lang}`;
    merge(actionContext?.telemetryProps, {
      [TelemetryProperty.TemplateName]: templateName,
    });

    await actionContext?.progressBar?.next(ProgressMessages.generateTemplate);
    ctx.logProvider.debug(`Downloading app template "${templateName}" to ${destinationPath}`);
    await this.generate(generatorContext, TemplateActionSeq);

    merge(actionContext?.telemetryProps, {
      [TelemetryProperty.Fallback]: generatorContext.fallback ? "true" : "false", // Track fallback cases.
    });
    return ok(undefined);
  }

  @hooks([
    ActionExecutionMW({
      enableProgressBar: true,
      progressTitle: ProgressTitles.create,
      progressSteps: 1,
      componentName: componentName,
      errorSource: errorSource,
      enableTelemetry: true,
      telemetryEventName: TelemetryEvent.GenerateSample,
    }),
  ])
  public static async generateSample(
    ctx: Context,
    destinationPath: string,
    sampleName: string,
    actionContext?: ActionContext
  ): Promise<Result<undefined, FxError>> {
    merge(actionContext?.telemetryProps, {
      [TelemetryProperty.SampleAppName]: sampleName,
      [TelemetryProperty.SampleDownloadDirectory]: "true",
    });
    // sample doesn't need replace function. Replacing projectId will be handled by core.

    const generatorContext: GeneratorContext = {
      name: sampleName,
      destination: destinationPath,
      logProvider: ctx.logProvider,
      timeoutInMs: sampleDefaultTimeoutInMs,
      onActionError: sampleDefaultOnActionError,
    };

    await actionContext?.progressBar?.next(ProgressMessages.generateSample(sampleName));
    ctx.logProvider.debug(`Downloading sample "${sampleName}" to ${destinationPath}`);
    await this.generate(generatorContext, SampleActionSeq);
    return ok(undefined);
  }

  public static async generate(
    context: GeneratorContext,
    actions: GeneratorAction[]
  ): Promise<void> {
    for (const action of actions) {
      try {
        await context.onActionStart?.(action, context);
        await action.run(context);
        await context.onActionEnd?.(action, context);
      } catch (e) {
        if (e instanceof Error) await context.onActionError(action, context, e);
      }
    }
  }
}

export function templateDefaultOnActionError(
  action: GeneratorAction,
  context: GeneratorContext,
  error: Error
): Promise<void> {
  switch (action.name) {
    case GeneratorActionName.ScaffoldRemoteTemplate:
      context.fallback = true;
      context.logProvider.debug(error.message);
      context.logProvider.info(LogMessages.getTemplateFromLocal);
      break;
    case GeneratorActionName.ScaffoldLocalTemplate:
      if (error instanceof BaseComponentInnerError) {
        return Promise.reject(error.toFxError());
      } else {
        context.logProvider.error(error.message);
        return Promise.reject(new ScaffoldLocalTemplateError(error).toFxError());
      }
    default:
      return Promise.reject(new Error(error.message));
  }
  return Promise.resolve();
}

export async function sampleDefaultOnActionError(
  action: GeneratorAction,
  context: GeneratorContext,
  error: Error
): Promise<void> {
  context.logProvider.error(error.message);
  if (error instanceof BaseComponentInnerError) throw error.toFxError();
  if (await fs.pathExists(context.destination)) {
    await fs.remove(context.destination);
  }
  switch (action.name) {
    case GeneratorActionName.FetchSampleInfo:
      throw new FetchSampleInfoError(error).toFxError();
    case GeneratorActionName.DownloadDirectory:
      const url = convertToUrl(context.sampleInfo!);
      if (isApiLimitError(error)) {
        throw new DownloadSampleApiLimitError(url, error).toFxError();
      } else {
        throw new DownloadSampleNetworkError(url, error).toFxError();
      }
    default:
      throw new Error(error.message);
  }
}
