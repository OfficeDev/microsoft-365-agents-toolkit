// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Inputs } from "@microsoft/teamsfx-api";
import os from "os";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { convertToAlphanumericOnly } from "../../../common/stringUtils";
import { LocalCrypto } from "../../../core/crypto";
import { QuestionNames } from "../../../question/constants";

export function getTemplateReplaceMap(inputs: Inputs): { [key: string]: string } {
  const appName = inputs[QuestionNames.AppName] as string;
  const safeProjectName =
    inputs[QuestionNames.SafeProjectName] ?? convertToAlphanumericOnly(appName);
  const solutionName = inputs[QuestionNames.SolutionName] ?? appName;
  const targetFramework = inputs.targetFramework;
  const placeProjectFileInSolutionDir = inputs.placeProjectFileInSolutionDir === "true";
  // Fall back to the CLI/UX default when the LLM service question was not asked
  // or omitted in non-interactive mode, so templates that depend on
  // {{#useAzureOpenAI}} blocks still emit a valid agent model.
  const llmService: string = inputs[QuestionNames.LLMService] ?? "llm-service-azure-openai";
  let openAIKey: string | undefined = inputs[QuestionNames.OpenAIKey];
  let azureOpenAIKey: string | undefined = inputs[QuestionNames.AzureOpenAIKey];
  let azureAISearchApiKey: string | undefined = inputs[QuestionNames.AzureAISearchApiKey];
  const azureOpenAIEndpoint: string | undefined = inputs[QuestionNames.AzureOpenAIEndpoint];
  const azureOpenAIDeploymentName: string | undefined =
    inputs[QuestionNames.AzureOpenAIDeploymentName];
  const azureAISearchEndpoint: string | undefined = inputs[QuestionNames.AzureAISearchEndpoint];
  const openAIEmbeddingModel: string | undefined = inputs[QuestionNames.OpenAIEmbeddingModel];
  const azureOpenAIEmbeddingDeploymentName: string | undefined =
    inputs[QuestionNames.AzureOpenAIEmbeddingDeploymentName];
  const gcName: string | undefined = inputs[QuestionNames.GCName];
  const gcConnectionId: string | undefined = inputs[QuestionNames.GCConnectionId];
  const foundryEndpoint: string | undefined = inputs[QuestionNames.FoundryEndpoint];
  const foundryAgentId: string | undefined = inputs[QuestionNames.FoundryAgentId];

  if (inputs.projectId !== undefined && (openAIKey || azureOpenAIKey)) {
    const cryptoProvider = new LocalCrypto(inputs.projectId);
    if (openAIKey) {
      const result = cryptoProvider.encrypt(openAIKey);
      openAIKey = (result as any).value;
    }
    if (azureOpenAIKey) {
      const result = cryptoProvider.encrypt(azureOpenAIKey);
      azureOpenAIKey = (result as any).value;
    }
    if (azureAISearchApiKey) {
      const result = cryptoProvider.encrypt(azureAISearchApiKey);
      azureAISearchApiKey = (result as any).value;
    }
  }

  return {
    appName: appName,
    ProjectName: appName,
    SolutionName: solutionName,
    TargetFramework: targetFramework ?? "net8.0",
    PlaceProjectFileInSolutionDir: placeProjectFileInSolutionDir ? "true" : "",
    SafeProjectName: safeProjectName,
    SafeProjectNameLowerCase: safeProjectName.toLocaleLowerCase(),
    useOpenAI: llmService === "llm-service-openai" ? "true" : "",
    useAzureOpenAI: llmService === "llm-service-azure-openai" ? "true" : "",
    openAIKey: openAIKey ?? "",
    originalOpenAIKey: inputs[QuestionNames.OpenAIKey] ?? "",
    azureOpenAIKey: azureOpenAIKey ?? "",
    originalAzureOpenAIKey: inputs[QuestionNames.AzureOpenAIKey] ?? "",
    azureAISearchApiKey: azureAISearchApiKey ?? "",
    originalAzureAISearchApiKey: inputs[QuestionNames.AzureAISearchApiKey] ?? "",
    azureOpenAIEndpoint: azureOpenAIEndpoint ?? "",
    azureOpenAIDeploymentName: azureOpenAIDeploymentName ?? "",
    azureOpenAIEmbeddingDeploymentName: azureOpenAIEmbeddingDeploymentName ?? "",
    azureAISearchEndpoint: azureAISearchEndpoint ?? "",
    gcName: gcName ?? "",
    gcConnectionId: gcConnectionId ?? "",
    FoundryEndpoint: foundryEndpoint ?? "",
    FoundryAgentId: foundryAgentId ?? "",
    openAIEmbeddingModel: openAIEmbeddingModel ?? "",
    NewProjectTypeName: process.env.TEAMSFX_NEW_PROJECT_TYPE_NAME ?? "M365Agent",
    NewProjectTypeExt: process.env.TEAMSFX_NEW_PROJECT_TYPE_EXTENSION ?? "atkproj",
    CEAEnabled: featureFlagManager.getBooleanValue(FeatureFlags.CEAEnabled) ? "true" : "",
    SensitivityLabelEnabled: featureFlagManager.getBooleanValue(
      FeatureFlags.SensitivityLabelEnabled
    )
      ? "true"
      : "",
    SandBoxedTeam: featureFlagManager.getBooleanValue(FeatureFlags.SandBoxedTeam) ? "true" : "",
    pathDelimiter: os.platform() === "win32" ? ";" : ":",
  };
}
