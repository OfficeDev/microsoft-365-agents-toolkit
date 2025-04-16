// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Inputs } from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { convertToAlphanumericOnly } from "../../../common/stringUtils";
import { QuestionNames } from "../../../question/constants";
import { LocalCrypto } from "../../../core/crypto";
import os from "os";

export function getTemplateReplaceMap(inputs: Inputs): { [key: string]: string } {
  const appName = inputs[QuestionNames.AppName] as string;
  const safeProjectName =
    inputs[QuestionNames.SafeProjectName] ?? convertToAlphanumericOnly(appName);
  const targetFramework = inputs.targetFramework;
  const placeProjectFileInSolutionDir = inputs.placeProjectFileInSolutionDir === "true";
  const llmService: string | undefined = inputs[QuestionNames.LLMService];
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
    TargetFramework: targetFramework ?? "net8.0",
    PlaceProjectFileInSolutionDir: placeProjectFileInSolutionDir ? "true" : "",
    SafeProjectName: safeProjectName,
    SafeProjectNameLowerCase: safeProjectName.toLocaleLowerCase(),
    enableTestToolByDefault: featureFlagManager.getBooleanValue(FeatureFlags.TestTool)
      ? "true"
      : "",
    enableMETestToolByDefault: featureFlagManager.getBooleanValue(FeatureFlags.METestTool)
      ? "true"
      : "",
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
    openAIEmbeddingModel: openAIEmbeddingModel ?? "",
    isNewProjectTypeEnabled: featureFlagManager.getBooleanValue(FeatureFlags.NewProjectType)
      ? "true"
      : "",
    NewProjectTypeName: process.env.TEAMSFX_NEW_PROJECT_TYPE_NAME ?? "TeamsApp",
    NewProjectTypeExt: process.env.TEAMSFX_NEW_PROJECT_TYPE_EXTENSION ?? "ttkproj",
    CEAEnabled: featureFlagManager.getBooleanValue(FeatureFlags.CEAEnabled) ? "true" : "",
    EmbeddedKnowledgeEnabled: featureFlagManager.getBooleanValue(
      FeatureFlags.EmbeddedKnowledgeEnabled
    )
      ? "true"
      : "",
    ShareEnabled: featureFlagManager.getBooleanValue(FeatureFlags.ShareEnabled) ? "true" : "",
    SensitivityLabelEnabled: featureFlagManager.getBooleanValue(
      FeatureFlags.SensitivityLabelEnabled
    )
      ? "true"
      : "",
    SandBoxedTeam: featureFlagManager.getBooleanValue(FeatureFlags.SandBoxedTeam) ? "true" : "",
    pathDelimiter: os.platform() === "win32" ? ";" : ":",
  };
}
