// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Standard Mustache replacement variables derived from project inputs.
 *
 * Matches the fx-core replace map so templates remain fully compatible.
 */
export function getTemplateReplaceMap(inputs: {
  appName: string;
  safeProjectName?: string;
  targetFramework?: string;
  llmService?: "openai" | "azure-openai";
  openAIKey?: string;
  azureOpenAIKey?: string;
  azureOpenAIEndpoint?: string;
  azureOpenAIDeploymentName?: string;
  foundryEndpoint?: string;
  foundryAgentId?: string;
  gcName?: string;
  gcConnectionId?: string;
  [key: string]: unknown;
}): Record<string, string> {
  const safeName = inputs.safeProjectName ?? toSafeProjectName(inputs.appName);
  const map: Record<string, string> = {
    appName: inputs.appName,
    ProjectName: inputs.appName,
    SolutionName: inputs.appName,
    SafeProjectName: safeName,
    SafeProjectNameLowerCase: safeName.toLowerCase(),
    TargetFramework: inputs.targetFramework ?? "net8.0",
    PlaceProjectFileInSolutionDir: "",
    pathDelimiter: process.platform === "win32" ? ";" : ":",
    // C# project type placeholders (used in .csproj/.atkproj template paths)
    NewProjectTypeName: process.env.TEAMSFX_NEW_PROJECT_TYPE_NAME ?? "M365Agent",
    NewProjectTypeExt: process.env.TEAMSFX_NEW_PROJECT_TYPE_EXTENSION ?? "atkproj",
  };

  // LLM service toggles
  if (inputs.llmService === "openai") {
    map.useOpenAI = "true";
    if (inputs.openAIKey) map.openAIKey = inputs.openAIKey;
  } else if (inputs.llmService === "azure-openai") {
    map.useAzureOpenAI = "true";
    if (inputs.azureOpenAIKey) map.azureOpenAIKey = inputs.azureOpenAIKey;
    if (inputs.azureOpenAIEndpoint) map.azureOpenAIEndpoint = inputs.azureOpenAIEndpoint;
    if (inputs.azureOpenAIDeploymentName) {
      map.azureOpenAIDeploymentName = inputs.azureOpenAIDeploymentName;
    }
  }

  // Foundry
  if (inputs.foundryEndpoint) map.FoundryEndpoint = inputs.foundryEndpoint;
  if (inputs.foundryAgentId) map.FoundryAgentId = inputs.foundryAgentId;

  // Graph Connector
  if (inputs.gcName) map.gcName = inputs.gcName;
  if (inputs.gcConnectionId) map.gcConnectionId = inputs.gcConnectionId;

  return map;
}

/**
 * Convert a project name to alphanumeric-safe form.
 */
function toSafeProjectName(name: string): string {
  return name.replace(/[^a-zA-Z0-9]/g, "");
}
