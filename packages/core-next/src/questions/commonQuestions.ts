// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as path from "path";
import * as os from "os";
import type {
  SingleSelectQuestion,
  TextInputQuestion,
  FolderQuestion,
  SingleFileOrInputQuestion,
  MultiSelectQuestion,
} from "../api/qm/question";
import type { OptionItem, Inputs } from "../api/types";
import type { QuestionSpec, Language } from "../templates/types";
import { Platform } from "../api/constants";
import { QuestionNames } from "./questionNames";

/**
 * App name validation pattern: starts with letter, at least 2 alphanumeric,
 * max 30 chars, no special chars.
 */
const appNamePattern =
  '^(?=(.*[\\da-zA-Z]){2})[a-zA-Z][^"<>:\\?/*&|\\u0000-\\u001F]*[^"\\s.<>:\\?/*&|\\u0000-\\u001F]$';

/**
 * Language display label mapping.
 */
const languageLabelMap: Record<string, string> = {
  typescript: "TypeScript",
  javascript: "JavaScript",
  python: "Python",
  csharp: "C#",
  common: "None",
};

/**
 * Create a "Project name" text question.
 */
export function projectNameQuestion(): TextInputQuestion {
  return {
    type: "text",
    name: QuestionNames.projectName,
    title: "Application name",
    default: "my-app",
    validation: {
      pattern: appNamePattern,
      maxLength: 30,
    },
  };
}

/**
 * Create a "Destination folder" question.
 * Defaults to "./" on CLI, ~/TeamsApps on VS Code.
 */
export function destinationFolderQuestion(): FolderQuestion {
  return {
    type: "folder",
    name: QuestionNames.destinationFolder,
    title: "Directory where the project folder will be created",
    default: (inputs: Inputs) =>
      inputs.platform === Platform.CLI ? "./" : path.join(os.homedir(), "TeamsApps"),
  };
}

/**
 * Create a "Programming language" single-select question from a list of supported languages.
 * If only one language is supported, the question is auto-skipped.
 */
export function languageQuestion(languages: Language[]): SingleSelectQuestion {
  const options: OptionItem[] = languages.map((lang) => ({
    id: lang,
    label: languageLabelMap[lang] ?? lang,
  }));
  return {
    type: "singleSelect",
    name: QuestionNames.language,
    title: "Programming language",
    staticOptions: options,
    default: options[0]?.id,
    skipSingleOption: true,
  };
}

/**
 * Create a "Project type" category selector from registered template categories.
 */
export function projectTypeQuestion(categories: OptionItem[]): SingleSelectQuestion {
  return {
    type: "singleSelect",
    name: QuestionNames.projectType,
    title: "Select a project type",
    staticOptions: categories,
    default: categories[0]?.id,
  };
}

/**
 * Create a "Select a template" single-select question from a list of descriptors.
 */
export function templateIdQuestion(templates: OptionItem[]): SingleSelectQuestion {
  return {
    type: "singleSelect",
    name: QuestionNames.templateId,
    title: "Select a template",
    staticOptions: templates,
    default: templates[0]?.id,
  };
}

/**
 * Create an "API spec path" file-or-text input question.
 */
export function apiSpecPathQuestion(): QuestionSpec {
  const question: SingleFileOrInputQuestion = {
    type: "singleFileOrText",
    name: QuestionNames.apiSpecPath,
    title: "Enter the path or URL to your OpenAPI specification",
    inputOptionItem: {
      id: "apiSpecUrl",
      label: "Enter a URL",
    },
    inputBoxConfig: {
      type: "innerText",
      name: `${QuestionNames.apiSpecPath}-url`,
      title: "OpenAPI specification URL",
      placeholder: "https://example.com/openapi.json",
    },
    filters: { OpenAPI: ["json", "yaml", "yml"] },
  };
  return { question };
}

/**
 * Create an "API operations" multi-select question.
 */
export function apiOperationsQuestion(): QuestionSpec {
  const question: MultiSelectQuestion = {
    type: "multiSelect",
    name: QuestionNames.apiOperations,
    title: "Select the API operations to include",
    staticOptions: [],
    dynamicOptions: () => Promise.resolve([]),
    validation: {
      minItems: 1,
    },
  };
  return { question, dependsOn: QuestionNames.apiSpecPath };
}

/**
 * Create an "LLM provider" selector (Azure OpenAI vs OpenAI).
 */
export function llmProviderQuestion(): QuestionSpec {
  const question: SingleSelectQuestion = {
    type: "singleSelect",
    name: QuestionNames.llmProvider,
    title: "Select your LLM provider",
    staticOptions: [
      { id: "azure-openai", label: "Azure OpenAI" },
      { id: "openai", label: "OpenAI" },
    ],
    default: "azure-openai",
  };
  return { question };
}

/**
 * Create "Azure OpenAI key" text question (conditional on llmProvider = azure-openai).
 */
export function azureOpenAiKeyQuestion(): QuestionSpec {
  const question: TextInputQuestion = {
    type: "text",
    name: QuestionNames.azureOpenAiKey,
    title: "Enter your Azure OpenAI API key",
    password: true,
  };
  return {
    question,
    dependsOn: QuestionNames.llmProvider,
    condition: (inputs: Inputs) => inputs[QuestionNames.llmProvider] === "azure-openai",
  };
}

/**
 * Create "Azure OpenAI endpoint" text question (conditional on llmProvider = azure-openai).
 */
export function azureOpenAiEndpointQuestion(): QuestionSpec {
  const question: TextInputQuestion = {
    type: "text",
    name: QuestionNames.azureOpenAiEndpoint,
    title: "Enter your Azure OpenAI endpoint URL",
    placeholder: "https://your-resource.openai.azure.com",
  };
  return {
    question,
    dependsOn: QuestionNames.llmProvider,
    condition: (inputs: Inputs) => inputs[QuestionNames.llmProvider] === "azure-openai",
  };
}

/**
 * Create "OpenAI API key" text question (conditional on llmProvider = openai).
 */
export function openAiKeyQuestion(): QuestionSpec {
  const question: TextInputQuestion = {
    type: "text",
    name: QuestionNames.openAiKey,
    title: "Enter your OpenAI API key",
    password: true,
  };
  return {
    question,
    dependsOn: QuestionNames.llmProvider,
    condition: (inputs: Inputs) => inputs[QuestionNames.llmProvider] === "openai",
  };
}

/**
 * Create "OAuth client ID" text question.
 */
export function oauthClientIdQuestion(): QuestionSpec {
  const question: TextInputQuestion = {
    type: "text",
    name: QuestionNames.oauthClientId,
    title: "Enter your OAuth client ID",
  };
  return { question };
}

/**
 * Create "OAuth client secret" text question.
 */
export function oauthClientSecretQuestion(): QuestionSpec {
  const question: TextInputQuestion = {
    type: "text",
    name: QuestionNames.oauthClientSecret,
    title: "Enter your OAuth client secret",
    password: true,
  };
  return { question, dependsOn: QuestionNames.oauthClientId };
}

/**
 * Create "Graph Connector display name" text question.
 */
export function graphConnectorNameQuestion(): QuestionSpec {
  const question: TextInputQuestion = {
    type: "text",
    name: QuestionNames.graphConnectorName,
    title: "Enter a display name for your Graph Connector",
    default: "My Graph Connector",
  };
  return { question };
}

/**
 * Create "Graph Connector connection ID" text question.
 */
export function graphConnectorConnectionIdQuestion(): QuestionSpec {
  const question: TextInputQuestion = {
    type: "text",
    name: QuestionNames.graphConnectorConnectionId,
    title: "Enter a unique connection ID for your Graph Connector",
    default: "myGraphConnector",
    validation: {
      pattern: "^[a-zA-Z][a-zA-Z0-9]*$",
      maxLength: 32,
    },
  };
  return { question, dependsOn: QuestionNames.graphConnectorName };
}

/**
 * Create "MCP server URL" text question (for remote MCP templates).
 */
export function mcpServerUrlQuestion(): QuestionSpec {
  const question: TextInputQuestion = {
    type: "text",
    name: QuestionNames.mcpServerUrl,
    title: "Enter the MCP server URL",
    placeholder: "https://example.com/mcp",
  };
  return { question };
}

/**
 * Create "Foundry project endpoint" text question.
 */
export function foundryEndpointQuestion(): QuestionSpec {
  const question: TextInputQuestion = {
    type: "text",
    name: QuestionNames.foundryEndpoint,
    title: "Enter your AI Foundry project endpoint",
    placeholder: "https://your-project.services.ai.azure.com",
  };
  return { question };
}

/**
 * Create "Foundry agent ID" text question (conditional on foundryEndpoint).
 */
export function foundryAgentIdQuestion(): QuestionSpec {
  const question: TextInputQuestion = {
    type: "text",
    name: QuestionNames.foundryAgentId,
    title: "Enter your Foundry agent ID",
    placeholder: "agent-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  };
  return {
    question,
    dependsOn: QuestionNames.foundryEndpoint,
    condition: (inputs: Inputs) => !!inputs[QuestionNames.foundryEndpoint],
  };
}

/**
 * Create "Existing Office Add-in project folder" question (for MetaOS upgrade).
 */
export function officeAddinFolderQuestion(): QuestionSpec {
  const question: FolderQuestion = {
    type: "folder",
    name: QuestionNames.officeAddinFolder,
    title: "Select your existing Office Add-in project folder",
  };
  return { question };
}
