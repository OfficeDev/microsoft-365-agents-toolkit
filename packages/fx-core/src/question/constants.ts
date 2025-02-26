// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { OptionItem } from "@microsoft/teamsfx-api";
import { FeatureFlags, featureFlagManager } from "../common/featureFlags";
import { getLocalizedString } from "../common/localizeUtils";

export enum QuestionNames {
  Scratch = "scratch",
  SctatchYes = "scratch-yes",
  AppName = "app-name",
  Folder = "folder",
  ProjectPath = "projectPath",
  ProgrammingLanguage = "programming-language",
  ProjectType = "project-type",
  Capabilities = "capabilities",
  BotTrigger = "bot-host-type-trigger",
  Runtime = "runtime",
  SPFxSolution = "spfx-solution",
  SPFxInstallPackage = "spfx-install-latest-package",
  SPFxFramework = "spfx-framework-type",
  SPFxWebpartName = "spfx-webpart-name",
  SPFxWebpartDesc = "spfx-webpart-desp",
  SPFxFolder = "spfx-folder",
  OfficeAddinFolder = "addin-project-folder",
  OfficeAddinManifest = "addin-project-manifest",
  OfficeAddinTemplate = "addin-template-select",
  OfficeAddinHost = "addin-host",
  OfficeAddinImport = "addin-import",
  OfficeAddinFramework = "office-addin-framework-type",
  Samples = "samples",
  ReplaceContentUrl = "replaceContentUrl",
  ReplaceWebsiteUrl = "replaceWebsiteUrl",
  ReplaceBotIds = "replaceBotIds",
  SafeProjectName = "safeProjectName",
  RepalceTabUrl = "tdp-tab-url",
  ValidateMethod = "validate-method",
  AppPackagePath = "appPackagePath",
  FromExistingApi = "from-existing-api", // group name for creating an App from existing api
  ApiSpecLocation = "openapi-spec-location",
  ApiOperation = "api-operation",
  ApiPluginManifestPath = "external-api-plugin-manifest-path", // manifest path for creating project from existing plugin manifest. Use in Kiota integration, etc.
  MeArchitectureType = "me-architecture",
  ApiSpecApiKey = "api-key",
  ApiSpecApiKeyConfirm = "api-key-confirm",
  ApiAuth = "api-auth",
  OauthClientSecret = "oauth-client-secret",
  OauthClientId = "oauth-client-id",
  OauthConfirm = "oauth-confirm",

  CustomCopilotRag = "custom-copilot-rag",
  CustomCopilotAssistant = "custom-copilot-agent",
  LLMService = "llm-service",
  OpenAIKey = "openai-key",
  OpenAIEmbeddingModel = "openai-embedding-model",
  AzureOpenAIKey = "azure-openai-key",
  AzureOpenAIEndpoint = "azure-openai-endpoint",
  AzureOpenAIDeploymentName = "azure-openai-deployment-name",
  AzureOpenAIEmbeddingDeploymentName = "azure-openai-embedding-deployment-name",
  AzureAISearchApiKey = "azure-ai-search-api-key",
  AzureAISearchEndpoint = "azure-ai-search-endpoint",

  Features = "features",
  Env = "env",
  SourceEnvName = "sourceEnvName",
  TargetEnvName = "targetEnvName",
  TargetResourceGroupName = "targetResourceGroupName",
  NewResourceGroupName = "newResourceGroupName",
  NewResourceGroupLocation = "newResourceGroupLocation",
  NewTargetEnvName = "newTargetEnvName",
  ExistingTabEndpoint = "existing-tab-endpoint",
  TeamsAppManifestFilePath = "manifest-path",
  LocalTeamsAppManifestFilePath = "local-manifest-path",
  AadAppManifestFilePath = "manifest-file-path",
  TeamsAppPackageFilePath = "app-package-file-path",
  ConfirmManifest = "confirmManifest",
  ConfirmLocalManifest = "confirmLocalManifest",
  ConfirmAadManifest = "confirmAadManifest",
  OutputZipPathParamName = "output-zip-path",
  OutputManifestParamName = "output-manifest-path",
  OutputFolderParamName = "output-folder",
  M365Host = "m365-host",

  ManifestPath = "manifest-path",
  ManifestId = "manifest-id",
  TeamsAppId = "teams-app-id",
  TitleId = "title-id",
  UserEmail = "email",

  UninstallMode = "mode",
  UninstallModeManifestId = "manifest-id",
  UninstallModeEnv = "env",
  UninstallModeTitleId = "title-id",
  UninstallOptions = "options",
  UninstallOptionM365 = "m365-app",
  UninstallOptionTDP = "app-registration",
  UninstallOptionBot = "bot-framework-registration",

  collaborationAppType = "collaborationType",
  DestinationApiSpecFilePath = "destination-api-spec-location",

  SyncManifest = "sync-manifest",
  ApiPluginType = "api-plugin-type",
  WithPlugin = "with-plugin",
  ImportPlugin = "import-plugin",
  PluginManifestFilePath = "plugin-manifest-path",
  PluginOpenApiSpecFilePath = "plugin-opeanapi-spec-path",
  KnowledgeSource = "knowledge-source",

  AuthName = "auth-name",
  TemplateName = "template-name",

  EmbeddedKnowledgeFiles = "embedded-knowledge-files",
  OAuthAuthorizationUrl = "oauth-authorization-url",
  OAuthTokenUrl = "oauth-token-url",
  OAuthRefreshUrl = "oauth-refresh-url",
  OAuthScope = "oauth-scope",
  OauthPKCE = "oauth-pkce",
  ApiKeyIn = "api-key-in",
  ApiKeyName = "api-key-name",
}

export enum ProjectTypeGroup {
  AIAgent = "AI Agent",
  M365Apps = "Apps for Microsoft 365",
}

export const AppNamePattern =
  '^(?=(.*[\\da-zA-Z]){2})[a-zA-Z][^"<>:\\?/*&|\u0000-\u001F]*[^"\\s.<>:\\?/*&|\u0000-\u001F]$';

export enum CliQuestionName {
  Capability = "capability",
}

export enum ProgrammingLanguage {
  JS = "javascript",
  TS = "typescript",
  CSharp = "csharp",
  PY = "python",
  Common = "common",
  None = "none",
}

export const apiPluginApiSpecOptionId = "api-spec";
export const capabilitiesHavePythonOption = [
  "custom-copilot-basic",
  "custom-copilot-rag-azureAISearch",
  "custom-copilot-rag-customize",
  "custom-copilot-agent-new",
  "custom-copilot-agent-assistants-api",
  "custom-copilot-rag-customApi",
];

export class ScratchOptions {
  static yes(): OptionItem {
    return {
      id: "yes",
      label: getLocalizedString("core.ScratchOptionYes.label"),
      detail: getLocalizedString("core.ScratchOptionYes.detail"),
    };
  }
  static no(): OptionItem {
    return {
      id: "no",
      label: getLocalizedString("core.ScratchOptionNo.label"),
      detail: getLocalizedString("core.ScratchOptionNo.detail"),
    };
  }
  static all(): OptionItem[] {
    return [ScratchOptions.yes(), ScratchOptions.no()];
  }
}

export class AddAuthActionAuthTypeOptions {
  static apiKey(): OptionItem {
    return {
      id: "api-key",
      label: "API Key",
    };
  }

  static bearerToken(): OptionItem {
    return {
      id: "bearer-token",
      label: "API Key (Bearer Token Auth)",
    };
  }

  static oauth(): OptionItem {
    return {
      id: "oauth",
      label: "OAuth",
    };
  }

  static microsoftEntra(): OptionItem {
    return {
      id: "microsoft-entra",
      label: "Microsoft Entra",
    };
  }

  static all(): OptionItem[] {
    return [
      AddAuthActionAuthTypeOptions.bearerToken(),
      AddAuthActionAuthTypeOptions.apiKey(),
      AddAuthActionAuthTypeOptions.oauth(),
      AddAuthActionAuthTypeOptions.microsoftEntra(),
    ];
  }
}

export enum HostType {
  AppService = "app-service",
  Functions = "azure-functions",
}

export const NotificationTriggers = {
  HTTP: "http",
  TIMER: "timer",
} as const;

export type NotificationTrigger = typeof NotificationTriggers[keyof typeof NotificationTriggers];

export interface HostTypeTriggerOptionItem extends OptionItem {
  hostType: HostType;
  triggers?: NotificationTrigger[];
}

export enum SPFxVersionOptionIds {
  installLocally = "true",
  globalPackage = "false",
}

export const recommendedLocations = [
  "South Africa North",
  "Australia East",
  "Central India",
  "East Asia",
  "Japan East",
  "Korea Central",
  "Southeast Asia",
  "Canada Central",
  "France Central",
  "Germany West Central",
  "Italy North",
  "North Europe",
  "Norway East",
  "Poland Central",
  "Sweden Central",
  "Switzerland North",
  "UK South",
  "West Europe",
  "Israel Central",
  "Qatar Central",
  "UAE North",
  "Brazil South",
  "Central US",
  "East US",
  "East US 2",
  "South Central US",
  "West US 2",
  "West US 3",
];

export class TeamsAppValidationOptions {
  static schema(): OptionItem {
    return {
      id: "validateAgainstSchema",
      label: getLocalizedString("core.selectValidateMethodQuestion.validate.schemaOption"),
    };
  }
  static package(): OptionItem {
    return {
      id: "validateAgainstPackage",
      label: getLocalizedString("core.selectValidateMethodQuestion.validate.appPackageOption"),
      detail: getLocalizedString(
        "core.selectValidateMethodQuestion.validate.appPackageOptionDescription"
      ),
    };
  }
  static testCases(): OptionItem {
    return {
      id: "validateWithTestCases",
      label: getLocalizedString("core.selectValidateMethodQuestion.validate.testCasesOption"),
      detail: getLocalizedString(
        "core.selectValidateMethodQuestion.validate.testCasesOptionDescription"
      ),
    };
  }
}

export enum HubTypes {
  teams = "teams",
  outlook = "outlook",
  office = "office",
}

export class HubOptions {
  static teams(): OptionItem {
    return {
      id: "teams",
      label: "Teams",
    };
  }
  static outlook(): OptionItem {
    return {
      id: "outlook",
      label: "Outlook",
    };
  }
  static office(): OptionItem {
    return {
      id: "office",
      label: "the Microsoft 365 app",
    };
  }
  static all(): OptionItem[] {
    return [this.teams(), this.outlook(), this.office()];
  }
}

export class KnowledgeSourceOptions {
  static webSearch(): OptionItem {
    return {
      id: "web-search",
      label: getLocalizedString("core.createProjectQuestion.capability.knowledgeWebSearch.label"),
      detail: getLocalizedString("core.createProjectQuestion.capability.knowledgeWebSearch.detail"),
    };
  }

  static oneDriveSharePoint(): OptionItem {
    return {
      id: "oneDrive-sharePoint",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.knowledgeOneDriveSharePoint.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.knowledgeOneDriveSharePoint.detail"
      ),
    };
  }

  static graphConnector(): OptionItem {
    return {
      id: "graph-connector",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.knowledgeGraphConnector.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.knowledgeGraphConnector.detail"
      ),
    };
  }

  static embeddedKnowledge(): OptionItem {
    return {
      id: "embedded-knowledge",
      label: getLocalizedString(
        "core.createProjectQuestion.capability.knowledgeEmbeddedKnowledge.label"
      ),
      detail: getLocalizedString(
        "core.createProjectQuestion.capability.knowledgeEmbeddedKnowledge.detail"
      ),
    };
  }

  static all(): OptionItem[] {
    const items: OptionItem[] = [
      KnowledgeSourceOptions.webSearch(),
      KnowledgeSourceOptions.oneDriveSharePoint(),
      KnowledgeSourceOptions.embeddedKnowledge(),
    ];
    return items;
  }

  static allWithFeatureFlags(): OptionItem[] {
    const items: OptionItem[] = [
      KnowledgeSourceOptions.webSearch(),
      KnowledgeSourceOptions.oneDriveSharePoint(),
    ];
    if (featureFlagManager.getBooleanValue(FeatureFlags.BuilderAPIEnabled)) {
      items.push(KnowledgeSourceOptions.embeddedKnowledge());
    }
    return items;
  }
}
