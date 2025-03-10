// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import { assign } from "lodash";
import { TOOLS, globalVars } from "./globalVars";
import { ProjectTypeResult } from "./projectTypeChecker";
import { maskSecret } from "./stringUtils";

export enum TelemetryProperty {
  TriggerFrom = "trigger-from",
  Component = "component",
  Components = "components",
  Feature = "feature",
  Hosting = "hosting",
  AppId = "appid",
  BotId = "botid",
  Success = "success",
  ErrorType = "error-type",
  ErrorCode = "error-code",
  ErrorCat = "error-cat",
  ErrorCat1 = "error-cat1",
  ErrorCat2 = "error-cat2",
  ErrorCat3 = "error-cat3",
  ErrorComponent = "error-component",
  ErrorInnerCode = "error-inner-code",
  ErrorMessage = "err-message",
  ErrorMethod = "error-method",
  ErrorName = "error-name",
  ErrorSource = "error-source",
  ErrorStack = "err-stack",
  ErrorData = "err-data",
  ErrorStage = "error-stage",
  SampleAppName = "sample-app-name",
  ProjectId = "project-id",
  NewProjectId = "new-project-id",
  CorrelationId = "correlation-id",
  Env = "env",
  CustomizeResourceGroupType = "customize-resource-group-type",
  EnvConfig = "env-config",
  Status = "status",
  HostType = "hostType",
  AzureResources = "azure-resources",
  Capabilities = "capabilities",
  ActivePlugins = "active-plugins",
  IsCopilotAllowed = "is-copilot-allowed",
  IsSideloadingAllowed = "is-sideloading-allowed",
  NeedMigrateAadManifest = "need-migrate-aad-manifest",
  CheckCopilotTracingId = "copilot-trace-id",
  CheckSideloadingStatusCode = "status-code",
  CheckSideloadingMethod = "method",
  CheckSideloadingUrl = "url",
  TemplateGroup = "template-group",
  TemplateLanguage = "template-language",
  TemplateScenario = "template-scenario",
  TemplateFallback = "template-fallback",
  TemplateName = "template-name",
  TenantId = "tenant-id",
  TimeCost = "time-cost",
  SampleDownloadDirectory = "sample-download-directory",
  Fallback = "fallback",
  HasSwitchedSubscription = "has-switched-subscription",
  HasSwitchedM365Tenant = "has-switched-m365",
  CustomizeSubscriptionType = "customize-subscription-type",
  IsFromTdp = "is-from-developer-portal",
  ToolkitVersion = "toolkit-version",
  YmlName = "yml-name",
  YmlSchemaVersion = "yml-schema-version",
  GraphPermission = "graph-permission",
  GraphPermissionHasRole = "graph-permission-has-role",
  GraphPermissionHasAdminScope = "graph-permission-has-admin-scope",
  GraphPermissionScopes = "graph-permission-scopes",
  GraphPermissionRoles = "graph-permission-roles",
  RscApplication = "rsc-application",
  RscDelegated = "rsc-delegated",
  WebApplicationId = "web-application-id",

  AadManifest = "aad-manifest",

  CustomCopilotAgent = "custom-copilot-agent",
  CustomCopilotRAG = "custom-copilot-rag",
  LlmService = "llm-service",
  HasAzureOpenAIKey = "has-azure-openai-key",
  HasAzureOpenAIEndpoint = "has-azure-openai-endpoint",
  HasAzureOpenAIDeploymentName = "has-azure-openai-deployment-name",
  HasOpenAIKey = "has-openai-key",

  TDPTraceId = "tdp-trace-id",
  MOSTraceId = "mos-trace-id",
  MOSPATH = "mos-api-path",

  AddAuthType = "add-auth-type",

  MosTitleId = "mos-title-id",
  MosAppId = "mos-app-id",
  IsDeclarativeAgent = "is-declarative-agent",
}

export const TelemetryConstants = {
  eventPrefix: "-start",
};

export enum TelemetryEvent {
  Scaffold = "scaffold",
  GenerateBicep = "generate-arm-templates",
  LocalDebug = "local-debug",
  PostLocalDebug = "post-local-debug",
  Provision = "provision",
  PostProvision = "post-provision",
  PreDeploy = "pre-deploy",
  Deploy = "deploy",
  DownloadSampleStart = "download-sample-start",
  DownloadSample = "download-sample",
  CreateProject = "create",
  AddFeature = "add-feature",
  ProjectUpgrade = "project-upgrade",
  ProjectUpgradeStart = "project-upgrade-start",
  ReadJson = "read-json",
  DecryptUserdata = "decrypt-userdata",
  CheckCopilot = "check-copilot",
  CheckResourceGroupStart = "check-resource-group-start",
  CheckResourceGroup = "check-resource-group",
  CheckSubscriptionStart = "check-subscription-start",
  CheckSubscription = "check-subscription",
  CheckSideloading = "check-sideloading",
  EnvConfig = "env-config",
  DisplayToolingUpdateNotification = "display-tooling-update-notification",
  ProjectMigratorNotificationStart = "project-migrator-notification-start",
  ProjectMigratorNotification = "project-migrator-notification",
  ProjectMigratorMigrateStart = "project-migrator-migrate-start",
  ProjectMigratorMigrate = "project-migrator-migrate",
  ProjectMigratorPrecheckFailed = "project-migrator-pre-check-failed",
  ProjectMigratorError = "project-migrator-error",
  ProjectConsolidateNotificationStart = "project-consolidate-notification-start",
  ProjectConsolidateNotification = "project-consolidate-notification",
  ProjectConsolidateUpgradeStart = "project-consolidate-upgrade-start",
  ProjectConsolidateUpgrade = "project-consolidate-upgrade",
  ProjectConsolidateAddLocalEnvStart = "project-consolidate-add-local-env-start",
  ProjectConsolidateAddLocalEnv = "project-consolidate-add-local-env",
  ProjectConsolidateAddSPFXManifestStart = "project-consolidate-add-spfx-manifest-start",
  ProjectConsolidateAddSPFXManifest = "project-consolidate-add-spfx-manifest",
  ProjectConsolidateCopyAzureManifestStart = "project-consolidate-copy-azure-manifest-start",
  ProjectConsolidateCopyAzureManifest = "project-consolidate-copy-azure-manifest",
  ProjectConsolidateAddAzureManifestStart = "project-consolidate-add-azure-manifest-start",
  ProjectConsolidateAddAzureManifest = "project-consolidate-add-azure-manifest",
  ProjectConsolidateBackupConfigStart = "project-consolidate-backup-config-start",
  ProjectConsolidateBackupConfig = "project-consolidate-backup-config",
  ProjectConsolidateGuideStart = "project-Consolidate-guide-start",
  ProjectConsolidateGuide = "project-consolidate-guide",
  ProjectConsolidateError = "project-consolidate-error",
  ProjectConsolidateCheckManifestError = "project-consolidate-check-manifest-error",
  DetectPortStart = "detect-port-start",
  DetectPort = "detect-port",
  FillProjectId = "fill-project-id",
  ScaffoldFromTemplatesStart = "scaffold-from-templates-start",
  ScaffoldFromTemplates = "scaffold-from-templates",
  GenerateTemplate = "generate-template",
  GenerateSample = "generate-sample",
  ConfirmProvision = "confirm-provision",
  CheckLocalDebugTenant = "check-local-debug-tenant",
  DebugSetUpSSO = "debug-set-up-sso",
  DeploymentInfoNotFound = "deployment-info-not-found",
  InstallScriptNotFound = "install-script-not-found",
  SkipDeploy = "skip-deploy",
  PublishInDeveloperPortal = "publish-in-developer-portal",
  MetaData = "metadata",
  ProjectType = "project-type",
  DependencyApi = "dependency-api",
  AppStudioApi = "app-studio-api",
  MOSApi = "ttk-mos-api",
  ViewPluginManifestAfterAdded = "view-plugin-manifest-after-added",
  AddAuthAction = "add-auth-action",
  ViewAgentManifestAfterAdded = "view-agent-manifest-after-added",
  MosSideloadEnd = "mos-sideload-end",
}

export enum ProjectTypeProps {
  IsTeamsFx = "is-teamsfx",
  TeamsfxConfigType = "teamsfx-config-type",
  TeamsfxConfigVersion = "teamsfx-config-version",
  TeamsfxVersionState = "teamsfx-version-state",
  TeamsfxProjectId = "teamsfx-project-id",
  TeamsManifest = "has-manifest",
  TeamsManifestVersion = "manifest-version",
  TeamsManifestAppId = "manifest-app-id",
  TeamsManifestCapabilities = "manifest-capabilities",
  TeamsJs = "teams-js",
  Lauguages = "languages",
  OfficeAddinProjectType = "office-addin-project-type",
}

export enum TelemetrySuccess {
  Yes = "yes",
  No = "no",
}

export enum TelemetryErrorType {
  UserError = "user",
  SystemError = "system",
}

export enum Component {
  vsc = "extension",
  cli = "cli",
  vs = "vs",
  core = "core",
}

export enum CustomizeResourceGroupType {
  CommandLine = "command-line",
  EnvConfig = "env-config",
  EnvState = "env-state",
  InteractiveCreateDefault = "interactive-create-default",
  InteractiveCreateCustomized = "interactive-create-customized",
  InteractiveUseExisting = "interactive-use-existing",
  FallbackDefault = "fallback-default",
}

export enum CustomizeSubscriptionType {
  CommandLine = "command-line",
  EnvConfig = "env-config",
  EnvState = "env-state",
  Default = "default",
}

export enum ProjectMigratorStatus {
  OK = "ok",
  Cancel = "cancel",
}

export enum WebApplicationIdValue {
  None = "none",
  Default = "default",
  Customized = "customized",
}

export enum ProjectMigratorGuideStatus {
  Reload = "reload",
  LearnMore = "learn-more",
  Cancel = "cancel",
}

export enum ApiSpecTelemetryPropertis {
  SpecNotValidDetails = "spec-not-valid-details",
  InvalidApiSpec = "invalid-api-spec",
}

export function getQuestionValidationErrorEventName(questionName: string) {
  return `invalid-${questionName}`;
}

export function sendTelemetryEvent(
  component: string,
  eventName: string,
  properties?: { [p: string]: string },
  measurements?: { [p: string]: number }
): void {
  if (!properties) {
    properties = {};
  }
  properties[TelemetryProperty.Component] = component;
  TOOLS.telemetryReporter?.sendTelemetryEvent(eventName, properties, measurements);
}

export function sendTelemetryErrorEvent(
  component: string,
  eventName: string,
  fxError: FxError,
  properties?: { [p: string]: string }
): void {
  if (!properties) {
    properties = {};
  }
  properties[TelemetryProperty.Component] = component;

  telemetryUtils.fillInErrorProperties(properties, fxError);

  TOOLS.telemetryReporter?.sendTelemetryErrorEvent(eventName, properties, {});
}

class TelemetryUtils {
  /**
   * fill in telemetry properties for FxError
   * @param error FxError
   * @param props teletry properties
   */
  fillInErrorProperties(props: Record<string, string>, error: FxError): void {
    const errorCode = error.source + "." + error.name;
    const errorType =
      error instanceof SystemError ? TelemetryErrorType.SystemError : TelemetryErrorType.UserError;
    props[TelemetryProperty.Success] = TelemetrySuccess.No;
    props[TelemetryProperty.ErrorCode] = props[TelemetryProperty.ErrorCode] || errorCode;
    props[TelemetryProperty.ErrorType] = errorType;
    props[TelemetryProperty.ErrorMessage] = error.skipProcessInTelemetry
      ? error.message
      : maskSecret(error.message);
    props[TelemetryProperty.ErrorStack] = this.extractMethodNamesFromErrorStack(error.stack); // error stack will not append in error-message any more
    props[TelemetryProperty.ErrorName] = error.name;
    if (error.name === "ScriptExecutionError" || error.name === "ScriptTimeoutError") {
      // don't send error message and error stack for these two types of error for the sake of complience issue
      props[TelemetryProperty.ErrorMessage] = "";
      props[TelemetryProperty.ErrorStack] = "";
    }
    // append global context properties
    props[TelemetryProperty.ErrorComponent] = globalVars.component;
    props[TelemetryProperty.ErrorStage] = globalVars.stage;
    props[TelemetryProperty.ErrorMethod] = globalVars.method;
    props[TelemetryProperty.ErrorSource] = globalVars.source;
    if (error.innerError && error.innerError["code"]) {
      props[TelemetryProperty.ErrorInnerCode] = error.innerError["code"];
    }

    if (error.categories) {
      props[TelemetryProperty.ErrorCat] = error.categories.join("|");
      props[TelemetryProperty.ErrorCat1] = error.categories[0];
      props[TelemetryProperty.ErrorCat2] = error.categories[1];
      props[TelemetryProperty.ErrorCat3] = error.categories[2];
    }

    if (error.telemetryProperties) {
      assign(props, error.telemetryProperties);
    }
  }

  fillinProjectTypeProperties(props: Record<string, string>, projectTypeRes: ProjectTypeResult) {
    const newProps = {
      [ProjectTypeProps.IsTeamsFx]: projectTypeRes.isTeamsFx ? "true" : "false",
      [ProjectTypeProps.TeamsfxConfigType]: projectTypeRes.teamsfxConfigType || "",
      [ProjectTypeProps.TeamsfxConfigVersion]: projectTypeRes.teamsfxConfigVersion || "",
      [ProjectTypeProps.TeamsfxVersionState]: projectTypeRes.teamsfxVersionState || "",
      [ProjectTypeProps.TeamsJs]: projectTypeRes.dependsOnTeamsJs ? "true" : "false",
      [ProjectTypeProps.TeamsManifest]: projectTypeRes.hasTeamsManifest ? "true" : "false",
      [ProjectTypeProps.TeamsManifestVersion]: projectTypeRes.manifestVersion || "",
      [ProjectTypeProps.TeamsManifestAppId]: projectTypeRes.manifestAppId || "",
      [ProjectTypeProps.TeamsfxProjectId]: projectTypeRes.teamsfxProjectId || "",
      [ProjectTypeProps.Lauguages]: projectTypeRes.lauguages.join(","),
      [ProjectTypeProps.TeamsManifestCapabilities]:
        projectTypeRes.manifestCapabilities?.join(",") || "",
      [ProjectTypeProps.OfficeAddinProjectType]: projectTypeRes.officeAddinProjectType || "",
    };
    assign(props, newProps);
  }

  extractMethodNamesFromErrorStack(stack?: string): string {
    if (!stack) return "";
    const methodNamesRegex = /at\s([\w.<>\[\]\s]+)\s\(/g;
    let match;
    const methodNames: string[] = [];
    while ((match = methodNamesRegex.exec(stack)) !== null) {
      methodNames.push(match[1]);
    }
    return methodNames.join(" | ");
  }
}

export const telemetryUtils = new TelemetryUtils();
