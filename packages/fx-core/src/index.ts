// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * File structure of this package:
 * ./common: contains common utilities and constants that are shared across different components.
 * ./component: contains the implementation of different components
 * ./core: contains the FxCore class that is the entry points implementing the lifecycle APIs of the Teams Toolkit.
 * ./error: contains the error classes used in the Teams Toolkit.
 * ./question: contains the question models used in the Teams Toolkit.
 * ./ui: contains the UI related components.
 */

import "reflect-metadata";
export { teamsDevPortalClient } from "./client/teamsDevPortalClient";
export { askSubscription } from "./common/azureUtils";
export {
  AppStudioScopes,
  AuthSvcScopes,
  AzureScopes,
  getAllowedAppMaps,
  GraphReadUserScopes,
  GraphScopes,
  SPFxScopes,
} from "./common/constants";
export { Correlator } from "./common/correlator";
export {
  featureFlagManager,
  FeatureFlagName,
  FeatureFlags,
  isFeatureFlagEnabled,
} from "./common/featureFlags";
export { globalStateGet, globalStateUpdate } from "./common/globalState";
export { AadSet } from "./common/globalVars";
export { getDefaultString, getLocalizedString } from "./common/localizeUtils";
export * from "./common/permissionInterface";
export * from "./common/projectSettingsHelper";
export {
  projectTypeChecker,
  ProjectTypeResult,
  TeamsfxConfigType,
  TeamsfxVersionState,
} from "./common/projectTypeChecker";
export { sendRequestWithRetry, sendRequestWithTimeout } from "./common/requestUtils";
export { SampleConfig, sampleProvider, SampleUrlInfo } from "./common/samples";
export {
  convertToAlphanumericOnly,
  getHashedEnv,
  getResourceGroupNameFromResourceId,
  getUuid,
  isValidHttpUrl,
  loadingDefaultPlaceholder,
  loadingOptionsPlaceholder,
  maskSecret,
  MaskSecretOptions,
  parseFromResourceId,
} from "./common/stringUtils";
export { telemetryUtils } from "./common/telemetry";
export { getSideloadingStatus, getSPFxTenant, listDevTunnels } from "./common/tools";
export { MetadataV3, VersionState } from "./common/versionMetadata";
export { SummaryConstant } from "./component/configManager/constant";
export { KiotaLastCommands } from "./component/constants";
export { CheckerFactory } from "./component/deps-checker/checkerFactory";
export {
  DepsCheckerEvent,
  TelemetryMessurement,
} from "./component/deps-checker/constant/telemetry";
export { CoreDepsLoggerAdapter } from "./component/deps-checker/coreDepsLoggerAdapter";
export { CoreDepsTelemetryAdapter } from "./component/deps-checker/coreDepsTelemetryAdapter";
export * from "./component/deps-checker/depsChecker";
export { DepsLogger, EmptyLogger } from "./component/deps-checker/depsLogger";
export { DepsManager } from "./component/deps-checker/depsManager";
export { DepsTelemetry, EmptyTelemetry } from "./component/deps-checker/depsTelemetry";
export { FuncToolChecker } from "./component/deps-checker/internal/funcToolChecker";
export { LtsNodeChecker } from "./component/deps-checker/internal/nodeChecker";
export { getPermissionMap } from "./component/driver/aad/permissions/index";
export { AppDefinition } from "./component/driver/teamsApp/interfaces/appdefinitions/appDefinition";
export { manifestUtils } from "./component/driver/teamsApp/utils/ManifestUtils";
export { pluginManifestUtils } from "./component/driver/teamsApp/utils/PluginManifestUtils";
export { generateScaffoldingSummary } from "./component/generator/apiSpec/helper";
export { DefaultTemplateGenerator } from "./component/generator/defaultGenerator";
export { HelperMethods } from "./component/generator/officeAddin/helperMethods";
export { TemplateInfo } from "./component/generator/templates/templateInfo";
export { getSampleFileInfo, runWithLimitedConcurrency } from "./component/generator/utils";
export * from "./component/local/constants";
export { LocalCertificateManager } from "./component/local/localCertificateManager";
export { LocalEnvManager } from "./component/local/localEnvManager";
export { LocalTelemetryReporter, TelemetryContext } from "./component/local/localTelemetryReporter";
export { loadTeamsFxDevScript } from "./component/local/packageJsonHelper";
export { Hub } from "./component/m365/constants";
export { PackageService } from "./component/m365/packageService";
export { MosServiceEndpoint, MosServiceScope } from "./component/m365/serviceConstant";
export * from "./component/middleware/actionExecutionMW";
export { outputScaffoldingWarningMessage } from "./component/utils/common";
export { DotenvOutput, envUtil } from "./component/utils/envUtil";
export { metadataUtil } from "./component/utils/metadataUtil";
export { pathUtils } from "./component/utils/pathUtils";
export { newResourceGroupOption, resourceGroupHelper } from "./component/utils/ResourceGroupHelper";
export { CoreCallbackFunc } from "./core/callback";
export { CollaborationConstants } from "./core/collaborator";
export { environmentManager } from "./core/environment";
export { environmentNameManager } from "./core/environmentName";
export { FxCore } from "./core/FxCore";
export { PreProvisionResForVS, VersionCheckRes } from "./core/types";
export { getProjectTypeByCapability } from "./question/scaffold/vsc/createRootNode";
export { isTdpTemplate } from "./question/scaffold/vsc/createFromTdpNode";
export * from "./error/index";
export * from "./question/constants";
export * from "./question/inputs";
export * from "./question/options";
