// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";
import {
  AuthType,
  ListAPIResult,
  ProjectType,
  SpecParser,
  SpecParserError,
  Utils,
} from "@microsoft/m365-spec-parser";
import {
  ApiOperation,
  AppPackageFolderName,
  BuildFolderName,
  Context,
  CoreCallbackEvent,
  CreateProjectInputs,
  CreateProjectResult,
  CryptoProvider,
  DefaultApiSpecFolderName,
  Func,
  FxError,
  IGenerator,
  IQTreeNode,
  Inputs,
  InputsWithProjectPath,
  ManifestUtil,
  Platform,
  PluginManifestSchema,
  ResponseTemplatesFolderName,
  Result,
  Stage,
  TeamsAppInputs,
  Tools,
  UserError,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import { DotenvParseOutput } from "dotenv";
import fs from "fs-extra";
import * as jsonschema from "jsonschema";
import * as os from "os";
import * as path from "path";
import "reflect-metadata";
import { Container } from "typedi";
import { pathToFileURL } from "url";
import { teamsDevPortalClient } from "../client/teamsDevPortalClient";
import { AppStudioScopes, VSCodeExtensionCommand } from "../common/constants";
import { FeatureFlags, featureFlagManager } from "../common/featureFlags";
import {
  ErrorContextMW,
  TOOLS,
  createContext,
  setErrorContext,
  setTools,
} from "../common/globalVars";
import { getLocalizedString } from "../common/localizeUtils";
import { ListCollaboratorResult, PermissionsResult } from "../common/permissionInterface";
import {
  getProjectMetadata,
  isValidProjectV2,
  isValidProjectV3,
} from "../common/projectSettingsHelper";
import {
  IsDeclarativeAgentManifest,
  ProjectTypeResult,
  projectTypeChecker,
} from "../common/projectTypeChecker";
import { TelemetryEvent, telemetryUtils } from "../common/telemetry";
import { generateDriverContext } from "../common/utils";
import { MetadataV3, VersionSource, VersionState } from "../common/versionMetadata";
import {
  APIKeyAuthType,
  MicrosoftEntraAuthType,
  OAuthAuthType,
} from "../component/configManager/constant";
import { ILifecycle, LifecycleName } from "../component/configManager/interface";
import { YamlParser } from "../component/configManager/parser";
import {
  AadConstants,
  KiotaLastCommands,
  SingleSignOnOptionItem,
  ViewAadAppHelpLinkV5,
} from "../component/constants";
import { coordinator } from "../component/coordinator";
import { UpdateAadAppArgs } from "../component/driver/aad/interface/updateAadAppArgs";
import { UpdateAadAppDriver } from "../component/driver/aad/update";
import { AadManifestHelper } from "../component/driver/aad/utility/aadManifestHelper";
import { buildAadManifest } from "../component/driver/aad/utility/buildAadManifest";
import { AddWebPartDriver } from "../component/driver/add/addWebPart";
import { AddWebPartArgs } from "../component/driver/add/interface/AddWebPartArgs";
import "../component/driver/index";
import { DriverContext } from "../component/driver/interface/commonArgs";
import "../component/driver/script/scriptDriver";
import { updateManifestV3 } from "../component/driver/teamsApp/appStudio";
import { CreateAppPackageDriver } from "../component/driver/teamsApp/createAppPackage";
import { AppStudioError } from "../component/driver/teamsApp/errors";
import { CreateAppPackageArgs } from "../component/driver/teamsApp/interfaces/CreateAppPackageArgs";
import { SyncManifestArgs } from "../component/driver/teamsApp/interfaces/SyncManifest";
import { ValidateAppPackageArgs } from "../component/driver/teamsApp/interfaces/ValidateAppPackageArgs";
import { ValidateManifestArgs } from "../component/driver/teamsApp/interfaces/ValidateManifestArgs";
import { ValidateWithTestCasesArgs } from "../component/driver/teamsApp/interfaces/ValidateWithTestCasesArgs";
import { AppStudioResultFactory } from "../component/driver/teamsApp/results";
import { SyncManifestDriver } from "../component/driver/teamsApp/syncManifest";
import { teamsappMgr } from "../component/driver/teamsApp/teamsappMgr";
import { copilotGptManifestUtils } from "../component/driver/teamsApp/utils/CopilotGptManifestUtils";
import { manifestUtils } from "../component/driver/teamsApp/utils/ManifestUtils";
import { pluginManifestUtils } from "../component/driver/teamsApp/utils/PluginManifestUtils";
import {
  containsUnsupportedFeature,
  getFeaturesFromAppDefinition,
  normalizePath,
} from "../component/driver/teamsApp/utils/utils";
import { ValidateManifestDriver } from "../component/driver/teamsApp/validate";
import { ValidateAppPackageDriver } from "../component/driver/teamsApp/validateAppPackage";
import { ValidateWithTestCasesDriver } from "../component/driver/teamsApp/validateTestCases";
import { createDriverContext } from "../component/driver/util/utils";
import "../component/feature/sso";
import { SSO } from "../component/feature/sso";
import {
  convertSpecParserErrorToFxError,
  generateFromApiSpec,
  generateScaffoldingSummary,
  getParserOptions,
  injectAuthAction,
  listOperations,
  listPluginExistingOperations,
} from "../component/generator/apiSpec/helper";
import { addExistingPlugin } from "../component/generator/copilotExtension/helper";
import { LaunchHelper } from "../component/m365/launchHelper";
import { PackageService } from "../component/m365/packageService";
import { MosServiceEndpoint, MosServiceScope } from "../component/m365/serviceConstant";
import { EnvLoaderMW, EnvWriterMW } from "../component/middleware/envMW";
import { QuestionMW } from "../component/middleware/questionMW";
import {
  expandEnvironmentVariable,
  outputScaffoldingWarningMessage,
} from "../component/utils/common";
import { envUtil } from "../component/utils/envUtil";
import { metadataUtil } from "../component/utils/metadataUtil";
import { pathUtils } from "../component/utils/pathUtils";
import { settingsUtil } from "../component/utils/settingsUtil";
import {
  FileNotFoundError,
  InputValidationError,
  InvalidProjectError,
  MissingRequiredInputError,
  MultipleServerError,
  UnhandledError,
  UserCancelError,
  assembleError,
} from "../error/common";
import { NoNeedUpgradeError } from "../error/upgrade";
import { YamlFieldMissingError } from "../error/yml";
import { SyncManifestInputs, UninstallInputs } from "../question";
import {
  ApiPluginStartOptions,
  AppNamePattern,
  HubTypes,
  ProjectTypeOptions,
  QuestionNames,
  SPFxVersionOptionIds,
  ScratchOptions,
  TeamsAppValidationOptions,
  apiPluginApiSpecOptionId,
} from "../question/constants";
import { createProjectCliHelpNode } from "../question/create";
import { ValidateTeamsAppInputs } from "../question/inputs/ValidateTeamsAppInputs";
import { isAadMainifestContainsPlaceholder } from "../question/other";
import { CallbackRegistry, CoreCallbackFunc } from "./callback";
import { checkPermission, grantPermission, listCollaborator } from "./collaborator";
import { LocalCrypto } from "./crypto";
import { environmentNameManager } from "./environmentName";
import { ConcurrentLockerMW } from "./middleware/concurrentLocker";
import { ContextInjectorMW } from "./middleware/contextInjector";
import { ErrorHandlerMW } from "./middleware/errorHandler";
import { ProjectMigratorMWV3, checkActiveResourcePlugins } from "./middleware/projectMigratorV3";
import {
  getProjectVersionFromPath,
  getTrackingIdFromPath,
  getVersionState,
} from "./middleware/utils/v3MigrationUtils";
import { CoreTelemetryEvent, CoreTelemetryProperty } from "./telemetry";
import { CoreHookContext, PreProvisionResForVS, VersionCheckRes } from "./types";

export class FxCore {
  constructor(tools: Tools) {
    setTools(tools);
  }

  /**
   * @todo this's a really primitive implement. Maybe could use Subscription Model to
   * refactor later.
   */
  public on(event: CoreCallbackEvent, callback: CoreCallbackFunc): void {
    return CallbackRegistry.set(event, callback);
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "createProject", reset: true }),
    ErrorHandlerMW,
    QuestionMW("createProject"),
  ])
  async createProject(inputs: Inputs): Promise<Result<CreateProjectResult, FxError>> {
    const context = createContext();
    if (inputs[QuestionNames.ProjectType] === ProjectTypeOptions.startWithGithubCopilot().id) {
      return ok({ projectPath: "", shouldInvokeTeamsAgent: true });
    }
    inputs[QuestionNames.Scratch] = ScratchOptions.yes().id;
    TOOLS.logProvider?.debug(`Scaffold inputs after question model: ${JSON.stringify(inputs)}`);
    const res = await coordinator.create(context, inputs);
    if (res.isOk()) {
      inputs.projectPath = res.value.projectPath;
    }
    return res;
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "createProjectFromTdp", reset: true }),
    ErrorHandlerMW,
    QuestionMW("createFromTdp"),
  ])
  async createProjectFromTdp(inputs: Inputs): Promise<Result<CreateProjectResult, FxError>> {
    const context = createContext();
    inputs[QuestionNames.Scratch] = ScratchOptions.yes().id;
    // should never happen as we do same check on Developer Portal.
    if (inputs.teamsAppFromTdp) {
      if (containsUnsupportedFeature(inputs.teamsAppFromTdp)) {
        return err(
          new InputValidationError("manifest.json", "Teams app contains unsupported features")
        );
      } else {
        context.telemetryReporter.sendTelemetryEvent(CoreTelemetryEvent.CreateFromTdpStart, {
          [CoreTelemetryProperty.TdpTeamsAppFeatures]: getFeaturesFromAppDefinition(
            inputs.teamsAppFromTdp
          ).join(","),
          [CoreTelemetryProperty.TdpTeamsAppId]: inputs.teamsAppFromTdp.teamsAppId,
        });
      }
    }
    TOOLS.logProvider?.debug(`Scaffold inputs after question model: ${JSON.stringify(inputs)}`);
    const res = await coordinator.create(context, inputs);
    inputs.projectPath = context.projectPath;
    return res;
  }

  @hooks([
    ErrorContextMW({
      component: "FxCore",
      stage: "createProjectByCustomizedGenerator",
      reset: true,
    }),
    ErrorHandlerMW,
  ])
  async createProjectByCustomizedGenerator(
    inputs: CreateProjectInputs,
    generator: IGenerator
  ): Promise<Result<CreateProjectResult, FxError>> {
    //1. input validation
    let folder = inputs["folder"];
    if (!folder) {
      return err(new MissingRequiredInputError("folder"));
    }
    folder = path.resolve(folder);
    const appName = inputs["app-name"];
    if (undefined === appName) return err(new MissingRequiredInputError(QuestionNames.AppName));
    const validateResult = jsonschema.validate(appName, {
      pattern: AppNamePattern,
    });
    if (validateResult.errors && validateResult.errors.length > 0) {
      return err(new InputValidationError(QuestionNames.AppName, validateResult.errors[0].message));
    }
    const projectPath = path.join(folder, appName);

    //2. run generator
    const context = createContext();
    const genRes = await generator.run(context, inputs, projectPath);
    if (genRes.isErr()) return err(genRes.error);
    //3. ensure unique projectId in teamsapp.yaml (optional)
    const ymlPath = path.join(projectPath, MetadataV3.configFile);
    const result: CreateProjectResult = { projectPath: projectPath };
    if (await fs.pathExists(ymlPath)) {
      const ensureRes = await coordinator.ensureTrackingId(projectPath, inputs.projectId);
      if (ensureRes.isErr()) return err(ensureRes.error);
      result.projectId = ensureRes.value;
    }
    return ok(result);
  }

  /**
   * lifecycle command: create new sample project
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "createSampleProject", reset: true }),
    ErrorHandlerMW,
    QuestionMW("createSampleProject"),
  ])
  async createSampleProject(inputs: Inputs): Promise<Result<CreateProjectResult, FxError>> {
    const context = createContext();
    inputs[QuestionNames.Scratch] = ScratchOptions.no().id;
    const res = await coordinator.create(context, inputs);
    inputs.projectPath = context.projectPath;
    return res;
  }
  /**
   * lifecycle commands: provision
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "provision", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    EnvLoaderMW(false),
    ConcurrentLockerMW,
    ContextInjectorMW,
    EnvWriterMW,
  ])
  async provisionResources(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<undefined, FxError>> {
    inputs.stage = Stage.provision;
    const context = createDriverContext(inputs);
    try {
      const res = await coordinator.provision(context, inputs as InputsWithProjectPath);
      if (res.isOk()) {
        ctx!.envVars = res.value;
        return ok(undefined);
      } else {
        // for partial success scenario, output is set in inputs object
        ctx!.envVars = inputs.envVars;
        return err(res.error);
      }
    } finally {
      //reset subscription
      try {
        await TOOLS.tokenProvider.azureAccountProvider.setSubscription("");
      } catch (e) {}
    }
  }

  /**
   * none lifecycle command, uninstall provisioned resources
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "uninstall", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    QuestionMW("uninstall"),
  ])
  async uninstall(inputs: UninstallInputs): Promise<Result<undefined, FxError>> {
    switch (inputs[QuestionNames.UninstallMode as string]) {
      case QuestionNames.UninstallModeManifestId:
        return await this.uninstallByManifestId(inputs);
      case QuestionNames.UninstallModeEnv:
        return await this.uninstallByEnv(inputs);
      case QuestionNames.UninstallModeTitleId:
        return await this.uninstallByTitleId(inputs);
      default:
        return err(new UnhandledError(new Error("Uninstall mode not supported"), "FxCore"));
    }
  }

  /**
   * uninstall provisioned resources by manifest ID
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "uninstallByManifestId", reset: true }),
    ErrorHandlerMW,
  ])
  async uninstallByManifestId(inputs: UninstallInputs): Promise<Result<undefined, FxError>> {
    const manifestId = inputs[QuestionNames.ManifestId as string] as string;
    if (!manifestId) {
      return err(new MissingRequiredInputError("manifest-id", "FxCore"));
    }
    const uninstallOptions = inputs[QuestionNames.UninstallOptions as string];
    const m356AppOption = uninstallOptions?.includes(QuestionNames.UninstallOptionM365);
    const tdpOption = uninstallOptions?.includes(QuestionNames.UninstallOptionTDP);
    const botOption = uninstallOptions?.includes(QuestionNames.UninstallOptionBot);

    if (m356AppOption) {
      const res = await this.uninstallM365App(undefined, manifestId);
      if (res.isErr()) {
        return err(res.error);
      }
    }
    if (botOption) {
      const res = await this.uninstallBotFrameworRegistration(undefined, manifestId);
      if (res.isErr()) {
        return err(res.error);
      }
    }
    // App registraion should be the last to remove, because we might need to query some metadata from TDP.
    if (tdpOption) {
      const res = await this.uninstallAppRegistration(manifestId);
      if (res.isErr()) {
        return err(res.error);
      }
    }

    return ok(undefined);
  }

  /**
   * uninstall provisioned resources by a given environment
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "uninstallByEnv", reset: true }),
    ErrorHandlerMW,
    EnvLoaderMW(true, true),
    ConcurrentLockerMW,
    ContextInjectorMW,
    EnvWriterMW,
  ])
  async uninstallByEnv(
    inputs: UninstallInputs,
    ctx?: CoreHookContext
  ): Promise<Result<undefined, FxError>> {
    if (!inputs.env) {
      return err(new MissingRequiredInputError("env", "FxCore"));
    }
    const teamsappYamlPath = pathUtils.getYmlFilePath(inputs.projectPath!, inputs.env);
    const yamlProjectModel = await metadataUtil.parse(teamsappYamlPath, inputs.env);
    if (yamlProjectModel.isErr()) {
      return err(yamlProjectModel.error);
    }
    const projectModel = yamlProjectModel.value;

    let teamsAppId;
    let botId;
    let m365TitleId;
    let teamsAppIdKeyName = "";
    let botIdKeyName = "";
    let m365TitleIdKeyName = "";
    for (const action of projectModel.provision?.driverDefs ?? []) {
      if (action.uses === "teamsApp/create") {
        teamsAppIdKeyName = action.writeToEnvironmentFile?.teamsAppId || "TEAMS_APP_ID";
        teamsAppId = process.env[teamsAppIdKeyName];
      } else if (action.uses === "botFramework/create") {
        botIdKeyName = action.writeToEnvironmentFile?.botId || "BOT_ID";
        botId = process.env[botIdKeyName];
      } else if (action.uses === "teamsApp/extendToM365") {
        m365TitleIdKeyName = action.writeToEnvironmentFile?.titleId || "M365_TITLE_ID";
        m365TitleId = process.env[m365TitleIdKeyName];
      }
    }

    const uninstallOptions = inputs[QuestionNames.UninstallOptions as string];
    const m356AppOption = uninstallOptions?.includes(QuestionNames.UninstallOptionM365);
    const tdpOption = uninstallOptions?.includes(QuestionNames.UninstallOptionTDP);
    const botOption = uninstallOptions?.includes(QuestionNames.UninstallOptionBot);

    if ((teamsAppId || m365TitleId) && m356AppOption) {
      const res = await this.uninstallM365App(m365TitleId, teamsAppId);
      if (res.isErr()) {
        return err(res.error);
      }
      this.resetEnvVar(teamsAppIdKeyName, ctx);
      this.resetEnvVar(m365TitleIdKeyName, ctx);
    }
    if (botId && botOption) {
      const res = await this.uninstallBotFrameworRegistration(botId);
      if (res.isErr()) {
        return err(res.error);
      }
      this.resetEnvVar(botIdKeyName, ctx);
    }
    // App registraion should be the last to remove, because we might need to query some metadata from TDP.
    if (teamsAppId && tdpOption) {
      const res = await this.uninstallAppRegistration(teamsAppId);
      if (res.isErr()) {
        return err(res.error);
      }
      this.resetEnvVar(teamsAppIdKeyName, ctx);
    }
    return ok(undefined);
  }
  resetEnvVar(key: string, ctx?: CoreHookContext, skipIfNotExist = true, resetValue = ""): void {
    if (!ctx) {
      return;
    }
    if (!ctx.envVars) {
      ctx.envVars = {};
    }
    if (skipIfNotExist && !ctx.envVars[key]) {
      return;
    }
    ctx.envVars[key] = resetValue;
    return;
  }
  /**
   * uninstall provisioned resources by title ID. Titlle mode only uninstalls M365 app.
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "uninstallByTitleId", reset: true }),
    ErrorHandlerMW,
  ])
  async uninstallByTitleId(inputs: UninstallInputs): Promise<Result<undefined, FxError>> {
    const titleId = inputs[QuestionNames.TitleId as string] as string;
    if (!titleId) {
      return err(new MissingRequiredInputError("title-id", "FxCore"));
    }
    const res = await this.uninstallM365App(titleId);
    if (res.isErr()) {
      return err(res.error);
    }
    return ok(undefined);
  }

  /**
   * uninstall sideloaded appps in M365
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "uninstallM365App", reset: true }),
    ErrorHandlerMW,
  ])
  async uninstallM365App(
    titleId?: string,
    manifestId?: string
  ): Promise<Result<undefined, FxError>> {
    if (titleId === undefined && manifestId === undefined) {
      return err(new MissingRequiredInputError("title id or manifest id", "FxCore"));
    }
    const sideloadingServiceEndpoint =
      process.env.SIDELOADING_SERVICE_ENDPOINT ?? MosServiceEndpoint;
    const sideloadingServiceScope = process.env.SIDELOADING_SERVICE_SCOPE ?? MosServiceScope;
    const sideloadingTokenRes = await TOOLS.tokenProvider.m365TokenProvider.getAccessToken({
      scopes: [sideloadingServiceScope],
    });
    if (sideloadingTokenRes.isErr()) {
      return err(sideloadingTokenRes.error);
    }
    const packageService = new PackageService(sideloadingServiceEndpoint, TOOLS.logProvider);
    if (titleId === undefined) {
      try {
        titleId = await packageService.retrieveTitleId(sideloadingTokenRes.value, manifestId ?? "");
      } catch (err: any) {
        await TOOLS.ui.showMessage(
          "info",
          getLocalizedString("core.uninstall.failed.titleId"),
          false
        );
        throw assembleError(err);
      }
    }
    const confirmRes = await TOOLS.ui.confirm?.({
      name: "uninstallM365App",
      title: getLocalizedString("core.uninstall.confirm.m365App", titleId),
      default: true,
    });
    if (confirmRes?.isOk() && confirmRes.value.result === true) {
      await packageService.unacquire(sideloadingTokenRes.value, titleId);
      await TOOLS.ui.showMessage(
        "info",
        getLocalizedString("core.uninstall.success.m365App", titleId),
        false
      );
      await TOOLS.ui.showMessage(
        "info",
        getLocalizedString("core.uninstall.success.delayWarning"),
        false
      );
    } else {
      await TOOLS.ui.showMessage(
        "info",
        getLocalizedString("core.uninstall.confirm.cancel.m365App"),
        false
      );
      return err(new UserCancelError("Uninstall M365 App"));
    }
    return ok(undefined);
  }

  /**
   * uninstall sideloaded apps in Teams Developer Portal
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "uninstallAppRegistration", reset: true }),
    ErrorHandlerMW,
  ])
  async uninstallAppRegistration(manifestId: string): Promise<Result<undefined, FxError>> {
    const appStudioTokenRes = await TOOLS.tokenProvider.m365TokenProvider.getAccessToken({
      scopes: AppStudioScopes,
    });
    if (appStudioTokenRes.isErr()) {
      return err(appStudioTokenRes.error);
    }
    const confirmRes = await TOOLS.ui.confirm?.({
      name: "uninstallAppRegistration",
      title: getLocalizedString("core.uninstall.confirm.tdp", manifestId),
      default: true,
    });
    if (confirmRes?.isOk() && confirmRes.value.result === true) {
      const token = appStudioTokenRes.value;
      await teamsDevPortalClient.deleteApp(token, manifestId);
      await TOOLS.ui.showMessage(
        "info",
        getLocalizedString("core.uninstall.success.tdp", manifestId),
        false
      );
      return ok(undefined);
    } else {
      await TOOLS.ui.showMessage(
        "info",
        getLocalizedString("core.uninstall.confirm.cancel.tdp"),
        false
      );
      return err(new UserCancelError("Uninstall App Registration"));
    }
  }

  /**
   * uninstall bots created in dev.botframework.com
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "uninstallBotFrameworRegistration", reset: true }),
    ErrorHandlerMW,
  ])
  async uninstallBotFrameworRegistration(
    botId?: string,
    manifestId?: string
  ): Promise<Result<undefined, FxError>> {
    if (!botId && !manifestId) {
      return err(new MissingRequiredInputError("bot id or manifest id", "FxCore"));
    }
    const appStudioTokenRes = await TOOLS.tokenProvider.m365TokenProvider.getAccessToken({
      scopes: AppStudioScopes,
    });
    if (appStudioTokenRes.isErr()) {
      return err(appStudioTokenRes.error);
    }
    const token = appStudioTokenRes.value;
    if (!botId) {
      const botIdRes = await teamsDevPortalClient.getBotId(token, manifestId!);
      if (!botIdRes) {
        const msg = getLocalizedString("core.uninstall.botNotFound", manifestId!);
        return err(new UserError("FxCore", "Uninstall", msg, msg));
      }
      botId = botIdRes;
    }
    const confirmRes = await TOOLS.ui.confirm?.({
      name: "uninstallBotFrameworRegistration",
      title: getLocalizedString("core.uninstall.confirm.bot", botId),
      default: true,
    });
    if (confirmRes?.isOk() && confirmRes.value.result === true) {
      await teamsDevPortalClient.deleteBot(token, botId);
      await TOOLS.ui.showMessage(
        "info",
        getLocalizedString("core.uninstall.success.bot", botId),
        false
      );
    } else {
      await TOOLS.ui.showMessage(
        "info",
        getLocalizedString("core.uninstall.confirm.cancel.bot"),
        false
      );
      return err(new UserCancelError("Uninstall Bot Framework Registration"));
    }
    return ok(undefined);
  }

  /**
   * lifecycle commands: deploy
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "deploy", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    EnvLoaderMW(false),
    ConcurrentLockerMW,
    ContextInjectorMW,
    EnvWriterMW,
  ])
  async deployArtifacts(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<undefined, FxError>> {
    inputs.stage = Stage.deploy;
    const context = createDriverContext(inputs);
    const res = await coordinator.deploy(context, inputs as InputsWithProjectPath);
    if (res.isOk()) {
      ctx!.envVars = res.value;
      return ok(undefined);
    } else {
      // for partial success scenario, output is set in inputs object
      ctx!.envVars = inputs.envVars;
      return err(res.error);
    }
  }
  @hooks([ErrorContextMW({ component: "FxCore", stage: "localDebug", reset: true })])
  async localDebug(inputs: Inputs): Promise<Result<undefined, FxError>> {
    inputs.env = environmentNameManager.getLocalEnvName();
    return this.provisionResources(inputs);
  }

  /**
   * none lifecycle command, v3 only
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "deployAadManifest", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    QuestionMW("deployAadManifest"),
    EnvLoaderMW(true, true),
    ConcurrentLockerMW,
    ContextInjectorMW,
  ])
  async deployAadManifest(inputs: Inputs): Promise<Result<undefined, FxError>> {
    inputs.stage = Stage.deployAad;
    const updateAadClient = Container.get<UpdateAadAppDriver>("aadApp/update");
    // In V3, the aad.template.json exist at .fx folder, and output to root build folder.
    const manifestTemplatePath: string = inputs[QuestionNames.AadAppManifestFilePath];
    if (!(await fs.pathExists(manifestTemplatePath))) {
      return err(new FileNotFoundError("deployAadManifest", manifestTemplatePath));
    }
    let manifestOutputPath: string = manifestTemplatePath;
    if (inputs.env && (await isAadMainifestContainsPlaceholder(inputs))) {
      await fs.ensureDir(path.join(inputs.projectPath!, "build"));
      manifestOutputPath = path.join(
        inputs.projectPath!,
        "build",
        `aad.manifest.${inputs.env as string}.json`
      );
    }
    const inputArgs: UpdateAadAppArgs = {
      manifestPath: manifestTemplatePath,
      outputFilePath: manifestOutputPath,
    };
    const Context: DriverContext = createDriverContext(inputs);
    setErrorContext({ component: "aadAppUpdate" });
    const res = await updateAadClient.execute(inputArgs, Context);
    if (res.result.isErr()) {
      return err(res.result.error);
    }
    if (Context.platform === Platform.CLI) {
      const msg = getLocalizedString("core.deploy.aadManifestOnCLISuccessNotice");
      void Context.ui!.showMessage("info", msg, false);
    } else {
      const msg = getLocalizedString("core.deploy.aadManifestSuccessNotice");
      void Context.ui!.showMessage(
        "info",
        msg,
        false,
        getLocalizedString("core.deploy.aadManifestLearnMore")
      ).then((result) => {
        const userSelected = result.isOk() ? result.value : undefined;
        if (userSelected === getLocalizedString("core.deploy.aadManifestLearnMore")) {
          void Context.ui!.openUrl(ViewAadAppHelpLinkV5);
        }
      });
    }
    return ok(undefined);
  }
  /**
   * none lifecycle command, v3 only
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "addWebpart", reset: true }),
    ErrorHandlerMW,
    QuestionMW("addWebpart"),
    ProjectMigratorMWV3,
    ConcurrentLockerMW,
  ])
  async addWebpart(inputs: Inputs): Promise<Result<undefined, FxError>> {
    setErrorContext({ component: "spfxAdd", method: "run" });
    const driver: AddWebPartDriver = Container.get<AddWebPartDriver>("spfx/add");
    const args: AddWebPartArgs = {
      manifestPath: inputs[QuestionNames.ManifestPath],
      localManifestPath: inputs[QuestionNames.LocalTeamsAppManifestFilePath],
      spfxFolder: inputs[QuestionNames.SPFxFolder],
      webpartName: inputs[QuestionNames.SPFxWebpartName],
      framework: inputs[QuestionNames.SPFxFramework],
      spfxPackage: SPFxVersionOptionIds.installLocally,
    };
    const Context: DriverContext = createDriverContext(inputs);
    const res = await driver.run(args, Context);
    if (res.isErr()) {
      return err(res.error);
    }
    return ok(undefined);
  }
  /**
   * lifecycle command: publish
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "publish", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    EnvLoaderMW(false),
    ConcurrentLockerMW,
    ContextInjectorMW,
    EnvWriterMW,
  ])
  async publishApplication(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<undefined, FxError>> {
    inputs.stage = Stage.publish;
    const context = createDriverContext(inputs);
    const res = await coordinator.publish(context, inputs as InputsWithProjectPath);
    if (res.isOk()) {
      ctx!.envVars = res.value;
      return ok(undefined);
    } else {
      // for partial success scenario, output is set in inputs object
      ctx!.envVars = inputs.envVars;
      return err(res.error);
    }
  }
  /**
   * most commands will be deprecated in V3
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "executeUserTask", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    EnvLoaderMW(false),
    ConcurrentLockerMW,
  ])
  async executeUserTask(func: Func, inputs: Inputs): Promise<Result<any, FxError>> {
    let res: Result<any, FxError> = ok(undefined);
    const context = createDriverContext(inputs);
    if (func.method === "addSso") {
      // used in v3 only in VS
      inputs.stage = Stage.addFeature;
      inputs[QuestionNames.Features] = SingleSignOnOptionItem.id;
      const component = Container.get<SSO>("sso");
      setErrorContext({ component: "sso", method: "add" });
      res = await component.add(context as unknown as Context, inputs as InputsWithProjectPath);
    }
    return res;
  }
  /**
   * v3 only none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "buildAadManifest", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    EnvLoaderMW(false),
    ConcurrentLockerMW,
  ])
  async buildAadManifest(inputs: Inputs): Promise<Result<undefined, FxError>> {
    const manifestTemplatePath: string = inputs.AAD_MANIFEST_FILE
      ? inputs.AAD_MANIFEST_FILE
      : path.join(inputs.projectPath!, AadConstants.DefaultTemplateFileName);
    if (!(await fs.pathExists(manifestTemplatePath))) {
      return err(new FileNotFoundError("buildAadManifest", manifestTemplatePath));
    }
    await fs.ensureDir(path.join(inputs.projectPath!, "build"));
    const manifestOutputPath: string = path.join(
      inputs.projectPath!,
      "build",
      // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
      `aad.${inputs.env}.json`
    );
    const Context: DriverContext = createDriverContext(inputs);
    await buildAadManifest(Context, manifestTemplatePath, manifestOutputPath);
    return ok(undefined);
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "convertAadToNewSchema", reset: true }),
    ErrorHandlerMW,
    QuestionMW("convertAadToNewSchema"),
  ])
  async convertAadToNewSchema(inputs: Inputs): Promise<Result<undefined, FxError>> {
    const manifestTemplatePath: string = inputs[QuestionNames.AadAppManifestFilePath];
    const projectPath = inputs[QuestionNames.ProjectPath] as string;
    return AadManifestHelper.convertManifestToNewSchemaAndOverride(
      manifestTemplatePath,
      projectPath
    );
  }

  /**
   * v3 only none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "deployTeamsManifest", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    QuestionMW("selectTeamsAppManifest"),
    EnvLoaderMW(true),
    ConcurrentLockerMW,
    ContextInjectorMW,
    EnvWriterMW,
  ])
  async deployTeamsManifest(
    inputs: Inputs,
    ctx?: CoreHookContext
  ): Promise<Result<undefined, FxError>> {
    inputs.manifestTemplatePath = inputs[QuestionNames.TeamsAppManifestFilePath] as string;
    const context = createContext();
    const res = await updateManifestV3(context, inputs as InputsWithProjectPath);
    if (res.isOk()) {
      ctx!.envVars = envUtil.map2object(res.value);
      return ok(undefined);
    }
    return err(res.error);
  }

  /******
   * CLI v3 commands
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "updateTeamsAppCLIV3", reset: true }),
    ErrorHandlerMW,
  ])
  async updateTeamsAppCLIV3(inputs: TeamsAppInputs): Promise<Result<undefined, FxError>> {
    const res = await teamsappMgr.updateTeamsApp(inputs);
    return res;
  }
  /******
   * CLI v3 commands
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "validateTeamsAppCLIV3", reset: true }),
    ErrorHandlerMW,
  ])
  async validateTeamsAppCLIV3(inputs: TeamsAppInputs): Promise<Result<undefined, FxError>> {
    const res = await teamsappMgr.validateTeamsApp(inputs);
    return res;
  }
  /******
   * CLI v3 commands
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "packageTeamsAppCLIV3", reset: true }),
    ErrorHandlerMW,
  ])
  async packageTeamsAppCLIV3(inputs: TeamsAppInputs): Promise<Result<undefined, FxError>> {
    const res = await teamsappMgr.packageTeamsApp(inputs);
    if (res.isErr()) {
      return err(res.error);
    }
    return ok(undefined);
  }
  /******
   * CLI v3 commands
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "publishTeamsAppCLIV3", reset: true }),
    ErrorHandlerMW,
  ])
  async publishTeamsAppCLIV3(inputs: TeamsAppInputs): Promise<Result<undefined, FxError>> {
    const res = await teamsappMgr.publishTeamsApp(inputs);
    return res;
  }

  /**
   * v3 only none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "validateApplication", reset: true }),
    QuestionMW("validateTeamsApp"),
  ])
  async validateApplication(inputs: ValidateTeamsAppInputs): Promise<Result<any, FxError>> {
    if (inputs["manifest-path"]) {
      return await this.validateManifest(inputs);
    } else if (inputs[QuestionNames.ValidateMethod] === TeamsAppValidationOptions.testCases().id) {
      return await this.validateWithTestCases(inputs);
    } else {
      return await this.validateAppPackage(inputs);
    }
  }
  /**
   * v3 only none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "validateManifest", reset: true }),
    ErrorHandlerMW,
    EnvLoaderMW(true),
    ConcurrentLockerMW,
  ])
  async validateManifest(inputs: ValidateTeamsAppInputs): Promise<Result<any, FxError>> {
    inputs.stage = Stage.validateApplication;
    const context: DriverContext = createDriverContext(inputs);
    const teamsAppManifestFilePath = inputs["manifest-path"] as string;
    const args: ValidateManifestArgs = {
      manifestPath: teamsAppManifestFilePath,
      showMessage: inputs?.showMessage != undefined ? inputs.showMessage : true,
    };
    const driver: ValidateManifestDriver = Container.get("teamsApp/validateManifest");
    const result = await driver.execute(args, context);
    return result.result;
  }
  /**
   * v3 only none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "validateAppPackage", reset: true }),
    ErrorHandlerMW,
    ConcurrentLockerMW,
  ])
  async validateAppPackage(inputs: ValidateTeamsAppInputs): Promise<Result<any, FxError>> {
    inputs.stage = Stage.validateApplication;
    const context: DriverContext = createDriverContext(inputs);
    const teamsAppPackageFilePath = inputs["app-package-file-path"] as string;
    const args: ValidateAppPackageArgs = {
      appPackagePath: teamsAppPackageFilePath,
      showMessage: true,
    };
    const driver: ValidateAppPackageDriver = Container.get("teamsApp/validateAppPackage");
    return (await driver.execute(args, context)).result;
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "validateWithTestCases", reset: true }),
    ErrorHandlerMW,
    ConcurrentLockerMW,
  ])
  async validateWithTestCases(inputs: ValidateTeamsAppInputs): Promise<Result<any, FxError>> {
    const context: DriverContext = createDriverContext(inputs);
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: inputs["app-package-file-path"] as string,
      showMessage: true,
      showProgressBar: true,
    };
    const driver: ValidateWithTestCasesDriver = Container.get("teamsApp/validateWithTestCases");
    return (await driver.execute(args, context)).result;
  }

  /**
   * v3 only none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "syncManifest", reset: true }),
    ErrorHandlerMW,
    QuestionMW("syncManifest"),
    ConcurrentLockerMW,
  ])
  async syncManifest(inputs: SyncManifestInputs): Promise<Result<any, FxError>> {
    const context: DriverContext = createDriverContext(inputs);
    const projectPath = inputs[QuestionNames.ProjectPath] as string;
    const env = inputs[QuestionNames.Env] as string;
    const teamsAppId = inputs[QuestionNames.TeamsAppId];
    const args: SyncManifestArgs = {
      projectPath: projectPath,
      env: env,
      teamsAppId: teamsAppId,
    };
    const driver: SyncManifestDriver = Container.get("teamsApp/syncManifest");
    return (await driver.execute(args, context)).result;
  }

  /**
   * v3 only none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "createAppPackage", reset: true }),
    ErrorHandlerMW,
    QuestionMW("selectTeamsAppManifest"),
    EnvLoaderMW(true),
    ConcurrentLockerMW,
  ])
  async createAppPackage(inputs: Inputs): Promise<Result<any, FxError>> {
    inputs.stage = Stage.createAppPackage;

    const context: DriverContext = createDriverContext(inputs);

    const teamsAppManifestFilePath = inputs?.[QuestionNames.TeamsAppManifestFilePath] as string;

    const driver: CreateAppPackageDriver = Container.get("teamsApp/zipAppPackage");
    const args: CreateAppPackageArgs = {
      manifestPath: teamsAppManifestFilePath,
      outputZipPath:
        inputs[QuestionNames.OutputZipPathParamName] ??
        // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
        `${inputs.projectPath}/${AppPackageFolderName}/${BuildFolderName}/appPackage.${process.env
          .TEAMSFX_ENV!}.zip`,
      outputFolder:
        inputs[QuestionNames.OutputManifestParamName] ??
        // eslint-disable-next-line @typescript-eslint/restrict-template-expressions
        `${inputs.projectPath}/${AppPackageFolderName}/${BuildFolderName}`,
    };
    const result = (await driver.execute(args, context)).result;
    if (context.platform === Platform.VSCode) {
      if (result.isOk()) {
        const isWindows = process.platform === "win32";
        let zipFileName = args.outputZipPath;
        if (!path.isAbsolute(zipFileName)) {
          zipFileName = path.join(context.projectPath, zipFileName);
        }
        let builtSuccess = getLocalizedString(
          "plugins.appstudio.buildSucceedNotice.fallback",
          zipFileName
        );
        if (isWindows) {
          const folderLink = pathToFileURL(path.dirname(zipFileName));
          const appPackageLink = `${
            VSCodeExtensionCommand.openFolder
          }?%5B%22${folderLink.toString()}%22%5D`;
          builtSuccess = getLocalizedString("plugins.appstudio.buildSucceedNotice", appPackageLink);
        }
        context.ui?.showMessage("info", builtSuccess, false);
      }
    }
    return result;
  }
  /**
   * get url to preview the app, may prompt to select env, hub and Teams manifest
   * v3 only none lifecycle command
   * @param {Inputs} inputs
   * @returns the url to preview the app
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "preview", reset: true }),
    ErrorHandlerMW,
    QuestionMW("previewWithTeamsAppManifest"),
    EnvLoaderMW(false),
    ConcurrentLockerMW,
  ])
  async previewWithManifest(inputs: Inputs): Promise<Result<string, FxError>> {
    inputs.stage = Stage.previewWithManifest;

    const hub = inputs[QuestionNames.M365Host] as HubTypes;
    const manifestFilePath = inputs[QuestionNames.TeamsAppManifestFilePath] as string;
    const context = createContext();

    const manifestRes = await manifestUtils.getManifestV3(
      manifestFilePath,
      generateDriverContext(context, inputs as InputsWithProjectPath),
      false
    );
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }

    const teamsAppId = manifestRes.value.id;
    const properties = ManifestUtil.parseCommonProperties(manifestRes.value);

    const launchHelper = new LaunchHelper(TOOLS.tokenProvider.m365TokenProvider, TOOLS.logProvider);
    const result = await launchHelper.getLaunchUrl(hub, teamsAppId, properties, true);
    return result;
  }
  /**
   * Warning: this API only works for CLI_HELP, it has no business with interactive run for CLI!
   */
  getQuestions(stage: Stage, inputs: Inputs): Result<IQTreeNode | undefined, FxError> {
    if (stage === Stage.create) {
      return ok(createProjectCliHelpNode());
    }
    return ok(undefined);
  }

  /**
   * get all dot envs
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "getDotEnvs", reset: true }),
    ErrorHandlerMW,
  ])
  async getDotEnvs(
    inputs: InputsWithProjectPath
  ): Promise<Result<{ [name: string]: DotenvParseOutput }, FxError>> {
    const envListRes = await envUtil.listEnv(inputs.projectPath);
    if (envListRes.isErr()) {
      return err(envListRes.error);
    }
    const res: { [name: string]: DotenvParseOutput } = {};
    for (const env of envListRes.value) {
      const envRes = await envUtil.readEnv(inputs.projectPath, env, false, false);
      if (envRes.isErr()) {
        return err(envRes.error);
      }
      res[env] = envRes.value as DotenvParseOutput;
    }
    return ok(res);
  }
  /**
   * given projectPath and filePath, return whether the filePath is a env file
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "isEnvFile", reset: true })])
  async isEnvFile(projectPath: string, inputFile: string): Promise<Result<boolean, FxError>> {
    const inputFileName = path.basename(inputFile);
    const envName = envUtil.extractEnvNameFromFileName(inputFileName);
    if (!envName) return ok(false);
    const folderRes = await pathUtils.getEnvFolderPath(projectPath);
    if (folderRes.isErr()) return err(folderRes.error);
    const envFolderPath = folderRes.value;
    if (!envFolderPath) return ok(false);
    const inputFileDir = path.dirname(inputFile);
    if (path.resolve(inputFileDir) !== path.resolve(envFolderPath)) return ok(false);
    return ok(true);
  }

  /**
   * get projectId
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "getProjectId", reset: true })])
  async getProjectId(projectPath: string): Promise<Result<string, FxError>> {
    const res = await this.getProjectMetadata(projectPath);
    if (res.isErr()) {
      return err(res.error);
    }
    return ok(res.value.projectId || "");
  }

  /**
   * @description get projectId and version from yml
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "getProjectMetadata", reset: true })])
  async getProjectMetadata(
    projectPath: string
  ): Promise<Result<{ version?: string; projectId?: string }, FxError>> {
    const res = getProjectMetadata(projectPath);
    if (!res) return ok({});
    return Promise.resolve(ok(res));
  }

  /**
   * get Teams App Name from yml
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "getTeamsAppName", reset: true })])
  async getTeamsAppName(projectPath: string): Promise<Result<string, FxError>> {
    const ymlPath = pathUtils.getYmlFilePath(projectPath, "dev");
    const maybeProjectModel = await metadataUtil.parse(ymlPath, "dev");
    if (maybeProjectModel.isErr()) {
      return err(maybeProjectModel.error);
    }
    const projectModel = maybeProjectModel.value as any;
    if (projectModel.provision) {
      const teamsAppCreate = projectModel.provision?.driverDefs.find(
        (d: any) => d.uses === "teamsApp/create"
      );
      if (teamsAppCreate) {
        let name = teamsAppCreate.with.name as string;
        if (name) {
          name = expandEnvironmentVariable(name, { APP_NAME_SUFFIX: "", TEAMSFX_ENV: " " }).trim();
          return ok(name);
        }
      }
    }
    return ok("");
  }

  /**
   * get project info
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "getProjectInfo", reset: true })])
  async getProjectInfo(
    projectPath: string,
    env: string
  ): Promise<
    Result<
      {
        projectId: string;
        teamsAppId: string;
        teamsAppName: string;
        m365TenantId: string;
      },
      FxError
    >
  > {
    const ymlPath = pathUtils.getYmlFilePath(projectPath, env);
    const maybeProjectModel = await metadataUtil.parse(ymlPath, env);
    if (maybeProjectModel.isErr()) {
      return err(maybeProjectModel.error);
    }
    const projectModel = maybeProjectModel.value;
    const readEnvRes = await envUtil.readEnv(projectPath, env, false, true);
    if (readEnvRes.isErr()) {
      return err(readEnvRes.error);
    }
    const envObject = readEnvRes.value;
    const res: {
      projectId: string;
      teamsAppId: string;
      teamsAppName: string;
      m365TenantId: string;
    } = {
      projectId: (projectModel as any).projectId || "",
      teamsAppId: "",
      teamsAppName: "",
      m365TenantId: envObject.TEAMS_APP_TENANT_ID || "",
    };
    if (projectModel.provision) {
      const teamsAppCreate = projectModel.provision.driverDefs.find(
        (d) => d.uses === "teamsApp/create"
      );
      if (teamsAppCreate) {
        const teamsAppIdEnvName = teamsAppCreate.writeToEnvironmentFile?.teamsAppId;
        if (teamsAppIdEnvName) {
          const teamsAppId = envObject[teamsAppIdEnvName];
          res.teamsAppId = teamsAppId;
        }
        const name = (teamsAppCreate.with as any).name;
        if (name) {
          res.teamsAppName = name.replace("-${{TEAMSFX_ENV}}", "") || "";
        }
      }
    }
    return ok(res);
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "grantPermission", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    QuestionMW("grantPermission"),
    EnvLoaderMW(false, true),
    ConcurrentLockerMW,
    EnvWriterMW,
  ])
  async grantPermission(inputs: Inputs): Promise<Result<PermissionsResult, FxError>> {
    inputs.stage = Stage.grantPermission;
    const context = createContext();
    setErrorContext({ component: "collaborator" });
    const res = await grantPermission(
      context,
      inputs as InputsWithProjectPath,
      TOOLS.tokenProvider
    );
    return res;
  }
  /**
   * none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "checkPermission", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    QuestionMW("listCollaborator"),
    EnvLoaderMW(false, true),
    ConcurrentLockerMW,
    EnvWriterMW,
  ])
  async checkPermission(inputs: Inputs): Promise<Result<PermissionsResult, FxError>> {
    inputs.stage = Stage.checkPermission;
    const context = createContext();
    const res = await checkPermission(
      context,
      inputs as InputsWithProjectPath,
      TOOLS.tokenProvider
    );
    return res;
  }
  /**
   * none lifecycle command
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "listCollaborator", reset: true }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    QuestionMW("listCollaborator"),
    EnvLoaderMW(false, true),
    ConcurrentLockerMW,
    EnvWriterMW,
  ])
  async listCollaborator(inputs: Inputs): Promise<Result<ListCollaboratorResult, FxError>> {
    inputs.stage = Stage.listCollaborator;
    const context = createContext();
    const res = await listCollaborator(
      context,
      inputs as InputsWithProjectPath,
      TOOLS.tokenProvider
    );
    return res;
  }
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "getSelectedEnv", reset: true }),
    ErrorHandlerMW,
    EnvLoaderMW(false),
  ])
  getSelectedEnv(inputs: Inputs): Promise<Result<string | undefined, FxError>> {
    return Promise.resolve(ok(inputs.env)); //work for both v2 and v3
  }

  @hooks([ErrorContextMW({ component: "FxCore", stage: "createLocalCrypto" })])
  async createLocalCrypto(projectPath: string): Promise<Result<CryptoProvider, FxError>> {
    const settingsRes = await settingsUtil.readSettings(projectPath);
    if (settingsRes.isErr()) {
      return err(settingsRes.error);
    }
    const projectId = settingsRes.value.trackingId;
    const cryptoProvider = new LocalCrypto(projectId);
    return ok(cryptoProvider);
  }

  /**
   * only for vs code extension
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "encrypt", reset: true }), ErrorHandlerMW])
  async encrypt(plaintext: string, inputs: Inputs): Promise<Result<string, FxError>> {
    const res = await this.createLocalCrypto(inputs.projectPath!);
    if (res.isErr()) {
      return err(res.error);
    }
    return res.value.encrypt(plaintext);
  }
  /**
   * only for vs code extension
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "decrypt", reset: true }), ErrorHandlerMW])
  async decrypt(ciphertext: string, inputs: Inputs): Promise<Result<string, FxError>> {
    const res = await this.createLocalCrypto(inputs.projectPath!);
    if (res.isErr()) {
      return err(res.error);
    }
    return res.value.decrypt(ciphertext);
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "createEnv", reset: true }),
    ErrorHandlerMW,
    QuestionMW("createNewEnv"),
    ConcurrentLockerMW,
  ])
  async createEnv(inputs: Inputs): Promise<Result<undefined, FxError>> {
    return this.createEnvCopyV3(
      inputs[QuestionNames.NewTargetEnvName]!,
      inputs[QuestionNames.SourceEnvName]!,
      inputs.projectPath!
    );
  }
  async createEnvCopyV3(
    targetEnvName: string,
    sourceEnvName: string,
    projectPath: string
  ): Promise<Result<undefined, FxError>> {
    let res = await pathUtils.getEnvFilePath(projectPath, sourceEnvName);
    if (res.isErr()) return err(res.error);
    const sourceDotEnvFile = res.value;

    res = await pathUtils.getEnvFilePath(projectPath, targetEnvName);
    if (res.isErr()) return err(res.error);
    const targetDotEnvFile = res.value;
    if (!sourceDotEnvFile || !targetDotEnvFile)
      return err(new YamlFieldMissingError("environmentFolderPath"));
    if (!(await fs.pathExists(sourceDotEnvFile)))
      return err(new FileNotFoundError("createEnvCopyV3", sourceDotEnvFile));
    const source = await fs.readFile(sourceDotEnvFile);
    const writeStream = fs.createWriteStream(targetDotEnvFile);
    source
      .toString()
      .split(/\r?\n/)
      .forEach((line) => {
        const reg = /^([a-zA-Z_][a-zA-Z0-9_]*=)/g;
        const match = reg.exec(line);
        if (match) {
          if (match[1].startsWith("TEAMSFX_ENV=")) {
            writeStream.write(`TEAMSFX_ENV=${targetEnvName}${os.EOL}`);
          } else if (match[1].startsWith("APP_NAME_SUFFIX=")) {
            writeStream.write(`APP_NAME_SUFFIX=${targetEnvName}${os.EOL}`);
          } else {
            writeStream.write(`${match[1]}${os.EOL}`);
          }
        } else {
          writeStream.write(`${line.trim()}${os.EOL}`);
        }
      });

    writeStream.end();
    TOOLS.logProvider.info(`env file created: ${targetDotEnvFile}`);
    return ok(undefined);
  }

  // a phantom migration method for V3
  @hooks([ErrorContextMW({ component: "FxCore", stage: "phantomMigrationV3", reset: true })])
  async phantomMigrationV3(inputs: Inputs): Promise<Result<undefined, FxError>> {
    // If the project is invalid or upgraded, the ProjectMigratorMWV3 will not take action.
    // Check invaliad/upgraded project here before call ProjectMigratorMWV3
    const projectPath = (inputs.projectPath as string) || "";
    const version = await getProjectVersionFromPath(projectPath);

    if (version.source === VersionSource.teamsapp) {
      return err(new NoNeedUpgradeError());
    } else if (version.source === VersionSource.projectSettings) {
      const isValid = await checkActiveResourcePlugins(projectPath);
      if (!isValid) {
        return err(new InvalidProjectError(projectPath));
      }
    }
    if (version.source === VersionSource.unknown) {
      return err(new InvalidProjectError(projectPath));
    }
    return this.innerMigrationV3(inputs);
  }

  @hooks([ErrorHandlerMW, ProjectMigratorMWV3])
  innerMigrationV3(inputs: Inputs): Result<undefined, FxError> {
    return ok(undefined);
  }

  // a project version check
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "projectVersionCheck", reset: true }),
    ErrorHandlerMW,
  ])
  async projectVersionCheck(inputs: Inputs): Promise<Result<VersionCheckRes, FxError>> {
    const projectPath = (inputs.projectPath as string) || "";
    if (isValidProjectV3(projectPath) || isValidProjectV2(projectPath)) {
      const versionInfo = await getProjectVersionFromPath(projectPath);
      if (!versionInfo.version) {
        return err(new InvalidProjectError(projectPath));
      }
      const trackingId = await getTrackingIdFromPath(projectPath);
      const isSupport = getVersionState(versionInfo);
      // if the project is upgradeable, check whether the project is valid and invalid project should not show upgrade option.
      if (isSupport === VersionState.upgradeable) {
        if (!(await checkActiveResourcePlugins(projectPath))) {
          return err(new InvalidProjectError(projectPath));
        }
      }
      return ok({
        currentVersion: versionInfo.version,
        trackingId,
        isSupport,
        versionSource: VersionSource[versionInfo.source],
      });
    } else {
      return err(new InvalidProjectError(projectPath));
    }
  }

  // apply the given yaml template to current project.
  async apply(
    inputs: Inputs,
    templatePath: string,
    lifecycleName: string
  ): Promise<Result<undefined, FxError>> {
    if (!inputs.projectPath) {
      return err(new InputValidationError("projectPath", "empty", "Core"));
    }
    const projectPath = inputs.projectPath;
    if (!inputs.env) {
      return err(new InputValidationError("env", "empty", "Core"));
    }
    const env = inputs.env;
    const lifecycleName_: LifecycleName = lifecycleName as LifecycleName;
    const result = await envUtil.readEnv(projectPath, env);
    if (result.isErr()) {
      return err(result.error);
    }

    const parser = new YamlParser();
    const maybeProjectModel = await parser.parse(templatePath);
    if (maybeProjectModel.isErr()) {
      return err(maybeProjectModel.error);
    }

    const projectModel = maybeProjectModel.value;
    const driverContext: DriverContext = {
      azureAccountProvider: TOOLS.tokenProvider.azureAccountProvider,
      m365TokenProvider: TOOLS.tokenProvider.m365TokenProvider,
      ui: TOOLS.ui,
      progressBar: undefined,
      logProvider: TOOLS.logProvider,
      telemetryReporter: TOOLS.telemetryReporter!,
      projectPath: projectPath,
      platform: inputs.platform,
    };
    const lifecycle = projectModel[lifecycleName_];
    if (lifecycle) {
      return this.runLifecycle(lifecycle, driverContext, env);
    } else {
      driverContext.logProvider.warning(`No definition found for ${lifecycleName}`);
      return ok(undefined);
    }
  }

  async runLifecycle(
    lifecycle: ILifecycle,
    driverContext: DriverContext,
    env: string
  ): Promise<Result<undefined, FxError>> {
    const r = await lifecycle.execute(driverContext);
    const runResult = r.result;
    if (runResult.isOk()) {
      driverContext.logProvider.info(`Lifecycle ${lifecycle.name} succeeded`);
      const writeResult = await envUtil.writeEnv(
        driverContext.projectPath,
        env,
        envUtil.map2object(runResult.value)
      );
      return writeResult.map(() => undefined);
    } else {
      const error = runResult.error;
      if (error.kind === "Failure") {
        driverContext.logProvider.error(
          `Failed to run ${lifecycle.name} due to ${error.error.name}: ${error.error.message}`
        );
        return err(error.error);
      } else {
        try {
          const failedDriver = error.reason.failedDriver;
          if (error.reason.kind === "UnresolvedPlaceholders") {
            const unresolved = error.reason.unresolvedPlaceHolders;
            driverContext.logProvider.warning(
              `Unresolved placeholders: ${unresolved.join(",")} for driver ${failedDriver.uses}`
            );
            return ok(undefined);
          } else {
            driverContext.logProvider.error(
              `Failed to run ${lifecycle.name} due to ${error.reason.error.name}: ${error.reason.error.message}. Failed driver: ${failedDriver.uses}`
            );
            return err(error.reason.error);
          }
        } finally {
          await envUtil.writeEnv(driverContext.projectPath, env, envUtil.map2object(error.env));
        }
      }
    }
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "preProvisionForVS" }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    EnvLoaderMW(false),
    ConcurrentLockerMW,
    ContextInjectorMW,
  ])
  async preProvisionForVS(inputs: Inputs): Promise<Result<PreProvisionResForVS, FxError>> {
    const context = createDriverContext(inputs);
    return coordinator.preProvisionForVS(context, inputs as InputsWithProjectPath);
  }
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "preCheckYmlAndEnvForVS" }),
    ErrorHandlerMW,
    ProjectMigratorMWV3,
    EnvLoaderMW(false),
    ConcurrentLockerMW,
    ContextInjectorMW,
  ])
  async preCheckYmlAndEnvForVS(inputs: Inputs): Promise<Result<undefined, FxError>> {
    const context = createDriverContext(inputs);
    const result = await coordinator.preCheckYmlAndEnvForVS(
      context,
      inputs as InputsWithProjectPath
    );
    return result;
  }
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "publishInDeveloperPortal" }),
    ErrorHandlerMW,
    ConcurrentLockerMW,
    ContextInjectorMW,
  ])
  async publishInDeveloperPortal(inputs: Inputs): Promise<Result<undefined, FxError>> {
    inputs.stage = Stage.publishInDeveloperPortal;
    const context = createContext();
    return await coordinator.publishInDeveloperPortal(context, inputs as InputsWithProjectPath);
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "copilotPluginAddAPI" }),
    ErrorHandlerMW,
    QuestionMW("copilotPluginAddAPI"),
    ConcurrentLockerMW,
  ])
  async copilotPluginAddAPI(inputs: Inputs): Promise<Result<string, FxError>> {
    const newOperations = inputs[QuestionNames.ApiOperation] as string[];
    const url = inputs[QuestionNames.ApiSpecLocation];
    const manifestPath = inputs[QuestionNames.ManifestPath];
    const isPlugin = inputs[QuestionNames.ApiPluginType] === apiPluginApiSpecOptionId;
    const context = createContext();

    // Get API spec file path from manifest
    const manifestRes = await manifestUtils._readAppManifest(manifestPath);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }

    const confirmRes = await context.userInteraction.showMessage(
      "warn",
      getLocalizedString("core.addApi.confirm", AppPackageFolderName),
      true,
      getLocalizedString("core.addApi.continue")
    );

    if (confirmRes.isErr()) {
      return err(confirmRes.error);
    } else if (confirmRes.value !== getLocalizedString("core.addApi.continue")) {
      return err(new UserCancelError());
    }

    // Merge existing operations in manifest.json
    const specParser = new SpecParser(
      url,
      getParserOptions(isPlugin ? ProjectType.Copilot : ProjectType.SME)
    );
    try {
      const listResult = await specParser.list();
      const apiResultList = listResult.APIs.filter((value) => value.isValid);

      let existingOperations: string[];
      let outputApiSpecPath: string;
      if (isPlugin) {
        if (!inputs[QuestionNames.DestinationApiSpecFilePath]) {
          return err(new MissingRequiredInputError(QuestionNames.DestinationApiSpecFilePath));
        }
        outputApiSpecPath = inputs[QuestionNames.DestinationApiSpecFilePath];
        existingOperations = await listPluginExistingOperations(
          manifestRes.value,
          manifestPath,
          inputs[QuestionNames.DestinationApiSpecFilePath]
        );
      } else {
        const existingOperationIds = manifestUtils.getOperationIds(manifestRes.value);
        existingOperations = apiResultList
          .filter((operation) => existingOperationIds.includes(operation.operationId))
          .map((operation) => operation.api);
        const apiSpecificationFile = manifestRes.value.composeExtensions![0].apiSpecificationFile;
        outputApiSpecPath = path.join(path.dirname(manifestPath), apiSpecificationFile!);
      }

      const operations = [...existingOperations, ...newOperations];

      const adaptiveCardFolder = path.join(
        inputs.projectPath!,
        AppPackageFolderName,
        ResponseTemplatesFolderName
      );

      const authNames: Set<string> = new Set();
      const serverUrls: Set<string> = new Set();
      const authNamesDict: Record<string, AuthType> = {};
      for (const api of operations) {
        const operation = apiResultList.find((op) => op.api === api);
        if (
          operation &&
          operation.auth &&
          (Utils.isBearerTokenAuth(operation.auth.authScheme) ||
            Utils.isOAuthWithAuthCodeFlow(operation.auth.authScheme))
        ) {
          authNames.add(operation.auth.name);
          serverUrls.add(operation.server);
          authNamesDict[operation.auth.name] = operation.auth.authScheme;
        }
      }

      if (serverUrls.size > 1) {
        throw new MultipleServerError(serverUrls);
      }

      if (authNames.size >= 1) {
        for (const authName of authNames) {
          await injectAuthAction(
            inputs.projectPath!,
            [...authNames][0],
            authNamesDict[authName],
            outputApiSpecPath,
            false
          );
        }
      }

      let pluginPath: string | undefined;

      if (isPlugin) {
        const pluginPathRes = await manifestUtils.getPluginFilePath(
          manifestRes.value,
          manifestPath
        );
        if (pluginPathRes.isErr()) {
          return err(pluginPathRes.error);
        }
        pluginPath = pluginPathRes.value;
      }
      const generateResult = await generateFromApiSpec(
        specParser,
        manifestPath,
        inputs,
        context,
        "copilotPluginAddAPI",
        isPlugin ? ProjectType.Copilot : ProjectType.SME,
        {
          destinationApiSpecFilePath: outputApiSpecPath,
          responseTemplateFolder: adaptiveCardFolder,
          pluginManifestFilePath: pluginPath,
        }
      );

      if (generateResult.isErr()) {
        return err(generateResult.error);
      }

      if (generateResult.value.warnings && generateResult.value.warnings.length > 0) {
        const warnSummary = await generateScaffoldingSummary(
          generateResult.value.warnings,
          manifestRes.value,
          path.relative(inputs.projectPath!, outputApiSpecPath),
          pluginPath === undefined ? undefined : path.relative(inputs.projectPath!, pluginPath),
          inputs.projectPath!
        );

        if (warnSummary) {
          context.logProvider.info(warnSummary);
        }
      }
    } catch (e) {
      let error: FxError;
      if (e instanceof SpecParserError) {
        error = convertSpecParserErrorToFxError(e);
      } else {
        error = assembleError(e);
      }
      return err(error);
    }

    const message = getLocalizedString(
      "core.copilot.addAPI.success",
      newOperations,
      inputs.projectPath
    );
    if (inputs.platform !== Platform.VS) {
      void context.userInteraction.showMessage("info", message, false);
    }
    return ok(message);
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "copilotPluginListApiSpecs" }),
    ErrorHandlerMW,
  ])
  async listPluginApiSpecs(inputs: Inputs): Promise<Result<string[], FxError>> {
    try {
      const manifestPath = inputs[QuestionNames.ManifestPath];
      const manifestRes = await manifestUtils._readAppManifest(manifestPath);
      if (manifestRes.isErr()) {
        return err(manifestRes.error);
      }
      const res = await pluginManifestUtils.getApiSpecFilePathFromTeamsManifest(
        manifestRes.value,
        manifestPath
      );
      if (res.isOk()) {
        return ok(res.value);
      } else {
        return err(res.error);
      }
    } catch (error) {
      return err(error as FxError);
    }
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: "copilotPluginListOperations" }),
    ErrorHandlerMW,
  ])
  async copilotPluginListOperations(inputs: Inputs): Promise<Result<ApiOperation[], FxError>> {
    const res = await listOperations(
      createContext(),
      inputs.apiSpecUrl,
      inputs,
      inputs.includeExistingAPIs,
      inputs.shouldLogWarning
    );
    if (res.isErr()) {
      const msg = res.error.map((e) => e.content).join("\n");
      return err(new UserError("FxCore", "ListOpenAPISpecOperationsError", msg, msg));
    } else {
      return ok(res.value);
    }
  }

  /**
   * check project type info
   */
  @hooks([ErrorContextMW({ component: "FxCore", stage: "checkProjectType" }), ErrorHandlerMW])
  async checkProjectType(projectPath: string): Promise<Result<ProjectTypeResult, FxError>> {
    const projectTypeRes = await projectTypeChecker.checkProjectType(projectPath);
    const props: Record<string, string> = {};
    telemetryUtils.fillinProjectTypeProperties(props, projectTypeRes);
    TOOLS.telemetryReporter?.sendTelemetryEvent(TelemetryEvent.ProjectType, props);
    return ok(projectTypeRes);
  }

  /**
   * Add plugin
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: Stage.addPlugin }),
    ErrorHandlerMW,
    QuestionMW("addPlugin"),
    ConcurrentLockerMW,
  ])
  async addPlugin(inputs: Inputs): Promise<Result<undefined | any, FxError>> {
    if (!inputs.projectPath) {
      throw new Error("projectPath is undefined"); // should never happen
    }

    // Call Kiota to select the OpenAPI spec file
    if (
      inputs.platform === Platform.VSCode &&
      featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
      inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id &&
      !!!inputs[QuestionNames.ApiPluginManifestPath]
    ) {
      return ok({
        projectPath: inputs.projectPath,
        lastCommand: KiotaLastCommands.addPlugin,
        manifestPath: inputs[QuestionNames.ManifestPath],
      });
    }

    const context = createContext();
    const teamsManifestPath = inputs[QuestionNames.ManifestPath];
    const appPackageFolder = path.dirname(teamsManifestPath);
    const isGenerateFromApiSpec =
      inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id;
    const isKiotaIntegration =
      inputs.platform === Platform.VSCode &&
      featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
      !!inputs[QuestionNames.ApiPluginManifestPath];

    // validate the project is valid for adding plugin
    const manifestRes = await manifestUtils._readAppManifest(teamsManifestPath);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }

    const teamsManifest = manifestRes.value;
    const declarativeGpt = teamsManifest.copilotExtensions
      ? teamsManifest.copilotExtensions.declarativeCopilots?.[0]
      : teamsManifest.copilotAgents?.declarativeAgents?.[0];
    if (!declarativeGpt?.file) {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.TeamsAppRequiredPropertyMissingError.name,
          AppStudioError.TeamsAppRequiredPropertyMissingError.message(
            "declarativeCopilots",
            teamsManifestPath
          )
        )
      );
    }
    const gptManifestFilePathRes = await copilotGptManifestUtils.getManifestPath(teamsManifestPath);
    if (gptManifestFilePathRes.isErr()) {
      return err(gptManifestFilePathRes.error);
    }

    const declarativeCopilotManifestPath = gptManifestFilePathRes.value;

    const declarativeCopilotManifesRes = await copilotGptManifestUtils.readCopilotGptManifestFile(
      declarativeCopilotManifestPath
    );
    if (declarativeCopilotManifesRes.isErr()) {
      return err(declarativeCopilotManifesRes.error);
    }

    const declarativeCopilotManifest = declarativeCopilotManifesRes.value;
    let confirmMessage = getLocalizedString(
      "core.addApi.confirm",
      path.relative(inputs.projectPath, appPackageFolder)
    );

    // Will be used if generating from API spec
    let specParser: SpecParser | undefined = undefined;
    let authNameAndSchemes: { authName: string; authScheme: AuthType }[] = [];

    if (isGenerateFromApiSpec) {
      specParser = new SpecParser(
        inputs[QuestionNames.ApiSpecLocation].trim(),
        getParserOptions(ProjectType.Copilot, true)
      );
      const listResult = await specParser.list();
      if (
        inputs.platform === Platform.VSCode &&
        featureFlagManager.getBooleanValue(FeatureFlags.KiotaIntegration) &&
        inputs[QuestionNames.ApiPluginType] === ApiPluginStartOptions.apiSpec().id &&
        !!inputs[QuestionNames.ApiPluginManifestPath]
      ) {
        inputs[QuestionNames.ApiOperation] = listResult.APIs.filter((value) => value.isValid).map(
          (value) => value.api
        );
      }
      authNameAndSchemes = this.parseAuthNameAndScheme(listResult, inputs);

      if (authNameAndSchemes.length > 0) {
        const doesLocalYamlPathExists = await fs.pathExists(
          path.join(inputs.projectPath, MetadataV3.localConfigFile)
        );
        confirmMessage = doesLocalYamlPathExists
          ? getLocalizedString(
              "core.addApi.confirm.localTeamsYaml",
              path.relative(inputs.projectPath, appPackageFolder),
              MetadataV3.localConfigFile,
              MetadataV3.configFile
            )
          : getLocalizedString(
              "core.addApi.confirm.teamsYaml",
              path.relative(inputs.projectPath, appPackageFolder),
              MetadataV3.configFile
            );
      }
    }

    // confirm

    const confirmRes = await context.userInteraction.showMessage(
      "warn",
      confirmMessage,
      true,
      getLocalizedString("core.addApi.continue")
    );

    if (confirmRes.isErr()) {
      return err(confirmRes.error);
    } else if (confirmRes.value !== getLocalizedString("core.addApi.continue")) {
      return err(new UserCancelError());
    }

    // find the next available action id
    let actionId = "";
    let suffix = 1;
    actionId = `action_${suffix}`;
    const existingActionIds = declarativeCopilotManifest.actions?.map((action) => action.id);
    while (existingActionIds?.includes(actionId)) {
      suffix += 1;
      actionId = `action_${suffix}`;
    }

    let destinationPluginManifestPath: string;
    // generate files
    if (isGenerateFromApiSpec && specParser) {
      destinationPluginManifestPath = isKiotaIntegration
        ? inputs[QuestionNames.ApiPluginManifestPath].trim()
        : await copilotGptManifestUtils.getDefaultNextAvailablePluginManifestPath(
            appPackageFolder,
            undefined
          );
      const destinationApiSpecPath = isKiotaIntegration
        ? inputs[QuestionNames.ApiSpecLocation].trim()
        : await pluginManifestUtils.getDefaultNextAvailableApiSpecPath(
            inputs[QuestionNames.ApiSpecLocation].trim(),
            path.join(appPackageFolder, DefaultApiSpecFolderName),
            undefined
          );

      const generateRes = await generateFromApiSpec(
        specParser,
        teamsManifestPath,
        inputs,
        context,
        Stage.addPlugin,
        ProjectType.Copilot,
        {
          destinationApiSpecFilePath: destinationApiSpecPath,
          pluginManifestFilePath: destinationPluginManifestPath,
        }
      );
      if (generateRes.isErr()) {
        return err(generateRes.error);
      }

      const warnings = generateRes.value.warnings;
      if (warnings && warnings.length > 0) {
        const warnSummary = await generateScaffoldingSummary(
          warnings,
          manifestRes.value,
          path.relative(inputs.projectPath, destinationApiSpecPath),
          path.relative(inputs.projectPath, destinationPluginManifestPath),
          inputs.projectPath
        );
        context.logProvider.info(warnSummary + "\n");
      }

      const addActionRes = await copilotGptManifestUtils.addAction(
        declarativeCopilotManifestPath,
        actionId,
        normalizePath(path.relative(appPackageFolder, destinationPluginManifestPath), true)
      );
      if (addActionRes.isErr()) {
        return err(addActionRes.error);
      }

      for (const authNameAndScheme of authNameAndSchemes) {
        await this.updateAuthActionInYaml(
          authNameAndScheme.authName,
          authNameAndScheme.authScheme,
          inputs.projectPath,
          destinationApiSpecPath,
          destinationPluginManifestPath
        );
      }
    } else {
      const addPluginRes = await addExistingPlugin(
        declarativeCopilotManifestPath,
        inputs[QuestionNames.PluginManifestFilePath].trim(),
        inputs[QuestionNames.PluginOpenApiSpecFilePath].trim(),
        actionId,
        context,
        Stage.addPlugin
      );

      if (addPluginRes.isErr()) {
        return err(addPluginRes.error);
      }
      destinationPluginManifestPath = addPluginRes.value.destinationPluginManifestPath;
      const warningMessage = outputScaffoldingWarningMessage(addPluginRes.value.warnings);
      context.logProvider.info(warningMessage);
    }

    if (inputs.platform === Platform.VSCode) {
      const successMessage = getLocalizedString("core.addPlugin.success.vsc", actionId);
      const viewPluginManifest = getLocalizedString("core.addPlugin.success.viewPluginManifest");
      void context.userInteraction
        .showMessage("info", successMessage, false, viewPluginManifest)
        .then((userRes) => {
          if (userRes.isOk() && userRes.value === viewPluginManifest) {
            context.telemetryReporter.sendTelemetryEvent(
              TelemetryEvent.ViewPluginManifestAfterAdded
            );
            void TOOLS?.ui?.openFile?.(destinationPluginManifestPath);
          }
        });
    } else {
      const successMessage = getLocalizedString(
        "core.addPlugin.success",
        actionId,
        destinationPluginManifestPath
      );
      void context.userInteraction.showMessage("info", successMessage, false);
    }

    return ok(undefined);
  }

  /**
   * Add Knowledge
   */
  // eslint-disable-next-line @typescript-eslint/require-await
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: Stage.addKnowledge }),
    ErrorHandlerMW,
    QuestionMW("addKnowledge"),
    ConcurrentLockerMW,
  ])
  async addKnowledge(inputs: Inputs): Promise<Result<undefined | any, FxError>> {
    return ok(undefined);
  }

  /**
   * Kiota regenerate
   * Need to update da manifest and update teamsapp.yml
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: Stage.addPlugin }),
    ErrorHandlerMW,
    QuestionMW("kiotaRegenerate"),
    ConcurrentLockerMW,
  ])
  async kiotaRegenerate(inputs: Inputs): Promise<Result<undefined, FxError>> {
    if (!inputs.projectPath) {
      throw new Error("projectPath is undefined"); // should never happen
    }

    const teamsManifestPath = inputs[QuestionNames.ManifestPath];
    const appPackageFolder = path.dirname(teamsManifestPath);
    const context = createContext();

    // validate the project is valid for adding plugin
    const manifestRes = await manifestUtils._readAppManifest(teamsManifestPath);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }

    const teamsManifest = manifestRes.value;
    const declarativeGpt = teamsManifest.copilotExtensions
      ? teamsManifest.copilotExtensions.declarativeCopilots?.[0]
      : teamsManifest.copilotAgents?.declarativeAgents?.[0];
    if (!declarativeGpt?.file) {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.TeamsAppRequiredPropertyMissingError.name,
          AppStudioError.TeamsAppRequiredPropertyMissingError.message(
            "declarativeCopilots",
            teamsManifestPath
          )
        )
      );
    }
    const gptManifestFilePathRes = await copilotGptManifestUtils.getManifestPath(teamsManifestPath);
    if (gptManifestFilePathRes.isErr()) {
      return err(gptManifestFilePathRes.error);
    }

    const declarativeCopilotManifestPath = gptManifestFilePathRes.value;

    const declarativeCopilotManifesRes = await copilotGptManifestUtils.readCopilotGptManifestFile(
      declarativeCopilotManifestPath
    );
    if (declarativeCopilotManifesRes.isErr()) {
      return err(declarativeCopilotManifesRes.error);
    }

    const declarativeCopilotManifest = declarativeCopilotManifesRes.value;
    const pluginManifestFilePath = inputs[QuestionNames.ApiPluginManifestPath];

    let actionId = "";
    declarativeCopilotManifest.actions?.forEach((action) => {
      const actionFilePath = path.normalize(path.join(appPackageFolder, action.file));
      if (actionFilePath.toLowerCase() === path.normalize(pluginManifestFilePath).toLowerCase()) {
        actionId = action.id;
      }
    });

    // Does not found same plugin manifest in da manifest, so add a new action
    // Should not happen when regenerate
    if (!actionId) {
      let suffix = 1;
      actionId = `action_${suffix}`;
      const existingActionIds = declarativeCopilotManifest.actions?.map((action) => action.id);
      while (existingActionIds?.includes(actionId)) {
        suffix += 1;
        actionId = `action_${suffix}`;
      }

      const addActionRes = await copilotGptManifestUtils.addAction(
        declarativeCopilotManifestPath,
        actionId,
        normalizePath(path.relative(appPackageFolder, pluginManifestFilePath), true)
      );

      if (addActionRes.isErr()) {
        return err(addActionRes.error);
      }
    }

    // Update teamsapp.local.yaml and teamsapp.yaml if need to add auth action
    const specPath = inputs[QuestionNames.ApiSpecLocation].trim();
    const specParser = new SpecParser(specPath, getParserOptions(ProjectType.Copilot, true));
    const listResult = await specParser.list();
    inputs[QuestionNames.ApiOperation] = listResult.APIs.filter((value) => value.isValid).map(
      (value) => value.api
    );

    const authNameAndSchemes = this.parseAuthNameAndScheme(listResult, inputs);
    for (const authInfo of authNameAndSchemes) {
      await this.updateAuthActionInYaml(
        authInfo.authName,
        authInfo.authScheme,
        inputs.projectPath,
        specPath,
        pluginManifestFilePath,
        false
      );
    }

    // Add dc to plugin manifest
    const generateRes = await generateFromApiSpec(
      specParser,
      teamsManifestPath,
      inputs,
      context,
      Stage.kiotaRegenerate,
      ProjectType.Copilot,
      {
        destinationApiSpecFilePath: inputs[QuestionNames.ApiSpecLocation].trim(),
        pluginManifestFilePath: inputs[QuestionNames.ApiPluginManifestPath].trim(),
      }
    );
    if (generateRes.isErr()) {
      return err(generateRes.error);
    }

    for (const action of declarativeCopilotManifest.actions!) {
      const actionPath = path.normalize(path.join(appPackageFolder, action.file));
      await copilotGptManifestUtils.updateConversationStarters(
        actionPath,
        declarativeCopilotManifest
      );
    }
    await copilotGptManifestUtils.writeCopilotGptManifestFile(
      declarativeCopilotManifest,
      declarativeCopilotManifestPath
    );

    return ok(undefined);
  }

  @hooks([
    ErrorContextMW({ component: "FxCore", stage: Stage.addAuthAction }),
    ErrorHandlerMW,
    QuestionMW("addAuthAction"),
    ConcurrentLockerMW,
  ])
  async addAuthAction(inputs: Inputs): Promise<Result<undefined, FxError>> {
    if (!inputs.projectPath) {
      throw new Error("projectPath is undefined"); // should never happen
    }

    const pluginManifestPath = inputs[QuestionNames.PluginManifestFilePath] as string;
    const apiSpecRelativePath = inputs[QuestionNames.ApiSpecLocation] as string;
    const apiOperation = inputs[QuestionNames.ApiOperation] as string[];
    const authName = inputs[QuestionNames.AuthName] as string;
    const apiSpecPath = path.normalize(
      path.join(path.dirname(pluginManifestPath), apiSpecRelativePath)
    );
    let authType;
    switch (inputs[QuestionNames.ApiAuth] as string) {
      case "api-key":
      default:
        authType = APIKeyAuthType;
        break;
      case "oauth":
        authType = OAuthAuthType;
        break;
      case "microsoft-entra":
        authType = MicrosoftEntraAuthType;
        break;
    }

    const addAuthActionRes = await injectAuthAction(
      inputs.projectPath,
      authName,
      undefined,
      apiSpecPath,
      true,
      authType
    );

    if (addAuthActionRes?.registrationIdEnvName) {
      const pluginManifest = (await fs.readJson(pluginManifestPath)) as PluginManifestSchema;
      pluginManifest.runtimes?.push({
        type: "OpenApi",
        auth: {
          type: authType as "None" | "OAuthPluginVault" | "ApiKeyPluginVault",
          reference_id: `\$\{\{${addAuthActionRes.registrationIdEnvName}\}\}`,
        },
        spec: {
          url: apiSpecRelativePath,
          progress_style: "ShowUsageWithInputAndOutput",
        },
        run_for_functions: apiOperation,
      });
      await fs.writeJson(pluginManifestPath, pluginManifest, { spaces: 4 });
    }

    return ok(undefined);
  }

  private async updateAuthActionInYaml(
    authName: string | undefined,
    authScheme: AuthType | undefined,
    projectPath: string,
    apSpecPath: string,
    pluginManifestPath: string,
    forceToAddNew = true
  ): Promise<void> {
    if (authName && authScheme) {
      const authInjectRes = await injectAuthAction(
        projectPath,
        authName,
        authScheme,
        apSpecPath,
        forceToAddNew
      );
      if (
        authInjectRes?.defaultRegistrationIdEnvName &&
        authInjectRes?.registrationIdEnvName &&
        authInjectRes.defaultRegistrationIdEnvName !== authInjectRes.registrationIdEnvName
      ) {
        const pluginManifestContent = await fs.readFile(pluginManifestPath, "utf-8");
        const updatedPluginManifestContext = pluginManifestContent.replace(
          authInjectRes.defaultRegistrationIdEnvName,
          authInjectRes.registrationIdEnvName
        );
        await fs.writeFile(pluginManifestPath, updatedPluginManifestContext);
      }
    }
  }

  private parseAuthNameAndScheme(
    listResult: ListAPIResult,
    inputs: Inputs
  ): { authName: string; authScheme: AuthType }[] {
    const authApis = listResult.APIs.filter((value) => value.isValid && !!value.auth);
    const result: {
      authName: string;
      authScheme: AuthType;
    }[] = [];
    for (const api of inputs[QuestionNames.ApiOperation] as string[]) {
      const operation = authApis.find((op) => op.api === api);
      if (
        operation &&
        operation.auth &&
        (Utils.isBearerTokenAuth(operation.auth.authScheme) ||
          Utils.isOAuthWithAuthCodeFlow(operation.auth.authScheme))
      ) {
        if (result.find((value) => value.authName === operation.auth!.name)) {
          continue;
        }
        result.push({ authName: operation.auth.name, authScheme: operation.auth.authScheme });
      }
    }
    return result;
  }

  async isDelcarativeAgentApp(inputs: Inputs): Promise<Result<any, FxError>> {
    const projectPath = inputs[QuestionNames.ProjectPath] as string;
    const manifestRes = await manifestUtils.readAppManifest(projectPath);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }
    return ok(IsDeclarativeAgentManifest(manifestRes.value));
  }
}
