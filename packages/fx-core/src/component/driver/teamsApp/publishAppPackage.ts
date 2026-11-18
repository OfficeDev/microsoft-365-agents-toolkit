// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  FxError,
  Platform,
  Result,
  SystemError,
  TeamsAppManifest,
  UserError,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import axios from "axios";
import fs from "fs-extra";
import { merge } from "lodash";
import path from "path";
import semver from "semver";
import { Service } from "typedi";
import { GraphClient } from "../../../client/graphClient";
import { SovereignCloudEnvironment } from "../../../common/accountUtils";
import { GraphTeamsAppCatalogReadWriteScopes } from "../../../common/constants";
import { FeatureFlagName } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import { getSPFxToken } from "../../../common/tools";
import { FileNotFoundError, InvalidActionInputError, UserCancelError } from "../../../error/common";
import { getAbsolutePath } from "../../utils/common";
import { getSPOTenant } from "../deploy/spfx/deployDriver";
import {
  isAmbiguousPublicationState,
  loadPreparedPublication,
  PreparedPublicationError,
  PreparedPublicationJournal,
  isPreparedCatalogUrl,
  updatePreparedPublication,
  createPreparedPublicationErrorData,
  PreparedPublicationErrorStage,
  verifyPreparedApplicationPackage,
  verifyPreparedArtifact,
  verifyPreparedSharePointPackage,
  verifyPreparedSources,
  withPreparedPublicationIdentityLock,
} from "../deploy/spfx/utility/preparedPublication";
import { verifyPreparedJournalIdentity } from "../deploy/spfx/utility/tokenIdentity";
import { PreparedSharePointRequestContext, SPOClient } from "../deploy/spfx/utility/spoClient";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { WrapDriverContext } from "../util/wrapUtil";
import { Constants } from "./constants";
import { AppStudioError } from "./errors";
import { PublishAppPackageArgs } from "./interfaces/PublishAppPackageArgs";
import { PublishingState } from "./interfaces/appdefinitions/IPublishingAppDefinition";
import { AppStudioResultFactory } from "./results";
import { verifyLocalMCPPluginCerts } from "./utils/McpCertVerification";
import { TelemetryPropertyKey } from "./utils/telemetry";

export const actionName = "teamsApp/publishAppPackage";

const defaultOutputNames = {
  publishedAppId: "TEAMS_APP_PUBLISHED_APP_ID",
};

@Service(actionName)
export class PublishAppPackageDriver implements StepDriver {
  description = getLocalizedString("driver.teamsApp.description.publishDriver");
  readonly progressTitle = getLocalizedString("driver.teamsApp.progressBar.publishTeamsAppStep2.2");

  public async execute(
    args: PublishAppPackageArgs,
    context: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(context, actionName, actionName);
    const res = await this.publish(args, wrapContext, outputEnvVarNames);
    return {
      result: res,
      summaries: wrapContext.summaries,
    };
  }

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async publish(
    args: PublishAppPackageArgs,
    context: WrapDriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<Result<Map<string, string>, FxError>> {
    if (this.hasPreparedArguments(args)) {
      return this.publishPrepared(args, context, outputEnvVarNames);
    }
    if (
      process.env[FeatureFlagName.SovereignCloudEnvironment] === SovereignCloudEnvironment.GCCH ||
      process.env[FeatureFlagName.SovereignCloudEnvironment] === SovereignCloudEnvironment.DOD
    ) {
      context.logProvider.warning(
        getLocalizedString("driver.teamsApp.warning.publishUnsupportedCloud", actionName)
      );
      return ok(new Map<string, string>());
    }

    const argsValidationResult = this.validateArgs(args);
    if (argsValidationResult.isErr()) {
      return err(argsValidationResult.error);
    }

    if (!outputEnvVarNames) {
      outputEnvVarNames = new Map(Object.entries(defaultOutputNames));
    }

    const appPackagePath = getAbsolutePath(args.appPackagePath, context.projectPath);
    if (!(await fs.pathExists(appPackagePath))) {
      return err(
        new FileNotFoundError(
          actionName,
          appPackagePath,
          "https://aka.ms/teamsfx-actions/teamsapp-publish"
        )
      );
    }
    const archivedFile = await fs.readFile(appPackagePath);

    const zipEntries = new AdmZip(archivedFile).getEntries();

    const manifestFile = zipEntries.find((x) => x.entryName === Constants.MANIFEST_FILE);
    if (!manifestFile) {
      return err(
        new FileNotFoundError(
          actionName,
          Constants.MANIFEST_FILE,
          "https://aka.ms/teamsfx-actions/teamsapp-publish"
        )
      );
    }
    const manifestString = manifestFile.getData().toString();
    const manifest = JSON.parse(manifestString) as TeamsAppManifest;

    const declarativeAgents =
      manifest.copilotExtensions?.declarativeCopilots || manifest.copilotAgents?.declarativeAgents;

    if (declarativeAgents && declarativeAgents.length > 0) {
      const declarativeAgentFile = zipEntries.find(
        (x) => x.entryName === declarativeAgents[0].file
      );

      if (declarativeAgentFile) {
        const declarativeAgentContent = declarativeAgentFile.getData().toString();
        const declarativeAgentManifest = JSON.parse(declarativeAgentContent);

        if (declarativeAgentManifest.actions) {
          for (const action of declarativeAgentManifest.actions) {
            const actionFile = zipEntries.find((x) => x.entryName === action.file);
            if (actionFile) {
              const isValid = await verifyLocalMCPPluginCerts(actionFile);
              if (!isValid) {
                const message = getLocalizedString(
                  "driver.teamsApp.error.localMcpCertVerificationFailed"
                );
                return err(
                  AppStudioResultFactory.UserError(AppStudioError.ValidationFailedError.name, [
                    message,
                    message,
                  ])
                );
              }
            }
          }
        }
      }
    }

    // manifest.id === externalID
    const graphTokenRes = await context.m365TokenProvider.getAccessToken({
      scopes: GraphTeamsAppCatalogReadWriteScopes,
    });
    if (graphTokenRes.isErr()) {
      return err(graphTokenRes.error);
    }
    const graphClient = new GraphClient(context.m365TokenProvider, context.logProvider);

    let result;

    const message = getLocalizedString("driver.teamsApp.progressBar.publishTeamsAppStep1");
    context.addSummary(message);

    try {
      const existApp = await graphClient.getStagedApp(graphTokenRes.value, manifest.id);
      if (existApp) {
        context.addSummary(
          getLocalizedString("driver.teamsApp.summary.publishTeamsAppExists", manifest.id)
        );
        let executePublishUpdate = false;
        let description = getLocalizedString(
          "plugins.appstudio.pubWarn",
          existApp.displayName,
          existApp.publishingState
        );
        if (existApp.lastModifiedDateTime) {
          description =
            description +
            getLocalizedString(
              "plugins.appstudio.lastModified",
              existApp.lastModifiedDateTime?.toLocaleString()
            );
        }
        description =
          description + getLocalizedString("plugins.appstudio.updatePublihsedAppConfirm");
        const confirm = getLocalizedString("core.option.confirm");
        const res = await context.ui?.showMessage("warn", description, true, confirm);
        if (res?.isOk() && res.value === confirm) executePublishUpdate = true;

        if (executePublishUpdate) {
          const message = getLocalizedString("driver.teamsApp.progressBar.publishTeamsAppStep2.1");
          context.addSummary(message);
          context.logProvider.debug(message);
          const appId = await graphClient.publishTeamsAppUpdate(
            graphTokenRes.value,
            manifest.id,
            archivedFile
          );
          result = new Map([[outputEnvVarNames.get("publishedAppId") as string, appId]]);
          merge(context.telemetryProperties, {
            [TelemetryPropertyKey.updateExistingApp]: "true",
            [TelemetryPropertyKey.publishedAppId]: appId,
          });
        } else {
          return err(new UserCancelError(actionName));
        }
      } else {
        context.addSummary(
          getLocalizedString("driver.teamsApp.summary.publishTeamsAppNotExists", manifest.id)
        );
        const message = getLocalizedString("driver.teamsApp.progressBar.publishTeamsAppStep2.2");
        context.addSummary(message);
        context.logProvider.debug(message);
        const appId = await graphClient.publishTeamsApp(
          graphTokenRes.value,
          manifest.id,
          archivedFile
        );
        result = new Map([[outputEnvVarNames.get("publishedAppId") as string, appId]]);
        merge(context.telemetryProperties, {
          [TelemetryPropertyKey.updateExistingApp]: "false",
        });
      }
    } catch (e: any) {
      return err(e);
    }

    context.logProvider.info(`Publish success!`);
    context.addSummary(
      getLocalizedString("driver.teamsApp.summary.publishTeamsAppSuccess", manifest.id)
    );
    if (context.platform === Platform.CLI) {
      const msg = getLocalizedString(
        "plugins.appstudio.publishSucceedNotice.cli",
        manifest.name.short,
        Constants.TEAMS_ADMIN_PORTAL,
        Constants.TEAMS_MANAGE_APP_DOC
      );
      context.ui?.showMessage("info", msg, false);
    }
    return ok(result);
  }

  private hasPreparedArguments(args: PublishAppPackageArgs): boolean {
    return (
      args?.preparedPublicationPath !== undefined ||
      args?.publicationId !== undefined ||
      args?.sourcePackageSha256 !== undefined ||
      args?.sourceAppPackageSha256 !== undefined ||
      args?.pluginSha256 !== undefined ||
      args?.sharePointPackageSha256 !== undefined ||
      args?.appPackageSha256 !== undefined
    );
  }

  private preparedError(
    error: Error,
    journal?: PreparedPublicationJournal,
    recoveryPath?: string,
    stage: PreparedPublicationErrorStage = "preparedPublication"
  ): FxError {
    const message = getLocalizedString("driver.teamsApp.prepared.error");
    const errorName =
      error instanceof PreparedPublicationError ? error.code : "PreparedPublicationFailed";
    const result = AppStudioResultFactory.UserError(errorName, [message, message]);
    result.data = createPreparedPublicationErrorData(error, journal, recoveryPath, stage);
    return result;
  }

  private isDefinitivePreparedRejection(error: unknown): boolean {
    if (!axios.isAxiosError(error) || error.response === undefined) {
      return false;
    }
    const status = error.response.status;
    return status >= 400 && status < 500 && status !== 408 && status !== 429;
  }

  private async publishPrepared(
    args: PublishAppPackageArgs,
    context: WrapDriverContext,
    outputEnvVarNames?: Map<string, string>,
    identityLocked = false
  ): Promise<Result<Map<string, string>, FxError>> {
    let journalForError: PreparedPublicationJournal | undefined;
    let currentStage: PreparedPublicationErrorStage = "preparedPublication";
    try {
      const preparedOutputNames = outputEnvVarNames ?? new Map(Object.entries(defaultOutputNames));
      const publishedAppOutputName = preparedOutputNames.get("publishedAppId");
      if (
        !publishedAppOutputName ||
        typeof args.appPackagePath !== "string" ||
        !args.appPackagePath
      ) {
        throw new PreparedPublicationError(
          "InvalidPreparedPublishArguments",
          "Prepared app package inputs and the published app ID output are required."
        );
      }
      if (
        process.env[FeatureFlagName.SovereignCloudEnvironment] === SovereignCloudEnvironment.GCCH ||
        process.env[FeatureFlagName.SovereignCloudEnvironment] === SovereignCloudEnvironment.DOD
      ) {
        throw new PreparedPublicationError(
          "UnsupportedCloud",
          "Prepared app package publication is not supported in this cloud."
        );
      }
      const requiredHashes = [
        args.sourcePackageSha256,
        args.sourceAppPackageSha256,
        args.pluginSha256,
        args.sharePointPackageSha256,
        args.appPackageSha256,
      ];
      if (
        typeof args.preparedPublicationPath !== "string" ||
        typeof args.publicationId !== "string" ||
        requiredHashes.some((hash) => typeof hash !== "string" || !/^[0-9a-f]{64}$/i.test(hash))
      ) {
        throw new PreparedPublicationError(
          "InvalidPreparedPublishArguments",
          "Prepared publication arguments are incomplete."
        );
      }
      const loaded = await loadPreparedPublication(
        context.projectPath,
        args.publicationId,
        args.preparedPublicationPath
      );
      if (loaded.isErr()) {
        throw loaded.error;
      }
      currentStage = "sharePointVerification";
      const journal = loaded.value.journal;
      journalForError = journal;
      if (!identityLocked) {
        return await withPreparedPublicationIdentityLock(
          loaded.value.root,
          loaded.value.journal,
          () => this.publishPrepared(args, context, outputEnvVarNames, true)
        );
      }
      if (
        args.sourcePackageSha256 !== journal.sourceSharePoint.sha256 ||
        args.sourceAppPackageSha256 !== journal.sourceApplication.sha256 ||
        args.pluginSha256 !== journal.plugin.sha256 ||
        args.sharePointPackageSha256 !== journal.sharePointArtifact.sha256 ||
        args.appPackageSha256 !== journal.applicationArtifact.sha256 ||
        isAmbiguousPublicationState(journal)
      ) {
        throw new PreparedPublicationError(
          "PreparedPublicationNotVerified",
          "The YAML action inputs do not match the prepared publication."
        );
      }
      const inputAppPath = await fs.realpath(
        getAbsolutePath(args.appPackagePath, context.projectPath)
      );
      const journalAppPath = await fs.realpath(
        path.resolve(loaded.value.root, ...journal.applicationArtifact.path.split("/"))
      );
      if (inputAppPath !== journalAppPath) {
        throw new PreparedPublicationError(
          "ApplicationArtifactPathMismatch",
          "The app package path does not match the prepared publication."
        );
      }
      const sources = await verifyPreparedSources(loaded.value.root, journal);
      if (sources.isErr()) {
        throw sources.error;
      }
      const appArtifact = await verifyPreparedArtifact(
        loaded.value.root,
        journal.applicationArtifact
      );
      if (appArtifact.isErr()) {
        throw appArtifact.error;
      }
      const appLineage = verifyPreparedApplicationPackage(journal, appArtifact.value);
      if (appLineage.isErr()) {
        throw appLineage.error;
      }
      if (
        journal.state === "submitted" &&
        journal.applicationArtifact.submissionState === "confirmed" &&
        typeof journal.applicationArtifact.submittedAppId === "string"
      ) {
        return ok(new Map([[publishedAppOutputName, journal.applicationArtifact.submittedAppId]]));
      }
      if (
        journal.state !== "sharePointVerified" ||
        journal.sharePointArtifact.deployState !== "confirmed"
      ) {
        throw new PreparedPublicationError(
          "PreparedPublicationNotVerified",
          "Application submission requires a SharePoint-verified prepared publication."
        );
      }
      const sharePointArtifact = await verifyPreparedArtifact(
        loaded.value.root,
        journal.sharePointArtifact
      );
      if (sharePointArtifact.isErr()) {
        throw sharePointArtifact.error;
      }
      const sharePointLineage = await verifyPreparedSharePointPackage(
        journal,
        sharePointArtifact.value
      );
      if (sharePointLineage.isErr()) {
        throw sharePointLineage.error;
      }
      if (!journal.sharePointArtifact.catalogUrl || !journal.sharePointArtifact.fileName) {
        throw new PreparedPublicationError(
          "MissingSharePointVerificationData",
          "The prepared SharePoint verification data is incomplete."
        );
      }
      if (context.signal?.aborted) {
        return err(new UserCancelError(actionName));
      }
      const tenant = await getSPOTenant(context.m365TokenProvider);
      const spoToken = await getSPFxToken(context.m365TokenProvider);
      if (!spoToken) {
        throw new PreparedPublicationError(
          "SharePointAuthenticationUnavailable",
          "A SharePoint access token is unavailable."
        );
      }
      const appCatalogSite = await SPOClient.getAppCatalogSiteForTenant(
        spoToken,
        new URL(tenant).origin,
        context.signal
      );
      if (
        !appCatalogSite ||
        !isPreparedCatalogUrl(appCatalogSite, tenant) ||
        this.normalizeCatalogUrl(appCatalogSite) !==
          this.normalizeCatalogUrl(journal.sharePointArtifact.catalogUrl)
      ) {
        throw new PreparedPublicationError(
          "SharePointCatalogChanged",
          "The current SharePoint app catalog does not match the prepared publication."
        );
      }
      await verifyPreparedJournalIdentity(
        context.m365TokenProvider,
        journal,
        tenant,
        appCatalogSite
      );
      const requestContext: PreparedSharePointRequestContext = {
        catalogUrl: appCatalogSite,
        signal: context.signal,
      };
      const remotePackage = await SPOClient.readPreparedAppPackage(
        spoToken,
        journal.sharePointArtifact.fileName,
        requestContext
      );
      if (!remotePackage) {
        throw new PreparedPublicationError(
          "SharePointPackageNotVisible",
          "The verified SharePoint package is no longer visible."
        );
      }
      const remotePackageVerification = await verifyPreparedSharePointPackage(
        journal,
        remotePackage
      );
      if (remotePackageVerification.isErr()) {
        throw remotePackageVerification.error;
      }
      const remoteApp = await SPOClient.getPreparedAvailableApp(
        spoToken,
        journal.sourceSharePoint.productId,
        requestContext
      );
      if (
        !remoteApp ||
        remoteApp.id.toLowerCase() !== journal.sourceSharePoint.productId.toLowerCase() ||
        !remoteApp.deployed ||
        !remoteApp.currentVersionDeployed ||
        remoteApp.skipFeatureDeployment !== journal.sourceSharePoint.skipFeatureDeployment
      ) {
        throw new PreparedPublicationError(
          "SharePointDeploymentNoLongerVerified",
          "SharePoint no longer reports the prepared package as the current deployed version."
        );
      }

      const graphTokenRes = await context.m365TokenProvider.getAccessToken({
        scopes: GraphTeamsAppCatalogReadWriteScopes,
      });
      if (graphTokenRes.isErr()) {
        throw graphTokenRes.error;
      }
      const graphClient = new GraphClient(context.m365TokenProvider, context.logProvider);
      currentStage = "applicationPreflight";
      const existingApp = await graphClient.getPreparedStagedApp(
        graphTokenRes.value,
        journal.applicationArtifact.appId,
        context.signal
      );
      if (existingApp) {
        if (
          !existingApp.definitions.some(
            (definition) => definition.publishingState === PublishingState.published
          ) ||
          existingApp.definitions.some(
            (definition) =>
              !semver.valid(definition.version) ||
              !semver.gt(journal.applicationArtifact.version, definition.version)
          )
        ) {
          throw new PreparedPublicationError(
            "TeamsAppVersionCollision",
            "The prepared app version must be newer than every existing app definition."
          );
        }
      }
      if (!context.ui) {
        throw new PreparedPublicationError(
          "ApprovalUnavailable",
          "Prepared publication requires an interactive approval surface."
        );
      }
      currentStage = "approval";
      const confirm = getLocalizedString("core.option.confirm");
      const approval = await context.ui.showMessage(
        "warn",
        getLocalizedString(
          "driver.teamsApp.prepared.approve",
          existingApp ? "update" : "create",
          journal.applicationArtifact.appId,
          journal.applicationArtifact.version
        ),
        true,
        confirm
      );
      if (!approval?.isOk() || approval.value !== confirm) {
        return err(new UserCancelError(actionName));
      }
      if (context.signal?.aborted) {
        return err(new UserCancelError(actionName));
      }
      const inFlight = await updatePreparedPublication(
        loaded.value.root,
        journal.publicationId,
        args.preparedPublicationPath,
        (current) => {
          if (
            current.state !== "sharePointVerified" ||
            current.applicationArtifact.submissionState !== "notStarted"
          ) {
            throw new PreparedPublicationError(
              "ConcurrentTeamsAppSubmission",
              "The prepared app submission state changed before the request."
            );
          }
          return {
            ...current,
            state: "appSubmissionInFlight",
            applicationArtifact: {
              ...current.applicationArtifact,
              submissionState: "inFlight",
            },
          };
        }
      );
      if (inFlight.isErr()) {
        throw inFlight.error;
      }
      if (context.signal?.aborted) {
        const cancelled = await updatePreparedPublication(
          loaded.value.root,
          journal.publicationId,
          args.preparedPublicationPath,
          (current) => {
            if (
              current.state !== "appSubmissionInFlight" ||
              current.applicationArtifact.submissionState !== "inFlight"
            ) {
              throw new PreparedPublicationError(
                "ConcurrentTeamsAppSubmission",
                "The prepared app submission state changed before cancellation settled."
              );
            }
            return {
              ...current,
              state: "sharePointVerified",
              applicationArtifact: {
                ...current.applicationArtifact,
                submissionState: "notStarted",
              },
            };
          }
        );
        if (cancelled.isErr()) {
          throw cancelled.error;
        }
        return err(new UserCancelError(actionName));
      }
      let publishedAppId: string;
      try {
        currentStage = "applicationSubmission";
        publishedAppId = await graphClient.publishPreparedTeamsApp(
          graphTokenRes.value,
          journal.applicationArtifact.appId,
          existingApp?.teamsAppId,
          appArtifact.value,
          context.signal
        );
      } catch (error) {
        const rejected = this.isDefinitivePreparedRejection(error);
        const failed = await updatePreparedPublication(
          loaded.value.root,
          journal.publicationId,
          args.preparedPublicationPath,
          (current) => {
            if (
              current.state !== "appSubmissionInFlight" ||
              current.applicationArtifact.submissionState !== "inFlight"
            ) {
              throw new PreparedPublicationError(
                "ConcurrentTeamsAppSubmission",
                "The prepared app submission state changed while the request was in progress."
              );
            }
            return {
              ...current,
              state: rejected ? "rejected" : "unknown",
              applicationArtifact: {
                ...current.applicationArtifact,
                submissionState: rejected ? "rejected" : "unknown",
              },
            };
          }
        );
        if (failed.isErr()) {
          throw failed.error;
        }
        throw new PreparedPublicationError(
          rejected ? "TeamsAppSubmissionRejected" : "TeamsAppSubmissionOutcomeUnknown",
          rejected
            ? "Graph rejected the prepared app package."
            : "The app package submission outcome is unknown; it will not be replayed.",
          { cause: error }
        );
      }
      const submitted = await updatePreparedPublication(
        loaded.value.root,
        journal.publicationId,
        args.preparedPublicationPath,
        (current) => {
          if (
            current.state !== "appSubmissionInFlight" ||
            current.applicationArtifact.submissionState !== "inFlight"
          ) {
            throw new PreparedPublicationError(
              "ConcurrentTeamsAppSubmission",
              "The prepared app submission state changed after the request."
            );
          }
          return {
            ...current,
            state: "submitted",
            applicationArtifact: {
              ...current.applicationArtifact,
              submissionState: "confirmed",
              submittedAppId: publishedAppId,
            },
          };
        }
      );
      if (submitted.isErr()) {
        throw submitted.error;
      }
      if (context.signal?.aborted) {
        return err(new UserCancelError(actionName));
      }
      merge(context.telemetryProperties, {
        [TelemetryPropertyKey.updateExistingApp]: existingApp ? "true" : "false",
        [TelemetryPropertyKey.publishedAppId]: publishedAppId,
      });
      context.addSummary(
        getLocalizedString(
          "driver.teamsApp.summary.publishTeamsAppSuccess",
          journal.applicationArtifact.appId
        )
      );
      return ok(new Map([[publishedAppOutputName, publishedAppId]]));
    } catch (error) {
      if (error instanceof UserCancelError) {
        return err(error);
      }
      let latestJournal = journalForError;
      if (
        typeof args.publicationId === "string" &&
        typeof args.preparedPublicationPath === "string"
      ) {
        const latest = await loadPreparedPublication(
          context.projectPath,
          args.publicationId,
          args.preparedPublicationPath
        );
        if (latest.isOk()) {
          latestJournal = latest.value.journal;
        }
      }
      if (context.signal?.aborted && latestJournal && !isAmbiguousPublicationState(latestJournal)) {
        return err(new UserCancelError(actionName));
      }
      if (
        (error instanceof UserError || error instanceof SystemError) &&
        error.data === undefined
      ) {
        error.data = createPreparedPublicationErrorData(
          error,
          latestJournal,
          args.preparedPublicationPath,
          currentStage
        );
        return err(error);
      }
      return err(
        this.preparedError(
          error instanceof Error ? error : new Error(),
          latestJournal,
          args.preparedPublicationPath,
          currentStage
        )
      );
    }
  }

  private normalizeCatalogUrl(value: string): string {
    const url = new URL(value);
    if (url.protocol !== "https:" || url.username || url.password || url.search || url.hash) {
      throw new PreparedPublicationError(
        "InvalidCatalogUrl",
        "The SharePoint app catalog URL is invalid."
      );
    }
    return `${url.origin}${url.pathname.replace(/\/+$/, "")}`.toLowerCase();
  }

  private validateArgs(args: PublishAppPackageArgs): Result<any, FxError> {
    const invalidParams: string[] = [];
    if (!args || !args.appPackagePath) {
      invalidParams.push("appPackagePath");
    }
    if (invalidParams.length > 0) {
      return err(
        new InvalidActionInputError(
          actionName,
          invalidParams,
          "https://aka.ms/teamsfx-actions/teamsapp-publish"
        )
      );
    } else {
      return ok(undefined);
    }
  }
}
