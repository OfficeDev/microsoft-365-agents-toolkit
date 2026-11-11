// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  FxError,
  M365TokenProvider,
  Platform,
  Result,
  UserError,
  UserErrorOptions,
} from "@microsoft/teamsfx-api";
import axios from "axios";
import fs from "fs-extra";
import path from "path";
import { Service } from "typedi";

import { GraphScopes } from "../../../../common/constants";
import { SovereignCloudEnvironment } from "../../../../common/accountUtils";
import { FeatureFlagName } from "../../../../common/featureFlags";
import { getLocalizedString } from "../../../../common/localizeUtils";
import { getSPFxToken } from "../../../../common/tools";
import { ErrorContextMW } from "../../../../common/globalVars";
import { FileNotFoundError, UserCancelError } from "../../../../error/common";
import { asBoolean, asFactory, asString, wrapRun } from "../../../utils/common";
import { DriverContext } from "../../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../../interface/stepDriver";
import { addStartAndEndTelemetry } from "../../middleware/addStartAndEndTelemetry";
import { WrapDriverContext } from "../../util/wrapUtil";
import { ErrorCategory } from "../../../../error/types";
import { CreateAppCatalogFailedError } from "./error/createAppCatalogFailedError";
import { GetGraphTokenFailedError } from "./error/getGraphTokenFailedError";
import { GetSPOTokenFailedError } from "./error/getSPOTokenFailedError";
import { GetTenantFailedError } from "./error/getTenantFailedError";
import { InsufficientPermissionError } from "./error/insufficientPermissionError";
import { NoSPPackageError } from "./error/noSPPackageError";
import { NoValidAppCatelog } from "./error/noValidAppCatelogError";
import { UploadAppPackageFailedError } from "./error/uploadAppPackageFailedError";
import { DeploySPFxArgs } from "./interface/deployArgs";
import { Constants, DeployProgressMessage } from "./utility/constants";
import { sleep } from "./utility/sleep";
import { PreparedSharePointRequestContext, SPOClient } from "./utility/spoClient";
import { inspectSharePointPackageIdentity } from "./utility/packageFile";
import {
  isAmbiguousPublicationState,
  compareSharePointPackageVersions,
  createPreparedPublicationErrorData,
  loadPreparedPublication,
  PreparedPublicationError,
  PreparedPublicationErrorStage,
  PreparedPublicationJournal,
  isPreparedCatalogUrl,
  updatePreparedPublication,
  verifyPreparedArtifact,
  verifyPreparedApplicationPackage,
  verifyPreparedSharePointPackage,
  verifyPreparedSources,
  withPreparedPublicationIdentityLock,
} from "./utility/preparedPublication";
import { verifyPreparedJournalIdentity } from "./utility/tokenIdentity";

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

@Service(Constants.DeployDriverName)
export class SPFxDeployDriver implements StepDriver {
  public readonly description = getLocalizedString("driver.spfx.deploy.description");
  public readonly progressTitle = getLocalizedString("driver.spfx.deploy.progressbar.stepMessage");

  private readonly EmptyMap = new Map<string, string>();

  private asDeployArgs = asFactory<
    Pick<DeploySPFxArgs, "createAppCatalogIfNotExist" | "packageSolutionPath">
  >({
    createAppCatalogIfNotExist: asBoolean,
    packageSolutionPath: asString,
  });

  @hooks([
    addStartAndEndTelemetry(Constants.TelemetryDeployEventName, Constants.TelemetryComponentName),
  ])
  public async run(
    args: DeploySPFxArgs,
    context: DriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    const wrapContext = new WrapDriverContext(
      context,
      Constants.TelemetryDeployEventName,
      Constants.TelemetryComponentName
    );
    return wrapRun(() => this.deploy(args, wrapContext), Constants.DeployDriverName);
  }

  public async execute(args: DeploySPFxArgs, ctx: DriverContext): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(
      ctx,
      Constants.TelemetryDeployEventName,
      Constants.TelemetryComponentName
    );
    const result = await this.run(args, wrapContext);
    return {
      result,
      summaries: wrapContext.summaries,
    };
  }
  @hooks([ErrorContextMW({ source: "SPFx", component: "SPFxDeployDriver" })])
  public async deploy(
    args: DeploySPFxArgs,
    context: WrapDriverContext
  ): Promise<Map<string, string>> {
    if (
      args?.preparedPublicationPath !== undefined ||
      args?.publicationId !== undefined ||
      args?.preparedPackagePath !== undefined ||
      args?.sourcePackageSha256 !== undefined ||
      args?.sourceAppPackageSha256 !== undefined ||
      args?.pluginSha256 !== undefined ||
      args?.sharePointPackageSha256 !== undefined ||
      args?.appPackageSha256 !== undefined
    ) {
      try {
        return await this.deployPreparedPublication(args, context);
      } catch (error) {
        if (error instanceof UserCancelError) {
          throw error;
        }
        let journal: PreparedPublicationJournal | undefined;
        if (
          typeof args.publicationId === "string" &&
          typeof args.preparedPublicationPath === "string"
        ) {
          const loaded = await loadPreparedPublication(
            context.projectPath,
            args.publicationId,
            args.preparedPublicationPath
          );
          if (loaded.isOk()) {
            journal = loaded.value.journal;
          }
        }
        if (context.signal?.aborted && journal && !isAmbiguousPublicationState(journal)) {
          throw new UserCancelError(Constants.DeployDriverName);
        }
        const cause =
          error instanceof UserError && error.innerError instanceof Error
            ? error.innerError
            : error instanceof Error
              ? error
              : new Error("Prepared SharePoint deployment failed.");
        const stage = this.preparedErrorStage(cause);
        throw this.toPreparedError(
          cause,
          Constants.DeployDriverName,
          journal,
          args.preparedPublicationPath,
          stage
        );
      }
    }
    const deployArgs = this.asDeployArgs(args);

    context.logProvider.debug(`Getting user tenant...`);
    const tenant = await this.getTenant(context.m365TokenProvider);
    SPOClient.setBaseUrl(tenant);
    context.logProvider.debug(`Succeeded to get user tenant: ${tenant}.`);

    const spoToken = await getSPFxToken(context.m365TokenProvider);
    if (!spoToken) {
      throw new GetSPOTokenFailedError();
    }

    context.logProvider.verbose(`Getting SharePoint app catalog site...`);
    let appCatalogSite = await SPOClient.getAppCatalogSite(spoToken);
    if (appCatalogSite) {
      context.logProvider.verbose(
        `Succeeded to get SharePoint app catalog site: ${appCatalogSite}.`
      );
      SPOClient.setBaseUrl(appCatalogSite);
      context.addSummary(DeployProgressMessage.SkipCreateSPAppCatalog());
    } else {
      context.logProvider.verbose(
        `Failed to get valid SharePoint app catalog site under current tenant.`
      );
      if (deployArgs.createAppCatalogIfNotExist) {
        context.logProvider.verbose(
          `Creating app catalog for user since there's no existing one...`
        );
        try {
          await SPOClient.createAppCatalog(spoToken);
          context.addSummary(DeployProgressMessage.CreateSPAppCatalog());
          context.logProvider.verbose(`Succeeded to create app catalog.`);
        } catch (e) {
          throw new CreateAppCatalogFailedError(e as Error);
        }
      } else {
        throw new NoValidAppCatelog();
      }
      let retry = 0;
      context.logProvider.verbose(`Getting newly created app catalog site...`);
      appCatalogSite = await SPOClient.getAppCatalogSite(spoToken);
      while (appCatalogSite == null && retry < Constants.APP_CATALOG_MAX_TIMES) {
        void context.logProvider.warning(
          getLocalizedString("driver.spfx.warn.noTenantAppCatalogFound", retry)
        );
        await sleep(Constants.APP_CATALOG_REFRESH_TIME);
        appCatalogSite = await SPOClient.getAppCatalogSite(spoToken);
        retry += 1;
      }
      if (appCatalogSite) {
        context.logProvider.verbose(
          `Succeeded to get newly created app catalog site: ${appCatalogSite}.`
        );
        SPOClient.setBaseUrl(appCatalogSite);
        void context.logProvider.info(
          getLocalizedString("driver.spfx.info.tenantAppCatalogCreated", appCatalogSite)
        );
        await sleep(Constants.APP_CATALOG_ACTIVE_TIME);
      } else {
        throw new CreateAppCatalogFailedError(
          new Error(getLocalizedString("driver.spfx.error.failedToGetAppCatalog"))
        );
      }
    }

    const packageSolutionPath = path.isAbsolute(deployArgs.packageSolutionPath)
      ? deployArgs.packageSolutionPath
      : path.join(context.projectPath, deployArgs.packageSolutionPath);
    context.logProvider.debug(
      `Getting zipped package path from package-solution.json file under ${packageSolutionPath}...`
    );
    const appPackage = await this.getPackagePath(packageSolutionPath);
    if (!(await fs.pathExists(appPackage))) {
      throw new NoSPPackageError(appPackage);
    }
    context.logProvider.debug(`Succeeded to get zipped package path: ${appPackage}.`);

    const fileName = path.parse(appPackage).base;
    const bytes = await fs.readFile(appPackage);
    try {
      context.logProvider.verbose(`Uploading SharePoint app package ${fileName}...`);
      await SPOClient.uploadAppPackage(spoToken, fileName, bytes);
      context.addSummary(DeployProgressMessage.Upload());
      context.logProvider.verbose(`Succeeded to upload SharePoint app package.`);
    } catch (e: any) {
      if (e.response?.status === 403) {
        throw new InsufficientPermissionError(appCatalogSite);
      } else {
        throw new UploadAppPackageFailedError(e);
      }
    }

    context.logProvider.debug(
      `Getting app id from package-solution.json file under ${packageSolutionPath}...`
    );
    const appID = await this.getAppID(packageSolutionPath);
    context.logProvider.verbose(`Deploying SharePoint app package with app id: ${appID}...`);
    await SPOClient.deployAppPackage(spoToken, appID);
    context.addSummary(DeployProgressMessage.Deploy());
    context.logProvider.verbose(`Succeeded to deploy SharePoint app package.`);
    const guidance = getLocalizedString(
      "plugins.spfx.deployNotice",
      appPackage,
      appCatalogSite,
      appCatalogSite
    );
    if (context.platform === Platform.CLI) {
      context.ui?.showMessage("info", guidance, false);
    } else {
      context.ui?.showMessage("info", guidance, false, "OK");
    }
    return this.EmptyMap;
  }

  private toPreparedError(
    error: Error,
    source: string,
    journal?: PreparedPublicationJournal,
    recoveryPath?: string,
    stage: PreparedPublicationErrorStage = this.preparedErrorStage(error)
  ): UserError {
    const options: UserErrorOptions = {
      source,
      error,
      message: error instanceof PreparedPublicationError ? error.code : "PreparedPublicationFailed",
      displayMessage: getLocalizedString("driver.spfx.prepared.error"),
      categories: [ErrorCategory.External],
    };
    const result = new UserError(options);
    result.data = createPreparedPublicationErrorData(error, journal, recoveryPath, stage);
    return result;
  }

  private preparedErrorStage(error: Error): PreparedPublicationErrorStage {
    if (error instanceof PreparedPublicationError) {
      if (error.code.startsWith("SharePointUpload")) {
        return "sharePointUpload";
      }
      if (error.code.startsWith("SharePointDeploy")) {
        return "sharePointDeploy";
      }
      if (error.code.startsWith("SharePoint")) {
        return "sharePointVerification";
      }
    }
    return "preparedPublication";
  }

  private async deployPreparedPublication(
    args: DeploySPFxArgs,
    context: WrapDriverContext,
    identityLocked = false
  ): Promise<Map<string, string>> {
    if (
      process.env[FeatureFlagName.SovereignCloudEnvironment] === SovereignCloudEnvironment.GCCH ||
      process.env[FeatureFlagName.SovereignCloudEnvironment] === SovereignCloudEnvironment.DOD
    ) {
      throw this.toPreparedError(
        new PreparedPublicationError(
          "UnsupportedCloud",
          "Prepared SharePoint publication is not supported in this cloud."
        ),
        Constants.DeployDriverName
      );
    }
    if (
      typeof args.preparedPublicationPath !== "string" ||
      typeof args.publicationId !== "string" ||
      typeof args.preparedPackagePath !== "string" ||
      args.createAppCatalogIfNotExist === true
    ) {
      throw this.toPreparedError(
        new PreparedPublicationError(
          "InvalidPreparedDeployArguments",
          "Prepared deployment requires a publication journal and must not create an app catalog."
        ),
        Constants.DeployDriverName
      );
    }
    const loaded = await loadPreparedPublication(
      context.projectPath,
      args.publicationId,
      args.preparedPublicationPath
    );
    if (loaded.isErr()) {
      throw this.toPreparedError(loaded.error, Constants.DeployDriverName);
    }
    if (!identityLocked) {
      try {
        return await withPreparedPublicationIdentityLock(
          loaded.value.root,
          loaded.value.journal,
          () => this.deployPreparedPublication(args, context, true)
        );
      } catch (error) {
        if (error instanceof PreparedPublicationError) {
          throw this.toPreparedError(error, Constants.DeployDriverName);
        }
        throw error;
      }
    }
    let journal = loaded.value.journal;
    if (
      args.sourcePackageSha256 !== journal.sourceSharePoint.sha256 ||
      args.sourceAppPackageSha256 !== journal.sourceApplication.sha256 ||
      args.pluginSha256 !== journal.plugin.sha256 ||
      args.sharePointPackageSha256 !== journal.sharePointArtifact.sha256 ||
      args.appPackageSha256 !== journal.applicationArtifact.sha256
    ) {
      throw this.toPreparedError(
        new PreparedPublicationError(
          "PreparedPublicationHashBindingMismatch",
          "The YAML action inputs do not match the prepared publication hashes."
        ),
        Constants.DeployDriverName
      );
    }
    if (args.preparedPackagePath !== journal.sharePointArtifact.path) {
      throw this.toPreparedError(
        new PreparedPublicationError(
          "SharePointArtifactPathMismatch",
          "The YAML action package path does not match the prepared publication."
        ),
        Constants.DeployDriverName
      );
    }
    const sources = await verifyPreparedSources(loaded.value.root, journal);
    if (sources.isErr()) {
      throw this.toPreparedError(sources.error, Constants.DeployDriverName);
    }
    const localArtifact = await verifyPreparedArtifact(
      loaded.value.root,
      journal.sharePointArtifact
    );
    if (localArtifact.isErr()) {
      throw this.toPreparedError(localArtifact.error, Constants.DeployDriverName);
    }
    const identity = await verifyPreparedSharePointPackage(journal, localArtifact.value);
    if (identity.isErr()) {
      throw this.toPreparedError(identity.error, Constants.DeployDriverName);
    }
    const applicationArtifact = await verifyPreparedArtifact(
      loaded.value.root,
      journal.applicationArtifact
    );
    if (applicationArtifact.isErr()) {
      throw this.toPreparedError(applicationArtifact.error, Constants.DeployDriverName);
    }
    const applicationLineage = verifyPreparedApplicationPackage(journal, applicationArtifact.value);
    if (applicationLineage.isErr()) {
      throw this.toPreparedError(applicationLineage.error, Constants.DeployDriverName);
    }
    if (
      journal.state === "submitted" &&
      journal.sharePointArtifact.uploadState === "confirmed" &&
      journal.sharePointArtifact.deployState === "confirmed" &&
      journal.applicationArtifact.submissionState === "confirmed" &&
      typeof journal.applicationArtifact.submittedAppId === "string" &&
      journal.applicationArtifact.submittedAppId.length > 0
    ) {
      return this.EmptyMap;
    }
    if (journal.state === "sharePointVerified") {
      return this.EmptyMap;
    }
    if (isAmbiguousPublicationState(journal)) {
      throw this.toPreparedError(
        new PreparedPublicationError(
          "PreparedSharePointOutcomeUnknown",
          "A previous SharePoint request has an unresolved outcome and will not be replayed."
        ),
        Constants.DeployDriverName
      );
    }
    if (journal.state !== "prepared" && journal.state !== "sharePointUploaded") {
      throw this.toPreparedError(
        new PreparedPublicationError(
          "InvalidPreparedSharePointState",
          "The prepared SharePoint publication is not deployable."
        ),
        Constants.DeployDriverName
      );
    }

    if (context.signal?.aborted) {
      throw new UserCancelError(Constants.DeployDriverName);
    }
    const tenant = await this.getTenant(context.m365TokenProvider);
    const spoToken = await getSPFxToken(context.m365TokenProvider);
    if (!spoToken) {
      throw new GetSPOTokenFailedError();
    }
    const appCatalogSite = await SPOClient.getAppCatalogSiteForTenant(
      spoToken,
      new URL(tenant).origin,
      context.signal
    );
    if (!appCatalogSite || !isPreparedCatalogUrl(appCatalogSite, tenant)) {
      throw this.toPreparedError(
        new PreparedPublicationError(
          "TenantAppCatalogUnavailable",
          "A supported tenant app catalog could not be discovered."
        ),
        Constants.DeployDriverName
      );
    }
    try {
      await verifyPreparedJournalIdentity(
        context.m365TokenProvider,
        journal,
        tenant,
        appCatalogSite
      );
    } catch (error) {
      if (error instanceof PreparedPublicationError) {
        throw this.toPreparedError(error, Constants.DeployDriverName);
      }
      throw error;
    }
    const requestContext: PreparedSharePointRequestContext = {
      catalogUrl: appCatalogSite,
      signal: context.signal,
    };

    if (journal.state === "prepared") {
      if (journal.sharePointArtifact.uploadState !== "notStarted") {
        throw this.toPreparedError(
          new PreparedPublicationError(
            "InvalidPreparedUploadState",
            "The prepared SharePoint upload state is invalid."
          ),
          Constants.DeployDriverName
        );
      }
      const fileName = path.posix.basename(journal.sourceSharePoint.path);
      const existingPackage = await SPOClient.readPreparedAppPackage(
        spoToken,
        fileName,
        requestContext
      );
      const existing = await SPOClient.getPreparedAvailableApp(
        spoToken,
        journal.sourceSharePoint.productId,
        requestContext
      );
      let overwrite = false;
      if (existingPackage) {
        const existingIdentity = await inspectSharePointPackageIdentity(existingPackage);
        if (
          existingIdentity.isErr() ||
          existingIdentity.value.productId.toLowerCase() !==
            journal.sourceSharePoint.productId.toLowerCase()
        ) {
          throw this.toPreparedError(
            new PreparedPublicationError(
              "SharePointPackageNameCollision",
              "The SharePoint package file name is already used by a different or invalid solution."
            ),
            Constants.DeployDriverName
          );
        }
        const versionComparison = compareSharePointPackageVersions(
          existingIdentity.value.version,
          journal.sourceSharePoint.version
        );
        if (versionComparison === undefined) {
          throw this.toPreparedError(
            new PreparedPublicationError(
              "InvalidSharePointCatalogVersion",
              "The existing SharePoint package version cannot be compared safely."
            ),
            Constants.DeployDriverName
          );
        }
        if (existing && existing.version !== existingIdentity.value.version) {
          throw this.toPreparedError(
            new PreparedPublicationError(
              "SharePointCatalogVersionMismatch",
              "The app catalog version does not match the package stored under its file name."
            ),
            Constants.DeployDriverName
          );
        }
        if (versionComparison > 0) {
          throw this.toPreparedError(
            new PreparedPublicationError(
              "SharePointPackageDowngrade",
              "The prepared SharePoint package version is older than the app catalog version."
            ),
            Constants.DeployDriverName
          );
        }
        if (versionComparison === 0) {
          const verified = await verifyPreparedSharePointPackage(journal, existingPackage);
          if (verified.isErr()) {
            throw this.toPreparedError(
              new PreparedPublicationError(
                "SharePointPackageVersionCollision",
                "The app catalog already contains different package content at this version."
              ),
              Constants.DeployDriverName
            );
          }
          const reused = await updatePreparedPublication(
            loaded.value.root,
            journal.publicationId,
            args.preparedPublicationPath,
            (current) => {
              if (
                current.state !== "prepared" ||
                current.sharePointArtifact.uploadState !== "notStarted"
              ) {
                throw new PreparedPublicationError(
                  "ConcurrentPreparedUpload",
                  "The prepared SharePoint upload state changed before reuse."
                );
              }
              return {
                ...current,
                state: "sharePointUploaded",
                sharePointArtifact: {
                  ...current.sharePointArtifact,
                  uploadState: "confirmed",
                  catalogUrl: appCatalogSite,
                  fileName,
                  ...(existing?.deployed &&
                  existing.currentVersionDeployed &&
                  existing.skipFeatureDeployment === current.sourceSharePoint.skipFeatureDeployment
                    ? { deployState: "confirmed" }
                    : {}),
                },
              };
            }
          );
          if (reused.isErr()) {
            throw this.toPreparedError(reused.error, Constants.DeployDriverName);
          }
          journal = reused.value;
        } else {
          overwrite = true;
        }
      } else if (existing) {
        throw this.toPreparedError(
          new PreparedPublicationError(
            "SharePointProductIdCollision",
            "An existing SharePoint solution cannot be matched to the package file, so it will not be replaced."
          ),
          Constants.DeployDriverName
        );
      }
      if (journal.state === "prepared") {
        await this.requirePreparedApproval(
          context,
          getLocalizedString(
            "driver.spfx.prepared.approveUpload",
            journal.sourceSharePoint.productId,
            journal.sourceSharePoint.version,
            appCatalogSite
          )
        );
        const prepared = await updatePreparedPublication(
          loaded.value.root,
          journal.publicationId,
          args.preparedPublicationPath,
          (current) => {
            if (
              current.state !== "prepared" ||
              current.sharePointArtifact.uploadState !== "notStarted"
            ) {
              throw new PreparedPublicationError(
                "ConcurrentPreparedUpload",
                "The prepared SharePoint upload state changed before the request."
              );
            }
            return {
              ...current,
              state: "sharePointUploadInFlight",
              sharePointArtifact: {
                ...current.sharePointArtifact,
                uploadState: "inFlight",
                catalogUrl: appCatalogSite,
                fileName,
              },
            };
          }
        );
        if (prepared.isErr()) {
          throw this.toPreparedError(prepared.error, Constants.DeployDriverName);
        }
        journal = prepared.value;
        if (context.signal?.aborted) {
          const cancelled = await updatePreparedPublication(
            loaded.value.root,
            journal.publicationId,
            args.preparedPublicationPath,
            (current) => {
              if (
                current.state !== "sharePointUploadInFlight" ||
                current.sharePointArtifact.uploadState !== "inFlight"
              ) {
                throw new PreparedPublicationError(
                  "ConcurrentPreparedUpload",
                  "The prepared SharePoint upload state changed before cancellation settled."
                );
              }
              return {
                ...current,
                state: "prepared",
                sharePointArtifact: {
                  ...current.sharePointArtifact,
                  uploadState: "notStarted",
                  catalogUrl: undefined,
                  fileName: undefined,
                },
              };
            }
          );
          if (cancelled.isErr()) {
            throw this.toPreparedError(cancelled.error, Constants.DeployDriverName);
          }
          throw new UserCancelError(Constants.DeployDriverName);
        }
        try {
          await SPOClient.uploadPreparedAppPackage(
            spoToken,
            fileName,
            localArtifact.value,
            overwrite,
            requestContext
          );
        } catch (error) {
          const rejected = this.isDefinitiveRejection(error);
          const failed = await updatePreparedPublication(
            loaded.value.root,
            journal.publicationId,
            args.preparedPublicationPath,
            (current) => {
              if (
                current.state !== "sharePointUploadInFlight" ||
                current.sharePointArtifact.uploadState !== "inFlight"
              ) {
                throw new PreparedPublicationError(
                  "ConcurrentPreparedUpload",
                  "The prepared SharePoint upload state changed while the request was in progress."
                );
              }
              return {
                ...current,
                state: rejected ? "rejected" : "unknown",
                sharePointArtifact: {
                  ...current.sharePointArtifact,
                  uploadState: rejected ? "rejected" : "unknown",
                },
              };
            }
          );
          if (failed.isErr()) {
            throw this.toPreparedError(failed.error, Constants.DeployDriverName);
          }
          throw this.toPreparedError(
            new PreparedPublicationError(
              rejected ? "SharePointUploadRejected" : "SharePointUploadOutcomeUnknown",
              rejected
                ? "SharePoint rejected the prepared package upload."
                : "The SharePoint package upload outcome is unknown; it will not be replayed.",
              { cause: error }
            ),
            Constants.DeployDriverName
          );
        }
        const uploaded = await updatePreparedPublication(
          loaded.value.root,
          journal.publicationId,
          args.preparedPublicationPath,
          (current) => {
            if (
              current.state !== "sharePointUploadInFlight" ||
              current.sharePointArtifact.uploadState !== "inFlight"
            ) {
              throw new PreparedPublicationError(
                "ConcurrentPreparedUpload",
                "The prepared SharePoint upload state changed after the request."
              );
            }
            return {
              ...current,
              state: "sharePointUploaded",
              sharePointArtifact: {
                ...current.sharePointArtifact,
                uploadState: "confirmed",
              },
            };
          }
        );
        if (uploaded.isErr()) {
          throw this.toPreparedError(uploaded.error, Constants.DeployDriverName);
        }
        journal = uploaded.value;
      }
    }

    if (
      journal.sharePointArtifact.uploadState !== "confirmed" ||
      typeof journal.sharePointArtifact.fileName !== "string" ||
      journal.sharePointArtifact.catalogUrl !== appCatalogSite
    ) {
      throw this.toPreparedError(
        new PreparedPublicationError(
          "InvalidPreparedUploadState",
          "The prepared SharePoint upload cannot be verified."
        ),
        Constants.DeployDriverName
      );
    }
    await this.waitForPreparedPackage(
      spoToken,
      journal.sharePointArtifact.fileName,
      journal,
      requestContext
    );
    const availableApp = await this.waitForPreparedAvailableApp(
      spoToken,
      journal.sourceSharePoint.productId,
      requestContext
    );
    if (
      availableApp.id.toLowerCase() !== journal.sourceSharePoint.productId.toLowerCase() ||
      availableApp.version !== journal.sourceSharePoint.version ||
      availableApp.skipFeatureDeployment !== journal.sourceSharePoint.skipFeatureDeployment
    ) {
      throw this.toPreparedError(
        new PreparedPublicationError(
          "SharePointCatalogIdentityMismatch",
          "The uploaded package does not match the catalog identity."
        ),
        Constants.DeployDriverName
      );
    }

    if (journal.sharePointArtifact.deployState === "notStarted") {
      await this.requirePreparedApproval(
        context,
        getLocalizedString(
          "driver.spfx.prepared.approveDeploy",
          journal.sourceSharePoint.productId,
          journal.sourceSharePoint.version,
          appCatalogSite
        )
      );
      const inFlight = await updatePreparedPublication(
        loaded.value.root,
        journal.publicationId,
        args.preparedPublicationPath,
        (current) => {
          if (
            current.state !== "sharePointUploaded" ||
            current.sharePointArtifact.deployState !== "notStarted"
          ) {
            throw new PreparedPublicationError(
              "ConcurrentPreparedDeploy",
              "The prepared SharePoint deployment state changed before the request."
            );
          }
          return {
            ...current,
            state: "sharePointDeployInFlight",
            sharePointArtifact: {
              ...current.sharePointArtifact,
              deployState: "inFlight",
            },
          };
        }
      );
      if (inFlight.isErr()) {
        throw this.toPreparedError(inFlight.error, Constants.DeployDriverName);
      }
      if (context.signal?.aborted) {
        const cancelled = await updatePreparedPublication(
          loaded.value.root,
          journal.publicationId,
          args.preparedPublicationPath,
          (current) => {
            if (
              current.state !== "sharePointDeployInFlight" ||
              current.sharePointArtifact.deployState !== "inFlight"
            ) {
              throw new PreparedPublicationError(
                "ConcurrentPreparedDeploy",
                "The prepared SharePoint deploy state changed before cancellation settled."
              );
            }
            return {
              ...current,
              state: "sharePointUploaded",
              sharePointArtifact: {
                ...current.sharePointArtifact,
                deployState: "notStarted",
              },
            };
          }
        );
        if (cancelled.isErr()) {
          throw this.toPreparedError(cancelled.error, Constants.DeployDriverName);
        }
        throw new UserCancelError(Constants.DeployDriverName);
      }
      try {
        await SPOClient.deployPreparedAppPackage(
          spoToken,
          journal.sourceSharePoint.productId,
          journal.sourceSharePoint.skipFeatureDeployment,
          requestContext
        );
      } catch (error) {
        const rejected = this.isDefinitiveRejection(error);
        const failed = await updatePreparedPublication(
          loaded.value.root,
          journal.publicationId,
          args.preparedPublicationPath,
          (current) => {
            if (
              current.state !== "sharePointDeployInFlight" ||
              current.sharePointArtifact.deployState !== "inFlight"
            ) {
              throw new PreparedPublicationError(
                "ConcurrentPreparedDeploy",
                "The prepared SharePoint deployment state changed while the request was in progress."
              );
            }
            return {
              ...current,
              state: rejected ? "rejected" : "unknown",
              sharePointArtifact: {
                ...current.sharePointArtifact,
                deployState: rejected ? "rejected" : "unknown",
              },
            };
          }
        );
        if (failed.isErr()) {
          throw this.toPreparedError(failed.error, Constants.DeployDriverName);
        }
        throw this.toPreparedError(
          new PreparedPublicationError(
            rejected ? "SharePointDeployRejected" : "SharePointDeployOutcomeUnknown",
            rejected
              ? "SharePoint rejected the prepared package deployment."
              : "The SharePoint deployment outcome is unknown; it will not be replayed.",
            { cause: error }
          ),
          Constants.DeployDriverName
        );
      }
      const confirmed = await updatePreparedPublication(
        loaded.value.root,
        journal.publicationId,
        args.preparedPublicationPath,
        (current) => {
          if (
            current.state !== "sharePointDeployInFlight" ||
            current.sharePointArtifact.deployState !== "inFlight"
          ) {
            throw new PreparedPublicationError(
              "ConcurrentPreparedDeploy",
              "The prepared SharePoint deployment state changed after the request."
            );
          }
          return {
            ...current,
            state: "sharePointUploaded",
            sharePointArtifact: {
              ...current.sharePointArtifact,
              deployState: "confirmed",
            },
          };
        }
      );
      if (confirmed.isErr()) {
        throw this.toPreparedError(confirmed.error, Constants.DeployDriverName);
      }
      journal = confirmed.value;
      if (context.signal?.aborted) {
        throw new UserCancelError(Constants.DeployDriverName);
      }
    } else if (journal.sharePointArtifact.deployState !== "confirmed") {
      throw this.toPreparedError(
        new PreparedPublicationError(
          "InvalidPreparedDeployState",
          "The prepared SharePoint deployment state is invalid."
        ),
        Constants.DeployDriverName
      );
    }

    await this.verifyPreparedDeployment(spoToken, journal, requestContext);
    const verified = await updatePreparedPublication(
      loaded.value.root,
      journal.publicationId,
      args.preparedPublicationPath,
      (current) => {
        if (
          current.state !== "sharePointUploaded" ||
          current.sharePointArtifact.deployState !== "confirmed"
        ) {
          throw new PreparedPublicationError(
            "ConcurrentPreparedDeploy",
            "The prepared SharePoint deployment state changed before verification completed."
          );
        }
        return { ...current, state: "sharePointVerified" };
      }
    );
    if (verified.isErr()) {
      throw this.toPreparedError(verified.error, Constants.DeployDriverName);
    }
    context.addSummary(getLocalizedString("driver.spfx.prepared.sharePointVerified"));
    return this.EmptyMap;
  }

  private async requirePreparedApproval(
    context: WrapDriverContext,
    message: string
  ): Promise<void> {
    if (!context.ui) {
      throw this.toPreparedError(
        new PreparedPublicationError(
          "ApprovalUnavailable",
          "Prepared publication requires an interactive approval surface."
        ),
        Constants.DeployDriverName
      );
    }
    const confirm = getLocalizedString("core.option.confirm");
    const response = await context.ui.showMessage("warn", message, true, confirm);
    if (!response?.isOk() || response.value !== confirm) {
      throw new UserCancelError(Constants.DeployDriverName);
    }
  }

  private async waitForPreparedPackage(
    token: string,
    fileName: string,
    journal: PreparedPublicationJournal,
    requestContext: PreparedSharePointRequestContext
  ): Promise<Buffer> {
    for (let attempt = 0; attempt < Constants.APP_CATALOG_MAX_TIMES; attempt += 1) {
      if (requestContext.signal?.aborted) {
        throw new UserCancelError(Constants.DeployDriverName);
      }
      const bytes = await SPOClient.readPreparedAppPackage(token, fileName, requestContext);
      if (bytes) {
        const verified = await verifyPreparedSharePointPackage(journal, bytes);
        if (verified.isErr()) {
          throw verified.error;
        }
        return bytes;
      }
      if (attempt + 1 < Constants.APP_CATALOG_MAX_TIMES) {
        await sleep(Constants.APP_CATALOG_REFRESH_TIME);
      }
    }
    throw new PreparedPublicationError(
      "SharePointPackageNotVisible",
      "The uploaded package did not become visible within the bounded verification period."
    );
  }

  private async waitForPreparedAvailableApp(
    token: string,
    productId: string,
    requestContext: PreparedSharePointRequestContext
  ): Promise<NonNullable<Awaited<ReturnType<typeof SPOClient.getPreparedAvailableApp>>>> {
    for (let attempt = 0; attempt < Constants.APP_CATALOG_MAX_TIMES; attempt += 1) {
      if (requestContext.signal?.aborted) {
        throw new UserCancelError(Constants.DeployDriverName);
      }
      const app = await SPOClient.getPreparedAvailableApp(token, productId, requestContext);
      if (app) {
        return app;
      }
      if (attempt + 1 < Constants.APP_CATALOG_MAX_TIMES) {
        await sleep(Constants.APP_CATALOG_REFRESH_TIME);
      }
    }
    throw new PreparedPublicationError(
      "SharePointAppNotVisible",
      "The uploaded package did not become visible in the app catalog."
    );
  }

  private async verifyPreparedDeployment(
    token: string,
    journal: PreparedPublicationJournal,
    requestContext: PreparedSharePointRequestContext
  ): Promise<void> {
    const fileName = journal.sharePointArtifact.fileName;
    if (typeof fileName !== "string") {
      throw new PreparedPublicationError(
        "MissingPreparedPackageName",
        "The prepared SharePoint package name is missing."
      );
    }
    for (let attempt = 0; attempt < Constants.APP_CATALOG_MAX_TIMES; attempt += 1) {
      if (requestContext.signal?.aborted) {
        throw new UserCancelError(Constants.DeployDriverName);
      }
      const bytes = await SPOClient.readPreparedAppPackage(token, fileName, requestContext);
      const app = await SPOClient.getPreparedAvailableApp(
        token,
        journal.sourceSharePoint.productId,
        requestContext
      );
      if (bytes && app) {
        const verified = await verifyPreparedSharePointPackage(journal, bytes);
        if (verified.isErr()) {
          throw verified.error;
        }
        if (
          app.id.toLowerCase() === journal.sourceSharePoint.productId.toLowerCase() &&
          app.deployed &&
          app.currentVersionDeployed &&
          app.skipFeatureDeployment === journal.sourceSharePoint.skipFeatureDeployment
        ) {
          return;
        }
      }
      if (attempt + 1 < Constants.APP_CATALOG_MAX_TIMES) {
        await sleep(Constants.APP_CATALOG_REFRESH_TIME);
      }
    }
    throw new PreparedPublicationError(
      "SharePointDeploymentNotVerified",
      "SharePoint did not report the prepared package as the current deployed version."
    );
  }

  private isDefinitiveRejection(error: unknown): boolean {
    if (!axios.isAxiosError(error) || error.response === undefined) {
      return false;
    }
    const status = error.response.status;
    return status >= 400 && status < 500 && status !== 408 && status !== 429;
  }

  public async getTenant(tokenProvider: M365TokenProvider): Promise<string> {
    const graphTokenRes = await tokenProvider.getAccessToken({
      scopes: GraphScopes,
    });
    const graphToken = graphTokenRes.isOk() ? graphTokenRes.value : undefined;
    if (!graphToken) {
      throw new GetGraphTokenFailedError();
    }

    const tokenJsonRes = await tokenProvider.getJsonObject({
      scopes: GraphScopes,
    });
    const tokenJson: unknown = tokenJsonRes.isOk() ? tokenJsonRes.value : undefined;
    const tokenValue = isRecord(tokenJson) ? tokenJson["value"] : undefined;
    const username =
      isRecord(tokenValue) && typeof tokenValue["unique_name"] === "string"
        ? tokenValue["unique_name"]
        : "";

    const instance = axios.create({
      baseURL: "https://graph.microsoft.com/v1.0",
    });
    instance.defaults.headers.common["Authorization"] = `Bearer ${graphToken}`;

    let tenant = "";
    try {
      const res = await instance.get("/sites/root?$select=webUrl");
      if (res && res.data && res.data.webUrl) {
        tenant = res.data.webUrl;
      } else {
        throw new GetTenantFailedError(username);
      }
    } catch (error) {
      throw new GetTenantFailedError(username, error instanceof Error ? error : new Error());
    }
    return tenant;
  }

  public async getPackagePath(solutionConfigPath: string): Promise<string> {
    if (!(await fs.pathExists(solutionConfigPath))) {
      throw new FileNotFoundError(Constants.DeployDriverName, solutionConfigPath);
    }
    const solutionConfig = await fs.readJson(solutionConfigPath);
    const sharepointFolder = path.dirname(solutionConfigPath).replace("config", "sharepoint");
    return path.resolve(sharepointFolder, solutionConfig.paths.zippedPackage);
  }

  public async getAppID(solutionConfigPath: string): Promise<string> {
    if (!(await fs.pathExists(solutionConfigPath))) {
      throw new FileNotFoundError(Constants.DeployDriverName, solutionConfigPath);
    }
    const solutionConfig = await fs.readJson(solutionConfigPath);
    const appID = solutionConfig["solution"]["id"];
    return appID;
  }
}

export async function getSPOTenant(tokenProvider: M365TokenProvider): Promise<string> {
  return new SPFxDeployDriver().getTenant(tokenProvider);
}
