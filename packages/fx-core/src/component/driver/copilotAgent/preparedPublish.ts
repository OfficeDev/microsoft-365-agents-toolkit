// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SystemError, UserError } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { Result } from "neverthrow";
import path from "path";
import { SovereignCloudEnvironment } from "../../../common/accountUtils";
import {
  getResourceServiceEndpoint,
  MosServiceScope,
  ResourceServiceType,
} from "../../../common/constants";
import { FeatureFlagName } from "../../../common/featureFlags";
import { getLocalizedString } from "../../../common/localizeUtils";
import { getSPFxToken } from "../../../common/tools";
import { UserCancelError } from "../../../error/common";
import { ErrorCategory } from "../../../error/types";
import { PackageValidationFailedError, PreparedPackageRequestError } from "../../m365/errors";
import { AppScope, PackageService, PreparedMosEndpoint } from "../../m365/packageService";
import { getAbsolutePath } from "../../utils/common";
import { getSPOTenant } from "../deploy/spfx/deployDriver";
import {
  createPreparedPublicationErrorData,
  isAmbiguousPublicationState,
  isPreparedCatalogUrl,
  loadPreparedPublication,
  PreparedPublicationError,
  PreparedPublicationErrorStage,
  PreparedPublicationJournal,
  PreparedPublicationState,
  PreparedWriteState,
  updatePreparedPublication,
  verifyPreparedApplicationPackage,
  verifyPreparedArtifact,
  verifyPreparedSharePointPackage,
  verifyPreparedSources,
  withPreparedPublicationIdentityLock,
} from "../deploy/spfx/utility/preparedPublication";
import { PreparedSharePointRequestContext, SPOClient } from "../deploy/spfx/utility/spoClient";
import {
  verifyPreparedMosTokenIdentity,
  verifyPreparedSubmissionIdentity,
} from "../deploy/spfx/utility/tokenIdentity";
import { DriverContext } from "../interface/commonArgs";
import { CopilotAgentPublishArgs } from "./interfaces/PublishArgs";

const sha256Pattern = /^[0-9a-f]{64}$/i;

const preparedInputNames: ReadonlyArray<keyof CopilotAgentPublishArgs> = [
  "preparedPublicationPath",
  "publicationId",
  "sourcePackageSha256",
  "sourceAppPackageSha256",
  "pluginSha256",
  "sharePointPackageSha256",
  "appPackageSha256",
];

const scopesByName = new Map<string, AppScope>(
  Object.values(AppScope).map((scope): [string, AppScope] => [scope.toLowerCase(), scope])
);

export interface PreparedPublishRequest {
  readonly args: CopilotAgentPublishArgs;
  readonly context: DriverContext;
  readonly outputEnvVarNames: Map<string, string> | undefined;
  readonly actionName: string;
}

export interface PreparedPublishResult {
  readonly output: Map<string, string>;
  readonly summaries: string[];
}

interface PreparedArgs {
  readonly appPackagePath: string;
  readonly preparedPublicationPath: string;
  readonly publicationId: string;
  readonly sourcePackageSha256: string;
  readonly sourceAppPackageSha256: string;
  readonly pluginSha256: string;
  readonly sharePointPackageSha256: string;
  readonly appPackageSha256: string;
}

interface PreparedOutputNames {
  readonly titleId: string;
  readonly appId: string;
  readonly shareLink?: string;
}

interface PreparedProgress {
  stage: PreparedPublicationErrorStage;
  journal?: PreparedPublicationJournal;
}

interface PreparedRun {
  readonly actionName: string;
  readonly context: DriverContext;
  readonly args: PreparedArgs;
  readonly outputs: PreparedOutputNames;
  readonly scope: AppScope;
  readonly progress: PreparedProgress;
}

interface MosSession {
  readonly service: PackageService;
  readonly token: string;
  readonly endpoint: PreparedMosEndpoint;
}

interface ConfirmedSubmission {
  readonly scope: string;
  readonly titleId: string;
  readonly appId: string;
}

type JournalUpdate = (current: PreparedPublicationJournal) => PreparedPublicationJournal;

interface SubmissionPhase {
  readonly state: PreparedPublicationState;
  readonly submissionState: PreparedWriteState;
}

/** Whether the YAML action selects the prepared SharePoint-first submission. */
export function hasPreparedPublishArguments(args: CopilotAgentPublishArgs): boolean {
  return preparedInputNames.some((name) => args?.[name] !== undefined);
}

/**
 * Submits the frozen application package of a SharePoint-verified prepared publication to the
 * Microsoft 365 app catalog (MOS3). Failures are thrown as sanitized, typed errors.
 */
export async function publishPreparedAgent(
  request: PreparedPublishRequest
): Promise<PreparedPublishResult> {
  const progress: PreparedProgress = { stage: "preparedPublication" };
  try {
    refuseUnsupportedCloud();
    const run: PreparedRun = {
      actionName: request.actionName,
      context: request.context,
      args: readPreparedArgs(request.args),
      outputs: readOutputNames(request.outputEnvVarNames),
      scope: readScope(request.args.scope),
      progress,
    };
    return await publishWithIdentityLock(run);
  } catch (error) {
    throw await toFailure(request, progress, error);
  }
}

function refuseUnsupportedCloud(): void {
  const cloud = process.env[FeatureFlagName.SovereignCloudEnvironment];
  if (cloud === SovereignCloudEnvironment.GCCH || cloud === SovereignCloudEnvironment.DOD) {
    throw new PreparedPublicationError(
      "UnsupportedCloud",
      "Prepared app package publication is not supported in this cloud."
    );
  }
}

function isSha256(value: unknown): value is string {
  return typeof value === "string" && sha256Pattern.test(value);
}

function readPreparedArgs(args: CopilotAgentPublishArgs): PreparedArgs {
  const {
    appPackagePath,
    preparedPublicationPath,
    publicationId,
    sourcePackageSha256,
    sourceAppPackageSha256,
    pluginSha256,
    sharePointPackageSha256,
    appPackageSha256,
  } = args;
  if (
    typeof appPackagePath === "string" &&
    appPackagePath !== "" &&
    typeof preparedPublicationPath === "string" &&
    preparedPublicationPath !== "" &&
    typeof publicationId === "string" &&
    publicationId !== "" &&
    isSha256(sourcePackageSha256) &&
    isSha256(sourceAppPackageSha256) &&
    isSha256(pluginSha256) &&
    isSha256(sharePointPackageSha256) &&
    isSha256(appPackageSha256)
  ) {
    return {
      appPackagePath,
      preparedPublicationPath,
      publicationId,
      sourcePackageSha256,
      sourceAppPackageSha256,
      pluginSha256,
      sharePointPackageSha256,
      appPackageSha256,
    };
  }
  throw new PreparedPublicationError(
    "InvalidPreparedPublishArguments",
    "Prepared publication arguments are incomplete."
  );
}

function readOutputNames(outputEnvVarNames: Map<string, string> | undefined): PreparedOutputNames {
  const titleId = outputEnvVarNames?.get("titleId");
  const appId = outputEnvVarNames?.get("appId");
  if (!titleId || !appId) {
    throw new PreparedPublicationError(
      "InvalidPreparedPublishArguments",
      "Prepared app package publication requires the titleId and appId outputs."
    );
  }
  return { titleId, appId, shareLink: outputEnvVarNames?.get("shareLink") };
}

function readScope(value: unknown): AppScope {
  if (value === undefined) {
    return AppScope.Personal;
  }
  const scope = typeof value === "string" ? scopesByName.get(value.toLowerCase()) : undefined;
  if (!scope) {
    throw new PreparedPublicationError(
      "InvalidPreparedPublishArguments",
      "The prepared publication scope is invalid."
    );
  }
  return scope;
}

function unwrap<T>(result: Result<T, Error>): T {
  if (result.isErr()) {
    throw result.error;
  }
  return result.value;
}

async function loadJournal(
  run: PreparedRun
): Promise<{ root: string; journal: PreparedPublicationJournal }> {
  return unwrap(
    await loadPreparedPublication(
      run.context.projectPath,
      run.args.publicationId,
      run.args.preparedPublicationPath
    )
  );
}

async function publishWithIdentityLock(run: PreparedRun): Promise<PreparedPublishResult> {
  const identity = await loadJournal(run);
  run.progress.stage = "sharePointVerification";
  return withPreparedPublicationIdentityLock(identity.root, identity.journal, async () => {
    // The journal may have changed while waiting for the lock, so decide on a fresh copy.
    const current = await loadJournal(run);
    run.progress.journal = current.journal;
    return publishUnderLock(run, current.root, current.journal);
  });
}

async function publishUnderLock(
  run: PreparedRun,
  root: string,
  journal: PreparedPublicationJournal
): Promise<PreparedPublishResult> {
  if (!matchesPreparedJournal(run.args, journal)) {
    throw new PreparedPublicationError(
      "PreparedPublicationNotVerified",
      "The YAML action inputs do not match the prepared publication."
    );
  }
  await verifyApplicationPackagePath(run, root, journal);
  unwrap(await verifyPreparedSources(root, journal));
  const applicationPackage = unwrap(
    await verifyPreparedArtifact(root, journal.applicationArtifact)
  );
  unwrap(verifyPreparedApplicationPackage(journal, applicationPackage));

  const confirmed = confirmedSubmission(journal);
  if (confirmed) {
    return replaySubmission(run, journal, confirmed);
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
  const tenantOrigin = await verifySharePointDeployment(run, root, journal);
  await requestApproval(run, journal, tenantOrigin);
  const session = await openMosSession(run);
  const submission = await submitApplication(run, root, session, applicationPackage);
  const shareLink = needsShareLink(run, submission)
    ? await readShareLink(run, session, submission.titleId)
    : undefined;
  return buildResult(run.outputs, submission, shareLink);
}

function matchesPreparedJournal(args: PreparedArgs, journal: PreparedPublicationJournal): boolean {
  return (
    args.sourcePackageSha256 === journal.sourceSharePoint.sha256 &&
    args.sourceAppPackageSha256 === journal.sourceApplication.sha256 &&
    args.pluginSha256 === journal.plugin.sha256 &&
    args.sharePointPackageSha256 === journal.sharePointArtifact.sha256 &&
    args.appPackageSha256 === journal.applicationArtifact.sha256 &&
    !isAmbiguousPublicationState(journal)
  );
}

async function verifyApplicationPackagePath(
  run: PreparedRun,
  root: string,
  journal: PreparedPublicationJournal
): Promise<void> {
  let requested: string;
  try {
    requested = await fs.realpath(
      getAbsolutePath(run.args.appPackagePath, run.context.projectPath)
    );
  } catch (error) {
    throw applicationPathMismatch(error);
  }
  const prepared = await fs.realpath(
    path.resolve(root, ...journal.applicationArtifact.path.split("/"))
  );
  if (requested !== prepared) {
    throw applicationPathMismatch();
  }
}

function applicationPathMismatch(cause?: unknown): PreparedPublicationError {
  return new PreparedPublicationError(
    "ApplicationArtifactPathMismatch",
    "The app package path does not match the prepared publication.",
    { cause }
  );
}

function confirmedSubmission(journal: PreparedPublicationJournal): ConfirmedSubmission | undefined {
  const submission = journal.applicationArtifact.submission;
  if (
    journal.state !== "submitted" ||
    journal.applicationArtifact.submissionState !== "confirmed" ||
    !submission?.titleId ||
    !submission.appId
  ) {
    return undefined;
  }
  return { scope: submission.scope, titleId: submission.titleId, appId: submission.appId };
}

async function replaySubmission(
  run: PreparedRun,
  journal: PreparedPublicationJournal,
  confirmed: ConfirmedSubmission
): Promise<PreparedPublishResult> {
  if (confirmed.scope !== run.scope) {
    throw new PreparedPublicationError(
      "PreparedSubmissionScopeMismatch",
      "The prepared publication was already submitted with a different scope."
    );
  }
  let shareLink: string | undefined;
  if (needsShareLink(run, confirmed)) {
    // Only the share link lookup may be retried; the confirmed submission is never repeated.
    run.progress.stage = "applicationPreflight";
    await verifyPreparedMosTokenIdentity(run.context.m365TokenProvider, journal.identity.tenantId);
    const session = await openMosSession(run);
    run.progress.stage = "applicationSubmission";
    shareLink = await readShareLink(run, session, confirmed.titleId);
  }
  return buildResult(run.outputs, confirmed, shareLink);
}

function normalizeCatalogUrl(value: string): string {
  let url: URL | undefined;
  try {
    url = new URL(value);
  } catch {
    url = undefined;
  }
  if (!url || url.protocol !== "https:" || url.username || url.password || url.search || url.hash) {
    throw new PreparedPublicationError(
      "InvalidCatalogUrl",
      "The SharePoint app catalog URL is invalid."
    );
  }
  return `${url.origin}${url.pathname.replace(/\/+$/, "")}`.toLowerCase();
}

function throwIfCancelled(run: PreparedRun): void {
  if (run.context.signal?.aborted) {
    throw new UserCancelError(run.actionName);
  }
}

/** Re-verifies, immediately before submission, that SharePoint still serves the frozen package. */
async function verifySharePointDeployment(
  run: PreparedRun,
  root: string,
  journal: PreparedPublicationJournal
): Promise<string> {
  const { context } = run;
  const sharePointPackage = unwrap(await verifyPreparedArtifact(root, journal.sharePointArtifact));
  unwrap(await verifyPreparedSharePointPackage(journal, sharePointPackage));
  const { catalogUrl, fileName } = journal.sharePointArtifact;
  if (!catalogUrl || !fileName) {
    throw new PreparedPublicationError(
      "MissingSharePointVerificationData",
      "The prepared SharePoint verification data is incomplete."
    );
  }
  throwIfCancelled(run);

  const tenant = await getSPOTenant(context.m365TokenProvider);
  const spoToken = await getSPFxToken(context.m365TokenProvider);
  if (!spoToken) {
    throw new PreparedPublicationError(
      "SharePointAuthenticationUnavailable",
      "A SharePoint access token is unavailable."
    );
  }
  const tenantOrigin = new URL(tenant).origin;
  const appCatalogSite = await SPOClient.getAppCatalogSiteForTenant(
    spoToken,
    tenantOrigin,
    context.signal
  );
  if (
    !appCatalogSite ||
    !isPreparedCatalogUrl(appCatalogSite, tenant) ||
    normalizeCatalogUrl(appCatalogSite) !== normalizeCatalogUrl(catalogUrl)
  ) {
    throw new PreparedPublicationError(
      "SharePointCatalogChanged",
      "The current SharePoint app catalog does not match the prepared publication."
    );
  }
  await verifyPreparedSubmissionIdentity(
    context.m365TokenProvider,
    journal,
    tenant,
    appCatalogSite
  );

  const requestContext: PreparedSharePointRequestContext = {
    catalogUrl: appCatalogSite,
    signal: context.signal,
  };
  const remotePackage = await SPOClient.readPreparedAppPackage(spoToken, fileName, requestContext);
  if (!remotePackage) {
    throw new PreparedPublicationError(
      "SharePointPackageNotVisible",
      "The verified SharePoint package is no longer visible."
    );
  }
  unwrap(await verifyPreparedSharePointPackage(journal, remotePackage));
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
  return tenantOrigin;
}

async function requestApproval(
  run: PreparedRun,
  journal: PreparedPublicationJournal,
  tenantOrigin: string
): Promise<void> {
  const { context } = run;
  if (!context.ui) {
    throw new PreparedPublicationError(
      "ApprovalUnavailable",
      "Prepared publication requires an interactive approval surface."
    );
  }
  run.progress.stage = "approval";
  const confirm = getLocalizedString("core.option.confirm");
  const approval = await context.ui.showMessage(
    "warn",
    getLocalizedString(
      "driver.m365.prepared.approve",
      journal.applicationArtifact.appId,
      journal.applicationArtifact.version,
      tenantOrigin,
      run.scope
    ),
    true,
    confirm
  );
  if (!approval?.isOk() || approval.value !== confirm) {
    throw new UserCancelError(run.actionName);
  }
  throwIfCancelled(run);
}

async function openMosSession(run: PreparedRun): Promise<MosSession> {
  const { context } = run;
  run.progress.stage = "applicationPreflight";
  const tokenResult = await context.m365TokenProvider.getAccessToken({
    scopes: MosServiceScope(),
  });
  if (tokenResult.isErr()) {
    throw tokenResult.error;
  }
  const service = new PackageService(
    getResourceServiceEndpoint(ResourceServiceType.MOS3),
    context.logProvider,
    context.m365TokenProvider
  );
  try {
    const endpoint = await service.resolvePreparedEndpoint(tokenResult.value, context.signal);
    return { service, token: tokenResult.value, endpoint };
  } catch (error) {
    throw new PreparedPublicationError(
      "MosEndpointUnavailable",
      "The Microsoft 365 app catalog service endpoint could not be resolved.",
      { cause: error }
    );
  }
}

function transition(from: SubmissionPhase, conflict: string, apply: JournalUpdate): JournalUpdate {
  return (current) => {
    if (
      current.state !== from.state ||
      current.applicationArtifact.submissionState !== from.submissionState
    ) {
      throw new PreparedPublicationError("ConcurrentMosSubmission", conflict);
    }
    return apply(current);
  };
}

const notStarted: SubmissionPhase = { state: "sharePointVerified", submissionState: "notStarted" };
const inFlight: SubmissionPhase = { state: "appSubmissionInFlight", submissionState: "inFlight" };

function beginSubmission(scope: AppScope): JournalUpdate {
  return transition(
    notStarted,
    "The prepared app submission state changed before the request.",
    (current) => ({
      ...current,
      state: "appSubmissionInFlight",
      applicationArtifact: {
        ...current.applicationArtifact,
        submissionState: "inFlight",
        submission: { scope },
      },
    })
  );
}

const abandonSubmission = transition(
  inFlight,
  "The prepared app submission state changed before cancellation settled.",
  (current) => {
    const applicationArtifact: PreparedPublicationJournal["applicationArtifact"] = {
      ...current.applicationArtifact,
      submissionState: "notStarted",
    };
    delete applicationArtifact.submission;
    return { ...current, state: "sharePointVerified", applicationArtifact };
  }
);

function settleFailedSubmission(rejected: boolean): JournalUpdate {
  return transition(
    inFlight,
    "The prepared app submission state changed while the request was in progress.",
    (current) => ({
      ...current,
      state: rejected ? "rejected" : "unknown",
      applicationArtifact: {
        ...current.applicationArtifact,
        submissionState: rejected ? "rejected" : "unknown",
      },
    })
  );
}

function confirmSubmission(scope: AppScope, titleId: string, appId: string): JournalUpdate {
  return transition(
    inFlight,
    "The prepared app submission state changed after the request.",
    (current) => ({
      ...current,
      state: "submitted",
      applicationArtifact: {
        ...current.applicationArtifact,
        submissionState: "confirmed",
        submission: { scope, titleId, appId },
      },
    })
  );
}

async function persist(run: PreparedRun, root: string, update: JournalUpdate): Promise<void> {
  unwrap(
    await updatePreparedPublication(
      root,
      run.args.publicationId,
      run.args.preparedPublicationPath,
      update
    )
  );
}

function isDefinitiveRejection(error: unknown): boolean {
  return (
    error instanceof PackageValidationFailedError ||
    (error instanceof PreparedPackageRequestError && error.definitive)
  );
}

/**
 * Persists the intent before the single upload attempt and the outcome right after it. The upload
 * is never retried or replayed: an outcome that cannot be proven stays `unknown` in the journal.
 */
async function submitApplication(
  run: PreparedRun,
  root: string,
  session: MosSession,
  applicationPackage: Buffer
): Promise<ConfirmedSubmission> {
  const { context, scope } = run;
  await persist(run, root, beginSubmission(scope));
  if (context.signal?.aborted) {
    await persist(run, root, abandonSubmission);
    throw new UserCancelError(run.actionName);
  }
  run.progress.stage = "applicationSubmission";
  let identifiers: [string, string];
  try {
    identifiers = await session.service.publishPreparedPackage(
      session.token,
      session.endpoint,
      applicationPackage,
      scope,
      context.signal
    );
  } catch (error) {
    const rejected = isDefinitiveRejection(error);
    await persist(run, root, settleFailedSubmission(rejected));
    throw new PreparedPublicationError(
      rejected ? "MosSubmissionRejected" : "MosSubmissionOutcomeUnknown",
      rejected
        ? "The Microsoft 365 app catalog rejected the prepared app package."
        : "The app package submission outcome is unknown; it will not be replayed.",
      { cause: error }
    );
  }
  const [titleId, appId] = identifiers;
  await persist(run, root, confirmSubmission(scope, titleId, appId));
  throwIfCancelled(run);
  return { scope, titleId, appId };
}

function needsShareLink(run: PreparedRun, submission: ConfirmedSubmission): boolean {
  return submission.scope === AppScope.Shared && run.outputs.shareLink !== undefined;
}

async function readShareLink(
  run: PreparedRun,
  session: MosSession,
  titleId: string
): Promise<string> {
  try {
    return await session.service.getPreparedShareLink(
      session.token,
      session.endpoint,
      titleId,
      run.context.signal
    );
  } catch (error) {
    throw new PreparedPublicationError(
      "MosShareLinkUnavailable",
      "The share link of the submitted app could not be retrieved.",
      { cause: error }
    );
  }
}

function buildResult(
  outputs: PreparedOutputNames,
  submission: ConfirmedSubmission,
  shareLink: string | undefined
): PreparedPublishResult {
  const output = new Map<string, string>([
    [outputs.titleId, submission.titleId],
    [outputs.appId, submission.appId],
  ]);
  if (outputs.shareLink && shareLink) {
    output.set(outputs.shareLink, shareLink);
  }
  return {
    output,
    summaries: [getLocalizedString("driver.m365.acquire.summary", submission.titleId)],
  };
}

async function toFailure(
  request: PreparedPublishRequest,
  progress: PreparedProgress,
  error: unknown
): Promise<Error> {
  if (error instanceof UserCancelError) {
    return error;
  }
  const { args, context, actionName } = request;
  const recoveryPath =
    typeof args.preparedPublicationPath === "string" ? args.preparedPublicationPath : undefined;
  let journal = progress.journal;
  if (recoveryPath !== undefined && typeof args.publicationId === "string") {
    const latest = await loadPreparedPublication(
      context.projectPath,
      args.publicationId,
      recoveryPath
    );
    if (latest.isOk()) {
      journal = latest.value.journal;
    }
  }
  if (context.signal?.aborted && journal && !isAmbiguousPublicationState(journal)) {
    return new UserCancelError(actionName);
  }
  if ((error instanceof UserError || error instanceof SystemError) && error.data === undefined) {
    error.data = createPreparedPublicationErrorData(error, journal, recoveryPath, progress.stage);
    return error;
  }
  const cause = error instanceof Error ? error : new Error("Prepared publication failed.");
  const code = cause instanceof PreparedPublicationError ? cause.code : "PreparedPublicationFailed";
  // The failure carries only allowlisted recovery data, never the raw cause or its request.
  const failure = new UserError({
    source: actionName,
    name: code,
    message: code,
    displayMessage: getLocalizedString("driver.m365.prepared.error"),
    categories: [ErrorCategory.External],
  });
  failure.data = createPreparedPublicationErrorData(cause, journal, recoveryPath, progress.stage);
  return failure;
}
