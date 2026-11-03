// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError, UserError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../../common/localizeUtils";
import { InvalidActionInputError, isUserCancelError, UserCancelError } from "../../error/common";

export const publishPolicyActionName = "workiq/publishPolicy";

export type DestinationState =
  "not-started" | "confirmed" | "mocked" | "failed" | "conflict" | "unknown";

export type PolicyPublicationOutcome =
  | "published"
  | "unchanged"
  | "simulated"
  | "incomplete"
  | "failed"
  | "unknown"
  | "conflict"
  | "cancelled";

export type PolicyPublicationNextAction =
  | "none"
  | "reconcile"
  | "resolve-conflict"
  | "complete-publication"
  | "review-error"
  | "verify-live-publication";

export interface PolicyDestinationReport {
  state: DestinationState;
  revision?: string;
  etag?: string;
}

export interface PolicyPublicationReport {
  reportVersion: 1;
  operationId: string;
  titleId?: string;
  policySha256?: string;
  outcome: PolicyPublicationOutcome;
  nextAction: PolicyPublicationNextAction;
  enforcement: "not-verified";
  mos: PolicyDestinationReport;
  workIq: PolicyDestinationReport;
}

export interface PolicyPublicationError extends FxError {
  report: PolicyPublicationReport;
  displayMessage?: string;
}

export function createPolicyError(
  name: string,
  messageKey: string,
  kind: "user" | "system" = "user",
  cause?: Error
): FxError {
  const options = {
    source: publishPolicyActionName,
    name,
    message: getDefaultString(messageKey),
    displayMessage: getLocalizedString(messageKey),
    error: cause,
  };
  return kind === "user" ? new UserError(options) : new SystemError(options);
}

export function normalizeOutcome(
  mos: DestinationState,
  workIq: DestinationState,
  cancelled = false
): PolicyPublicationOutcome {
  if (cancelled) return "cancelled";
  if (mos === "unknown" || workIq === "unknown") return "unknown";
  if (mos === "conflict" || workIq === "conflict") return "conflict";
  if (mos === "confirmed" && workIq === "confirmed") return "published";
  if (
    (mos === "confirmed" || mos === "mocked") &&
    (workIq === "confirmed" || workIq === "mocked")
  ) {
    return "simulated";
  }
  if (mos === "confirmed" || workIq === "confirmed") return "incomplete";
  return "failed";
}

function snapshotDestination(receipt: PolicyDestinationReport): PolicyDestinationReport {
  return {
    state: receipt.state,
    ...(receipt.etag !== undefined ? { etag: receipt.etag } : {}),
    ...(receipt.revision !== undefined ? { revision: receipt.revision } : {}),
  };
}

export function finalizePolicyPublicationReport(
  report: PolicyPublicationReport,
  cancelled = false
): PolicyPublicationReport {
  const mos = snapshotDestination(report.mos);
  const workIq = snapshotDestination(report.workIq);
  const receiptOutcome = normalizeOutcome(mos.state, workIq.state);
  const outcome = cancelled ? "cancelled" : receiptOutcome;
  const partiallyConfirmed = (mos.state === "confirmed") !== (workIq.state === "confirmed");
  let nextAction: PolicyPublicationNextAction = "none";
  if (mos.state === "unknown" || workIq.state === "unknown") nextAction = "reconcile";
  else if (mos.state === "conflict" || workIq.state === "conflict") nextAction = "resolve-conflict";
  else if (receiptOutcome === "simulated") nextAction = "verify-live-publication";
  else if (partiallyConfirmed) {
    nextAction = "complete-publication";
  } else if (
    receiptOutcome === "failed" &&
    (!cancelled || mos.state !== "not-started" || workIq.state !== "not-started")
  )
    nextAction = "review-error";
  return { ...report, mos, workIq, outcome, nextAction };
}

export function createPolicyPublicationFailure(
  error: FxError,
  report: PolicyPublicationReport,
  cancelled = false
): PolicyPublicationError {
  const snapshot = finalizePolicyPublicationReport(report, cancelled || isUserCancelError(error));
  let safeError: FxError;
  if (isUserCancelError(error)) {
    safeError = new UserCancelError(publishPolicyActionName);
  } else if (snapshot.mos.state === "unknown" || snapshot.workIq.state === "unknown") {
    safeError = createPolicyError(
      "POLICY_PUBLISH_UNKNOWN",
      "driver.workiq.error.outcomeUnknown",
      "system"
    );
  } else if (snapshot.mos.state === "conflict" || snapshot.workIq.state === "conflict") {
    safeError = createPolicyError(
      "POLICY_REVISION_CONFLICT",
      snapshot.mos.state === "conflict"
        ? "driver.workiq.error.conflict"
        : "driver.workiq.error.workIqConflict"
    );
  } else if (normalizeOutcome(snapshot.mos.state, snapshot.workIq.state) === "incomplete") {
    safeError = createPolicyError(
      "POLICY_PUBLISH_INCOMPLETE",
      "driver.workiq.error.incomplete",
      error instanceof UserError ? "user" : "system"
    );
  } else if (error instanceof InvalidActionInputError) {
    safeError = new InvalidActionInputError(publishPolicyActionName, [
      "titleId",
      "policyPath",
      "callerId",
      "etag",
      "workIqEtag",
    ]);
  } else {
    const messages: Record<string, string> = {
      InvalidWorkIqPolicy: "driver.workiq.error.invalidPolicy",
      WorkIqPolicyReadFailed: "driver.workiq.error.readFailed",
      PolicyPathError: "driver.workiq.error.outsideProject",
      MosDiscoveryError: "driver.workiq.error.serviceDiscovery",
      POLICY_AUTH_REQUIRED: "driver.workiq.error.authentication",
      MosPolicyRejected: "driver.workiq.error.rejected",
      WORKIQ_CONFIGURATION_INVALID: "driver.workiq.error.workIqConfiguration",
      WORKIQ_APP_BINDING_INVALID: "driver.workiq.error.appBinding",
      WORKIQ_INPUT_INVALID: "driver.workiq.error.workIqInput",
      WORKIQ_AUTH_REQUIRED: "driver.workiq.error.workIqAuthentication",
    };
    const known = Object.hasOwn(messages, error.name);
    const beforeWrites =
      snapshot.mos.state === "not-started" && snapshot.workIq.state === "not-started";
    safeError = createPolicyError(
      known ? error.name : beforeWrites ? "POLICY_PREFLIGHT_FAILED" : "POLICY_PUBLISH_FAILED",
      known
        ? messages[error.name]
        : beforeWrites
          ? "driver.workiq.error.preflight"
          : "driver.workiq.error.failed",
      error instanceof UserError ? "user" : "system"
    );
  }
  return Object.assign(safeError, { report: snapshot });
}

export function cancelPolicyPublication(report: PolicyPublicationReport): PolicyPublicationError {
  return createPolicyPublicationFailure(new UserCancelError(publishPolicyActionName), report, true);
}
