// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError, UserError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../../common/localizeUtils";
import { InvalidActionInputError, isUserCancelError, UserCancelError } from "../../error/common";

export const publishPolicyActionName = "copilotAgent/publishPolicy";

export type PolicyPublicationTarget = "mos" | "workiq" | "both";

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
  target?: PolicyPublicationTarget;
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
  cancelled = false,
  target: PolicyPublicationTarget = "both"
): PolicyPublicationOutcome {
  const states = target === "mos" ? [mos] : target === "workiq" ? [workIq] : [mos, workIq];
  if (cancelled) return "cancelled";
  if (states.includes("unknown")) return "unknown";
  if (states.includes("conflict")) return "conflict";
  if (states.every((state) => state === "confirmed")) return "published";
  if (states.every((state) => state === "confirmed" || state === "mocked")) {
    return "simulated";
  }
  if (states.includes("confirmed")) return "incomplete";
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
  const receiptOutcome = normalizeOutcome(mos.state, workIq.state, false, report.target);
  const outcome = cancelled ? "cancelled" : receiptOutcome;
  let nextAction: PolicyPublicationNextAction = "none";
  if (receiptOutcome === "unknown") nextAction = "reconcile";
  else if (receiptOutcome === "conflict") nextAction = "resolve-conflict";
  else if (receiptOutcome === "simulated") nextAction = "verify-live-publication";
  else if (receiptOutcome === "incomplete") {
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
  const receiptOutcome = normalizeOutcome(
    snapshot.mos.state,
    snapshot.workIq.state,
    false,
    snapshot.target
  );
  let safeError: FxError;
  if (isUserCancelError(error)) {
    safeError = new UserCancelError(publishPolicyActionName);
  } else if (receiptOutcome === "unknown") {
    safeError = createPolicyError(
      "POLICY_PUBLISH_UNKNOWN",
      "driver.workiq.error.outcomeUnknown",
      "system"
    );
  } else if (receiptOutcome === "conflict" && error.name === "POLICY_MOS_POLICY_EXISTS") {
    safeError = createPolicyError(
      "POLICY_MOS_POLICY_EXISTS",
      "driver.workiq.error.mosPolicyExists"
    );
  } else if (receiptOutcome === "conflict") {
    safeError = createPolicyError(
      "POLICY_REVISION_CONFLICT",
      snapshot.mos.state === "conflict"
        ? "driver.workiq.error.conflict"
        : "driver.workiq.error.workIqConflict"
    );
  } else if (receiptOutcome === "incomplete") {
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
    ]);
  } else {
    const messages: Record<string, string> = {
      InvalidWorkIqPolicy: "driver.workiq.error.invalidPolicy",
      WorkIqPolicyReadFailed: "driver.workiq.error.readFailed",
      PolicyPathError: "driver.workiq.error.outsideProject",
      MosDiscoveryError: "driver.workiq.error.serviceDiscovery",
      POLICY_AUTH_REQUIRED: "driver.workiq.error.authentication",
      POLICY_MOS_TOKEN_REJECTED: "driver.workiq.error.mosTokenRejected",
      POLICY_MOS_ACCESS_DENIED: "driver.workiq.error.mosAccessDenied",
      POLICY_MOS_ENV_INVALID: "driver.workiq.error.mosEnvironment",
      POLICY_PUBLISH_TARGET_INVALID: "driver.workiq.error.publishTarget",
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
