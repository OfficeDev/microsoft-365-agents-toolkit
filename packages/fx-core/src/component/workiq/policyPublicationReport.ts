// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError, UserError } from "@microsoft/teamsfx-api";
import { getDefaultString, getLocalizedString } from "../../common/localizeUtils";
import { maskSecret } from "../../common/stringUtils";
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

const mosHttpDiagnostic = Symbol("mosHttpDiagnostic");

interface MosHttpDiagnostic {
  message: string;
  displayMessage: string;
  publicValues: string[];
}

function mosHttpPublicValues(error: FxError): string[] {
  if (mosHttpDiagnostic in error) {
    const diagnostic = error[mosHttpDiagnostic];
    if (
      typeof diagnostic === "object" &&
      diagnostic !== null &&
      "publicValues" in diagnostic &&
      Array.isArray(diagnostic.publicValues)
    ) {
      return diagnostic.publicValues.filter(
        (value): value is string =>
          typeof value === "string" && /^[A-Za-z0-9][A-Za-z0-9_.-]{0,63}$/.test(value)
      );
    }
  }
  return [];
}

export function maskPolicyPublicationError(error: FxError): string {
  const message =
    "displayMessage" in error && typeof error.displayMessage === "string"
      ? error.displayMessage
      : error.message;
  return maskSecret(message, {
    whiteList: mosHttpPublicValues(error).flatMap((value) => [value, `${value}.`]),
  });
}

function appendMosHttpDiagnostic(error: FxError, diagnostic: MosHttpDiagnostic): FxError {
  const displayMessage =
    "displayMessage" in error && typeof error.displayMessage === "string"
      ? error.displayMessage
      : error.message;
  return Object.assign(error, {
    message: `${error.message}\n${diagnostic.message}`,
    displayMessage: `${displayMessage}\n${diagnostic.displayMessage}`,
    [mosHttpDiagnostic]: diagnostic,
  });
}

export function withMosPolicyHttpDiagnostic(
  error: FxError,
  status: number,
  data: unknown,
  requestIds: unknown[],
  accessToken: string
): FxError {
  if (!Number.isInteger(status) || status < 400 || status >= 600) return error;
  const mask = (value: string, publicValue = false) =>
    maskSecret(
      accessToken ? value.replaceAll(accessToken, "<REDACTED:secret>") : value,
      publicValue ? { whiteList: [value] } : undefined
    );
  const serviceError =
    typeof data === "object" && data !== null && !Array.isArray(data) && "Error" in data
      ? data.Error
      : undefined;
  let code: string | undefined;
  let message: string | undefined;
  if (typeof serviceError === "object" && serviceError !== null && !Array.isArray(serviceError)) {
    if (
      "Code" in serviceError &&
      typeof serviceError.Code === "string" &&
      /^[A-Za-z][A-Za-z0-9_.-]{0,63}$/.test(serviceError.Code) &&
      serviceError.Code !== accessToken
    ) {
      code = mask(serviceError.Code, true);
    }
    if (
      status < 500 &&
      code &&
      "Message" in serviceError &&
      typeof serviceError.Message === "string" &&
      !/https?:\/\/|[A-Za-z]:[\\/]|(?:^|\s)\/[\w.-]+\/|<[^>]*>|(?:^|\s)at\s+\S+.*:\d/i.test(
        serviceError.Message
      )
    ) {
      const text = serviceError.Message.replace(/\u001b\[[0-?]*[ -/]*[@-~]/g, "")
        .replace(/[\p{Cc}\p{Cf}\u2028\u2029]/gu, " ")
        .replace(/\s+/g, " ")
        .trim();
      const masked = mask(text);
      if (masked) message = masked.length > 1024 ? `${masked.slice(0, 1021)}...` : masked;
    }
  }
  const requestId = requestIds.find(
    (value): value is string =>
      typeof value === "string" &&
      value !== accessToken &&
      /^(?:[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}|[0-9a-f]{2}-[0-9a-f]{32}-[0-9a-f]{16}-[0-9a-f]{2})$/i.test(
        value
      )
  );
  const fields: [string, string | number][] = [["driver.workiq.error.mosHttpStatus", status]];
  if (code) fields.push(["driver.workiq.error.mosServiceCode", code]);
  if (message) fields.push(["driver.workiq.error.mosServiceMessage", message]);
  const publicRequestId = requestId ? mask(requestId, true) : undefined;
  if (publicRequestId) fields.push(["driver.workiq.error.mosRequestId", publicRequestId]);
  return appendMosHttpDiagnostic(error, {
    message: fields.map(([key, value]) => getDefaultString(key, value)).join(" "),
    displayMessage: fields.map(([key, value]) => getLocalizedString(key, value)).join(" "),
    publicValues: [code, publicRequestId].filter(
      (value): value is string => typeof value === "string"
    ),
  });
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
  // Only diagnostics created at the HTTP boundary survive error normalization.
  if (!cancelled && !isUserCancelError(error) && mosHttpDiagnostic in error) {
    const diagnostic = error[mosHttpDiagnostic];
    if (
      typeof diagnostic === "object" &&
      diagnostic !== null &&
      "message" in diagnostic &&
      typeof diagnostic.message === "string" &&
      "displayMessage" in diagnostic &&
      typeof diagnostic.displayMessage === "string"
    ) {
      safeError = appendMosHttpDiagnostic(safeError, {
        message: diagnostic.message,
        displayMessage: diagnostic.displayMessage,
        publicValues: mosHttpPublicValues(error),
      });
    }
  }
  return Object.assign(safeError, { report: snapshot });
}

export function cancelPolicyPublication(report: PolicyPublicationReport): PolicyPublicationError {
  return createPolicyPublicationFailure(new UserCancelError(publishPolicyActionName), report, true);
}
