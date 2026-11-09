// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SystemError, UserError } from "@microsoft/teamsfx-api";
import { describe, expect, it } from "vitest";
import { InvalidActionInputError } from "../../../../src/error/common";
import {
  cancelPolicyPublication,
  createPolicyError,
  createPolicyPublicationFailure,
  DestinationState,
  finalizePolicyPublicationReport,
  normalizeOutcome,
  PolicyPublicationNextAction,
  PolicyPublicationOutcome,
  PolicyPublicationReport,
  withMosPolicyHttpDiagnostic,
} from "../../../../src/component/workiq/policyPublicationReport";

const report = (
  mos: DestinationState,
  workIq: DestinationState,
  target?: "mos" | "workiq" | "both"
): PolicyPublicationReport => ({
  reportVersion: 1,
  operationId: "operation",
  ...(target ? { target } : {}),
  titleId: "title",
  enforcement: "not-verified",
  outcome: "failed",
  nextAction: "review-error",
  mos: { state: mos, etag: '"mos-receipt"' },
  workIq: { state: workIq },
});

const cases: {
  mos: DestinationState;
  workIq: DestinationState;
  cancelled?: boolean;
  target?: "mos" | "workiq" | "both";
  outcome: PolicyPublicationOutcome;
  nextAction: PolicyPublicationNextAction;
}[] = [
  { mos: "confirmed", workIq: "confirmed", outcome: "published", nextAction: "none" },
  {
    mos: "confirmed",
    workIq: "mocked",
    outcome: "simulated",
    nextAction: "verify-live-publication",
  },
  { mos: "mocked", workIq: "mocked", outcome: "simulated", nextAction: "verify-live-publication" },
  { mos: "confirmed", workIq: "failed", outcome: "incomplete", nextAction: "complete-publication" },
  {
    mos: "confirmed",
    workIq: "not-started",
    outcome: "incomplete",
    nextAction: "complete-publication",
  },
  { mos: "unknown", workIq: "not-started", outcome: "unknown", nextAction: "reconcile" },
  { mos: "confirmed", workIq: "unknown", outcome: "unknown", nextAction: "reconcile" },
  { mos: "unknown", workIq: "conflict", outcome: "unknown", nextAction: "reconcile" },
  { mos: "conflict", workIq: "not-started", outcome: "conflict", nextAction: "resolve-conflict" },
  { mos: "confirmed", workIq: "conflict", outcome: "conflict", nextAction: "resolve-conflict" },
  { mos: "failed", workIq: "not-started", outcome: "failed", nextAction: "review-error" },
  {
    mos: "not-started",
    workIq: "not-started",
    cancelled: true,
    outcome: "cancelled",
    nextAction: "none",
  },
  {
    mos: "confirmed",
    workIq: "not-started",
    cancelled: true,
    outcome: "cancelled",
    nextAction: "complete-publication",
  },
  {
    mos: "confirmed",
    workIq: "unknown",
    cancelled: true,
    outcome: "cancelled",
    nextAction: "reconcile",
  },
  {
    mos: "conflict",
    workIq: "not-started",
    cancelled: true,
    outcome: "cancelled",
    nextAction: "resolve-conflict",
  },
  {
    mos: "confirmed",
    workIq: "confirmed",
    cancelled: true,
    outcome: "cancelled",
    nextAction: "none",
  },
  {
    mos: "failed",
    workIq: "not-started",
    cancelled: true,
    outcome: "cancelled",
    nextAction: "review-error",
  },
  {
    mos: "confirmed",
    workIq: "not-started",
    target: "mos",
    outcome: "published",
    nextAction: "none",
  },
  {
    mos: "not-started",
    workIq: "confirmed",
    target: "workiq",
    outcome: "published",
    nextAction: "none",
  },
  {
    mos: "not-started",
    workIq: "mocked",
    target: "workiq",
    outcome: "simulated",
    nextAction: "verify-live-publication",
  },
  {
    mos: "confirmed",
    workIq: "not-started",
    target: "mos",
    cancelled: true,
    outcome: "cancelled",
    nextAction: "none",
  },
  {
    mos: "not-started",
    workIq: "unknown",
    target: "workiq",
    cancelled: true,
    outcome: "cancelled",
    nextAction: "reconcile",
  },
];

describe("Policy publication normalization", () => {
  for (const entry of cases) {
    it(`WIQ-POL-15: ${entry.mos}/${entry.workIq}, target=${entry.target ?? "both"}, cancelled=${!!entry.cancelled} gives ${entry.outcome}/${entry.nextAction}`, () => {
      expect(normalizeOutcome(entry.mos, entry.workIq, entry.cancelled, entry.target)).toBe(
        entry.outcome
      );
      const result = finalizePolicyPublicationReport(
        report(entry.mos, entry.workIq, entry.target),
        entry.cancelled
      );
      expect(result.outcome).toBe(entry.outcome);
      expect(result.nextAction).toBe(entry.nextAction);
      expect(result.enforcement).toBe("not-verified");
    });
  }

  it("WIQ-POL-22: runtime binding or ETag conflicts are not described as MOS failures", () => {
    const result = createPolicyPublicationFailure(
      new UserError({
        source: "workiq",
        name: "WORKIQ_POLICY_CONFLICT",
        message: "private-binding",
      }),
      report("confirmed", "conflict")
    );
    expect(result.report.outcome).toBe("conflict");
    expect(result.message).toContain("Work IQ");
    expect(result.message).not.toContain("MOS reported");
    expect(result.message).not.toContain("private-binding");
  });

  it("WIQ-POL-18: failures copy receipts and never mutate or expose an original client error", () => {
    const original = new SystemError({
      source: "service",
      name: "RemoteError",
      message: "private-token https://private.example/account",
      error: new Error("private-policy-content"),
    });
    const firstReport = report("confirmed", "failed");
    const first = createPolicyPublicationFailure(original, firstReport);
    const second = createPolicyPublicationFailure(original, report("unknown", "not-started"));
    firstReport.mos.state = "failed";
    expect(first).not.toBe(original);
    expect(second).not.toBe(first);
    expect(original).not.toHaveProperty("report");
    expect(first.report.mos.state).toBe("confirmed");
    expect(first.name).toBe("POLICY_PUBLISH_INCOMPLETE");
    expect(second.name).toBe("POLICY_PUBLISH_UNKNOWN");
    expect(JSON.stringify([first, second])).not.toContain("private-");
  });

  it("WIQ-POL-18: an unknown prerequisite error gets a stable safe category", () => {
    const original = new UserError({
      source: "auth",
      name: "RemoteError",
      message: "private-account",
    });
    const result = createPolicyPublicationFailure(original, report("not-started", "not-started"));
    expect(result.name).toBe("POLICY_PREFLIGHT_FAILED");
    expect(result).toBeInstanceOf(UserError);
    expect(result.message).not.toContain("private-account");
  });

  it("WIQ-POL-18: invalid-input guidance does not request publication ETags", () => {
    const original = new InvalidActionInputError("copilotAgent/publishPolicy", ["etag"]);
    const result = createPolicyPublicationFailure(original, report("not-started", "not-started"));
    expect(result).toBeInstanceOf(InvalidActionInputError);
    expect(result.message).toContain("titleId");
    expect(result.message).toContain("policyPath");
    expect(result.message).toContain("callerId");
    expect(result.message).not.toMatch(/etag/i);
  });

  it("WIQ-POL-16: late cancellation snapshots retain ambiguous receipts and reconcile guidance", () => {
    const original = report("confirmed", "unknown");
    const result = cancelPolicyPublication(original);
    expect(result.name).toBe("UserCancel");
    expect(result.report.outcome).toBe("cancelled");
    expect(result.report.nextAction).toBe("reconcile");
    expect(original.outcome).toBe("failed");
    expect(result.report.mos.etag).toBe('"mos-receipt"');
  });

  it("WIQ-POL-18: cancellation cannot label an already-partial publication as a preflight failure", () => {
    const cause = new UserError({ source: "workiq", name: "Rejected", message: "private-binding" });
    const result = createPolicyPublicationFailure(cause, report("confirmed", "failed"), true);
    expect(result.report.outcome).toBe("cancelled");
    expect(result.name).toBe("POLICY_PUBLISH_INCOMPLETE");
    expect(result.report.nextAction).toBe("complete-publication");
  });

  it("WIQ-POL-18: an unknown error category after a rejected write is not a preflight failure", () => {
    const cause = new UserError({
      source: "service",
      name: "Rejected",
      message: "private-service",
    });
    const result = createPolicyPublicationFailure(cause, report("failed", "not-started"));
    expect(result.name).toBe("POLICY_PUBLISH_FAILED");
    expect(result.report.outcome).toBe("failed");
    expect(result.message).not.toContain("private-service");
  });

  it("WIQ-POL-45: repeated normalization copies safe HTTP diagnostics without duplication or mutating the client error", () => {
    const original = withMosPolicyHttpDiagnostic(
      createPolicyError("MosPolicyRejected", "driver.workiq.error.rejected"),
      400,
      { Error: { Code: "BadRequest", Message: "Fix the policy capabilities." } },
      [],
      "synthetic-test-token"
    );
    const first = createPolicyPublicationFailure(original, report("failed", "not-started"));
    const second = createPolicyPublicationFailure(first, first.report);

    expect(first).not.toBe(original);
    expect(second).not.toBe(first);
    expect(original).not.toHaveProperty("report");
    expect(first.message).toBe(original.message);
    expect(second.message).toBe(first.message);
    expect(second.message.match(/Service message:/g)).toHaveLength(1);
    expect(second.displayMessage).toContain("Fix the policy capabilities.");
  });

  it("WIQ-POL-45: cancellation keeps reconciliation guidance without service diagnostics", () => {
    const original = withMosPolicyHttpDiagnostic(
      createPolicyError("WorkIqMosOutcomeUnknown", "driver.workiq.error.outcomeUnknown", "system"),
      503,
      { Error: { Code: "ServiceUnavailable", Message: "private-service-details" } },
      [],
      "synthetic-test-token"
    );

    const result = createPolicyPublicationFailure(original, report("unknown", "not-started"), true);

    expect(result.report.outcome).toBe("cancelled");
    expect(result.report.nextAction).toBe("reconcile");
    expect(result.message).not.toMatch(/HTTP 503|ServiceUnavailable|private-service-details/);
  });

  it("WIQ-POL-45: client-supplied text properties cannot bypass HTTP diagnostic normalization", () => {
    const original = Object.assign(
      new UserError({
        source: "client",
        name: "MosPolicyRejected",
        message: "private-client-details",
      }),
      {
        mosHttpDiagnostic: {
          message: "private-service-details",
          displayMessage: "private-details",
        },
      }
    );

    const result = createPolicyPublicationFailure(original, report("failed", "not-started"));

    expect(result.message).toContain("MOS rejected the policy publication");
    expect(JSON.stringify(result)).not.toContain("private-");
  });
});
