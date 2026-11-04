// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result } from "@microsoft/teamsfx-api";
import { tokenize } from "comment-json";
import { createHash, randomUUID } from "crypto";
import { readFile, realpath } from "fs/promises";
import path from "path";
import { InvalidActionInputError, isUserCancelError, UserCancelError } from "../../error/common";
import {
  createPolicyError,
  createPolicyPublicationFailure,
  finalizePolicyPublicationReport,
  PolicyPublicationError,
  PolicyPublicationReport,
  publishPolicyActionName,
} from "./policyPublicationReport";
export * from "./policyPublicationReport";

export interface PublishPolicyArgs {
  titleId: string;
  policyPath: string;
  callerId: string;
}

export interface CapturedPolicy {
  bytes: Buffer;
  sha256: string;
}

export interface PolicyPublicationContext {
  projectPath: string;
  signal?: AbortSignal;
}

export interface MosPolicyPublicationRequest {
  titleId: string;
  policy: CapturedPolicy;
  signal?: AbortSignal;
}

export interface WorkIqPolicyPublicationRequest {
  titleId: string;
  callerId: string;
  policy: CapturedPolicy;
  signal?: AbortSignal;
}

export interface PolicyDestinationReceipt {
  state: "confirmed" | "mocked";
  revision?: string;
  etag?: string;
}

export interface PolicyDestinationFailure {
  state: "not-started" | "failed" | "conflict" | "unknown";
  error: FxError;
}

export type PolicyDestinationResult = Result<PolicyDestinationReceipt, PolicyDestinationFailure>;

export interface PreparedPolicyPublication {
  publish(): Promise<PolicyDestinationResult>;
}

export interface MosPolicyPublisher {
  prepare(
    request: MosPolicyPublicationRequest
  ): Promise<Result<PreparedPolicyPublication, FxError>>;
}

export interface WorkIqPolicyPublisher {
  prepare(
    request: WorkIqPolicyPublicationRequest
  ): Promise<Result<PreparedPolicyPublication, FxError>>;
}

export function isStrongPolicyEtag(value: string): boolean {
  return /^"[\x21\x23-\x7e\x80-\xff]*"$/.test(value);
}

function validateArgs(value: unknown): Result<PublishPolicyArgs, FxError> {
  if (typeof value !== "object" || value === null || Array.isArray(value)) {
    return err(new InvalidActionInputError(publishPolicyActionName, ["with"]));
  }
  if (
    !("titleId" in value) ||
    typeof value.titleId !== "string" ||
    !value.titleId.trim() ||
    !("callerId" in value) ||
    typeof value.callerId !== "string" ||
    !value.callerId.trim() ||
    !("policyPath" in value) ||
    typeof value.policyPath !== "string" ||
    !value.policyPath.trim()
  ) {
    return err(
      new InvalidActionInputError(publishPolicyActionName, ["titleId", "callerId", "policyPath"])
    );
  }
  const unexpected = Object.keys(value).filter(
    (key) => !["titleId", "callerId", "policyPath"].includes(key)
  );
  if (unexpected.length) {
    return err(new InvalidActionInputError(publishPolicyActionName, unexpected));
  }
  return ok({
    titleId: value.titleId,
    callerId: value.callerId,
    policyPath: value.policyPath,
  });
}

function isWithinProject(root: string, filePath: string): boolean {
  const relative = path.relative(root, filePath);
  return relative !== ".." && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative);
}

function hasDuplicateKeys(source: string): boolean {
  // Escape only a validation copy for the JS tokenizer; never change captured JSON bytes.
  const tokens = tokenize(source.replace(/\u2028/g, "\\u2028").replace(/\u2029/g, "\\u2029"));
  const scopes: (Set<string> | undefined)[] = [];
  for (let index = 0; index < tokens.length; index++) {
    const token = tokens[index];
    if (token.value === "{") scopes.push(new Set());
    else if (token.value === "[") scopes.push(undefined);
    else if (token.value === "}" || token.value === "]") scopes.pop();
    else if (token.type === "String" && tokens[index + 1]?.value === ":") {
      const key: unknown = JSON.parse(token.value);
      const keys = scopes[scopes.length - 1];
      if (typeof key === "string" && keys) {
        if (keys.has(key)) return true;
        keys.add(key);
      }
    }
  }
  return false;
}

export function isPolicyObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function parsePolicyObject(bytes: Buffer): Result<Record<string, unknown>, FxError> {
  try {
    const source = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    const value: unknown = JSON.parse(source);
    if (!isPolicyObject(value) || hasDuplicateKeys(source)) {
      return err(createPolicyError("InvalidWorkIqPolicy", "driver.workiq.error.invalidPolicy"));
    }
    return ok(value);
  } catch {
    return err(createPolicyError("InvalidWorkIqPolicy", "driver.workiq.error.invalidPolicy"));
  }
}

async function capturePolicy(
  policyPath: string,
  context: PolicyPublicationContext
): Promise<Result<CapturedPolicy, FxError>> {
  let bytes: Buffer;
  try {
    const root = await realpath(context.projectPath);
    const candidate = path.resolve(context.projectPath, policyPath);
    if (!isWithinProject(path.resolve(context.projectPath), candidate)) {
      return err(createPolicyError("PolicyPathError", "driver.workiq.error.outsideProject"));
    }
    const resolved = await realpath(candidate);
    if (!isWithinProject(root, resolved)) {
      return err(createPolicyError("PolicyPathError", "driver.workiq.error.outsideProject"));
    }
    bytes = await readFile(resolved, { signal: context.signal });
  } catch (error) {
    return err(
      context.signal?.aborted
        ? new UserCancelError(publishPolicyActionName)
        : createPolicyError(
            "WorkIqPolicyReadFailed",
            "driver.workiq.error.readFailed",
            "user",
            error instanceof Error ? error : undefined
          )
    );
  }
  const parsed = parsePolicyObject(bytes);
  if (parsed.isErr()) return err(parsed.error);
  return ok({ bytes, sha256: createHash("sha256").update(bytes).digest("hex") });
}

export class PolicyPublicationOperation {
  constructor(
    private readonly clients: { mos: MosPolicyPublisher; workIq: WorkIqPolicyPublisher }
  ) {}

  public async execute(
    input: unknown,
    context: PolicyPublicationContext
  ): Promise<Result<PolicyPublicationReport, PolicyPublicationError>> {
    const report: PolicyPublicationReport = {
      reportVersion: 1,
      operationId: randomUUID(),
      outcome: "failed",
      nextAction: "review-error",
      enforcement: "not-verified",
      mos: { state: "not-started" },
      workIq: { state: "not-started" },
    };
    const failure = (error: FxError): Result<PolicyPublicationReport, PolicyPublicationError> =>
      err(createPolicyPublicationFailure(error, report, context.signal?.aborted));
    if (context.signal?.aborted) {
      return failure(new UserCancelError(publishPolicyActionName));
    }
    const args = validateArgs(input);
    if (args.isErr()) return failure(args.error);
    report.titleId = args.value.titleId;
    const captured = await capturePolicy(args.value.policyPath, context);
    if (captured.isErr()) return failure(captured.error);
    const policy = captured.value;
    report.policySha256 = policy.sha256;

    let mos: Result<PreparedPolicyPublication, FxError>;
    let workIq: Result<PreparedPolicyPublication, FxError>;
    try {
      workIq = await this.clients.workIq.prepare({
        titleId: args.value.titleId,
        callerId: args.value.callerId,
        policy,
        signal: context.signal,
      });
      if (workIq.isErr()) return failure(workIq.error);
      if (context.signal?.aborted) return failure(new UserCancelError(publishPolicyActionName));
      mos = await this.clients.mos.prepare({
        titleId: args.value.titleId,
        policy,
        signal: context.signal,
      });
      if (mos.isErr()) return failure(mos.error);
    } catch (error) {
      return failure(
        error instanceof Error && isUserCancelError(error)
          ? new UserCancelError(publishPolicyActionName)
          : createPolicyError("POLICY_PREFLIGHT_FAILED", "driver.workiq.error.preflight", "system")
      );
    }
    const publications: ["mos" | "workIq", PreparedPolicyPublication][] = [
      ["mos", mos.value],
      ["workIq", workIq.value],
    ];
    for (const [destination, prepared] of publications) {
      if (context.signal?.aborted) {
        return failure(new UserCancelError(publishPolicyActionName));
      }
      let result: PolicyDestinationResult;
      try {
        result = await prepared.publish();
      } catch (error) {
        report[destination] = { state: "unknown" };
        return failure(
          error instanceof Error && isUserCancelError(error)
            ? new UserCancelError(publishPolicyActionName)
            : createPolicyError(
                "PolicyOutcomeUnknown",
                "driver.workiq.error.outcomeUnknown",
                "system"
              )
        );
      }
      if (result.isErr()) {
        report[destination] = { state: result.error.state };
        return failure(result.error.error);
      }
      report[destination] = result.value;
      if (destination === "mos" && result.value.state !== "confirmed") {
        return failure(
          createPolicyError("PolicyMosNotConfirmed", "driver.workiq.error.outcomeUnknown", "system")
        );
      }
    }
    if (context.signal?.aborted) return failure(new UserCancelError(publishPolicyActionName));
    return ok(finalizePolicyPublicationReport(report));
  }
}
