// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, M365TokenProvider, ok, Result } from "@microsoft/teamsfx-api";
import { AxiosInstance, AxiosResponse } from "axios";
import { WorkIqServiceScope } from "../common/constants";
import { WrappedAxiosClient } from "../common/wrappedAxiosClient";
import {
  createPolicyError,
  isPolicyObject,
  isStrongPolicyEtag,
  parsePolicyObject,
  PolicyDestinationResult,
  PreparedPolicyPublication,
  publishPolicyActionName,
  WorkIqPolicyPublicationRequest,
  WorkIqPolicyPublisher,
} from "../component/workiq/policyPublication";
import { isUserCancelError, UserCancelError } from "../error/common";

export type WorkIqPolicyInspection =
  { state: "found"; policy: Record<string, unknown>; etag: string } | { state: "not-found" };

export interface WorkIqPolicyDeleteRequest {
  titleId: string;
  callerId: string;
  etag: string;
  signal?: AbortSignal;
}

interface WorkIqContext {
  baseUrl: string;
  token: string;
}

interface WorkIqWrite {
  method: "PUT" | "DELETE";
  path: string;
  body: string;
  etag?: string;
  expectedStatus: 200 | 201 | 204;
  signal?: AbortSignal;
}

function isAppId(value: string): boolean {
  return (
    /^[a-f0-9]{8}(?:-[a-f0-9]{4}){3}-[a-f0-9]{12}$/i.test(value) &&
    value !== "00000000-0000-0000-0000-000000000000"
  );
}

function appBindingError(): FxError {
  return createPolicyError("WORKIQ_APP_BINDING_INVALID", "driver.workiq.error.appBinding");
}

function inputError(): FxError {
  return createPolicyError("WORKIQ_INPUT_INVALID", "driver.workiq.error.workIqInput");
}

function acknowledgementError(): FxError {
  return createPolicyError(
    "WORKIQ_ACKNOWLEDGEMENT_INVALID",
    "driver.workiq.error.outcomeUnknown",
    "system"
  );
}

function readResponseError(): FxError {
  return createPolicyError("WORKIQ_RESPONSE_INVALID", "driver.workiq.error.workIqRead", "system");
}

function agentPath(titleId: string): Result<string, FxError> {
  return segment(titleId).map((id) => `/api/policies/agents/${id}`);
}

function segment(value: string): Result<string, FxError> {
  if (
    typeof value !== "string" ||
    !value.trim() ||
    value === "." ||
    value === ".." ||
    /[\u0000-\u001f\u007f]/.test(value)
  ) {
    return err(inputError());
  }
  try {
    return ok(encodeURIComponent(value));
  } catch {
    return err(inputError());
  }
}

function responseEtag(response: AxiosResponse<unknown>): string | undefined {
  const value: unknown = response.headers.etag ?? response.headers.ETag;
  return typeof value === "string" && isStrongPolicyEtag(value) ? value : undefined;
}

export class WorkIqPolicyClient implements WorkIqPolicyPublisher {
  private readonly scopes: readonly string[];

  constructor(
    public readonly endpoint = "",
    scopes: readonly string[] = [],
    private readonly http: AxiosInstance = WrappedAxiosClient.create()
  ) {
    this.scopes = [...scopes];
  }

  public async prepare(
    request: WorkIqPolicyPublicationRequest,
    provider?: M365TokenProvider
  ): Promise<Result<PreparedPolicyPublication, FxError>> {
    if (request.signal?.aborted) return err(new UserCancelError(publishPolicyActionName));
    if (!isAppId(request.callerId)) return err(appBindingError());
    const path = agentPath(request.titleId);
    if (path.isErr()) return err(path.error);
    const policy = parsePolicyObject(request.policy.bytes);
    if (policy.isErr()) return err(policy.error);
    const appId = policy.value.appId;
    if (
      appId !== undefined &&
      (typeof appId !== "string" || appId.toLowerCase() !== request.callerId.toLowerCase())
    ) {
      return err(appBindingError());
    }
    const write: WorkIqWrite = {
      method: "PUT",
      path: path.value,
      body: JSON.stringify({ ...policy.value, appId: request.callerId }),
      expectedStatus: 201,
      signal: request.signal,
    };
    if (!this.endpoint) {
      return ok({
        publish: () =>
          Promise.resolve(
            write.signal?.aborted
              ? err({ state: "not-started", error: new UserCancelError(publishPolicyActionName) })
              : ok({ state: "mocked" })
          ),
      });
    }
    const context = await this.authenticate("write", provider, write.signal);
    if (context.isErr()) return err(context.error);
    // Temporarily paused for integration; retain the Contract preflight.
    /*
    const contracts = await this.readPrepared(
      "/api/policies/contracts",
      context.value,
      write.signal
    );
    if (contracts.isErr()) return err(contracts.error);
    */
    const inspection = await this.readPolicyPrepared(path.value, context.value, write.signal);
    if (inspection.isErr()) return err(inspection.error);
    if (inspection.value.state === "found") {
      write.etag = inspection.value.etag;
      write.expectedStatus = 200;
    }
    if (write.signal?.aborted) return err(new UserCancelError(publishPolicyActionName));
    // TODO: Validate Capability versions, writability and configuration schemas once the response contract is confirmed.
    return ok({ publish: () => this.write(write, context.value) });
  }

  public async publish(
    request: WorkIqPolicyPublicationRequest,
    provider?: M365TokenProvider
  ): Promise<PolicyDestinationResult> {
    const prepared = await this.prepare(request, provider);
    if (prepared.isErr()) return err({ state: "not-started", error: prepared.error });
    return prepared.value.publish();
  }

  public async getPolicy(
    titleId: string,
    provider: M365TokenProvider,
    signal?: AbortSignal
  ): Promise<Result<WorkIqPolicyInspection, FxError>> {
    const path = agentPath(titleId);
    if (path.isErr()) return err(path.error);
    const context = await this.authenticate("read", provider, signal);
    if (context.isErr()) return err(context.error);
    return this.readPolicyPrepared(path.value, context.value, signal);
  }

  private async readPolicyPrepared(
    path: string,
    context: WorkIqContext,
    signal?: AbortSignal
  ): Promise<Result<WorkIqPolicyInspection, FxError>> {
    const response = await this.readPrepared(path, context, signal, true);
    if (response.isErr()) return err(response.error);
    if (response.value.status === 404) return ok({ state: "not-found" });
    const etag = responseEtag(response.value);
    if (!etag || !isPolicyObject(response.value.data)) return err(readResponseError());
    return ok({ state: "found", policy: response.value.data, etag });
  }

  public async deletePolicy(
    request: WorkIqPolicyDeleteRequest,
    provider: M365TokenProvider
  ): Promise<PolicyDestinationResult> {
    if (!isAppId(request.callerId)) return err({ state: "not-started", error: appBindingError() });
    if (!isStrongPolicyEtag(request.etag))
      return err({ state: "not-started", error: inputError() });
    const path = agentPath(request.titleId);
    if (path.isErr()) return err({ state: "not-started", error: path.error });
    const write: WorkIqWrite = {
      method: "DELETE",
      path: path.value,
      body: JSON.stringify({ appId: request.callerId }),
      etag: request.etag,
      expectedStatus: 204,
      signal: request.signal,
    };
    const context = await this.authenticate("write", provider, write.signal);
    if (context.isErr()) return err({ state: "not-started", error: context.error });
    return this.write(write, context.value);
  }

  public async getCapabilityContracts(
    provider: M365TokenProvider,
    signal?: AbortSignal
  ): Promise<Result<unknown, FxError>> {
    const response = await this.read("/api/policies/contracts", provider, signal);
    if (response.isErr()) return err(response.error);
    return typeof response.value.data === "object" && response.value.data !== null
      ? ok(response.value.data)
      : err(readResponseError());
  }

  public async getCapabilityContract(
    name: string,
    version: string,
    provider: M365TokenProvider,
    signal?: AbortSignal
  ): Promise<Result<unknown, FxError>> {
    const encodedName = segment(name);
    if (encodedName.isErr()) return err(encodedName.error);
    const encodedVersion = segment(version);
    if (encodedVersion.isErr()) return err(encodedVersion.error);
    const response = await this.read(
      `/api/policies/contracts/${encodedName.value}/${encodedVersion.value}`,
      provider,
      signal
    );
    if (response.isErr()) return err(response.error);
    return isPolicyObject(response.value.data) ? ok(response.value.data) : err(readResponseError());
  }

  private async authenticate(
    permission: "read" | "write",
    provider?: M365TokenProvider,
    signal?: AbortSignal
  ): Promise<Result<WorkIqContext, FxError>> {
    if (signal?.aborted) return err(new UserCancelError(publishPolicyActionName));
    const configurationError = () =>
      createPolicyError("WORKIQ_CONFIGURATION_INVALID", "driver.workiq.error.workIqConfiguration");
    let url: URL;
    try {
      url = new URL(this.endpoint);
    } catch {
      return err(configurationError());
    }
    if (
      url.protocol !== "https:" ||
      url.username ||
      url.password ||
      url.href.includes("?") ||
      url.href.includes("#")
    ) {
      return err(configurationError());
    }
    const resources = new Set(this.scopes.map((scope) => scope.slice(0, scope.lastIndexOf("/"))));
    const allowed =
      permission === "write"
        ? ["WorkIQSettings.ReadWrite.All", ".default"]
        : ["WorkIQSettings.Read.All", "WorkIQSettings.ReadWrite.All", ".default"];
    if (
      !this.scopes.length ||
      resources.size !== 1 ||
      this.scopes.some(
        (scope) => !scope.trim() || scope !== scope.trim() || scope.lastIndexOf("/") <= 0
      ) ||
      !this.scopes.some(
        (scope) =>
          scope.lastIndexOf("/") > 0 && allowed.includes(scope.slice(scope.lastIndexOf("/") + 1))
      )
    ) {
      return err(configurationError());
    }
    const authError = () =>
      createPolicyError("WORKIQ_AUTH_REQUIRED", "driver.workiq.error.workIqAuthentication");
    if (!provider) return err(authError());
    try {
      const token = await provider.getAccessToken({ scopes: [...this.scopes] });
      if (token.isErr()) {
        return err(
          isUserCancelError(token.error)
            ? new UserCancelError(publishPolicyActionName)
            : authError()
        );
      }
      if (!token.value.trim()) return err(authError());
      if (signal?.aborted) return err(new UserCancelError(publishPolicyActionName));
      return ok({ baseUrl: url.href.replace(/\/+$/, ""), token: token.value });
    } catch (error) {
      return err(
        error instanceof Error && isUserCancelError(error)
          ? new UserCancelError(publishPolicyActionName)
          : authError()
      );
    }
  }

  private async write(
    write: WorkIqWrite,
    context: WorkIqContext
  ): Promise<PolicyDestinationResult> {
    if (write.signal?.aborted)
      return err({ state: "not-started", error: new UserCancelError(publishPolicyActionName) });
    try {
      const response = await this.http.request<unknown>({
        method: write.method,
        url: `${context.baseUrl}${write.path}`,
        data: write.body,
        headers: {
          Authorization: `Bearer ${context.token}`,
          "Accept-Encoding": "identity",
          "Content-Type": "application/json",
          ...(write.etag !== undefined ? { "If-Match": write.etag } : { "If-None-Match": "*" }),
        },
        signal: write.signal,
        timeout: 30000,
        maxRedirects: 0,
        validateStatus: () => true,
      });
      if (response.status === write.expectedStatus) {
        if (write.method === "DELETE") return ok({ state: "confirmed" });
        const etag = responseEtag(response);
        return etag
          ? ok({ state: "confirmed", etag })
          : err({ state: "unknown", error: acknowledgementError() });
      }
      if (response.status === 412) {
        return err({
          state: "conflict",
          error: createPolicyError("WORKIQ_POLICY_CONFLICT", "driver.workiq.error.workIqConflict"),
        });
      }
      if ([400, 401, 403, 404, 413, 422, 428].includes(response.status)) {
        return err({ state: "failed", error: this.rejection(response.status) });
      }
    } catch {
      return err({ state: "unknown", error: acknowledgementError() });
    }
    return err({ state: "unknown", error: acknowledgementError() });
  }

  private async read(
    path: string,
    provider: M365TokenProvider,
    signal?: AbortSignal,
    allowNotFound = false
  ): Promise<Result<AxiosResponse<unknown>, FxError>> {
    const context = await this.authenticate("read", provider, signal);
    if (context.isErr()) return err(context.error);
    return this.readPrepared(path, context.value, signal, allowNotFound);
  }

  private async readPrepared(
    path: string,
    context: WorkIqContext,
    signal?: AbortSignal,
    allowNotFound = false
  ): Promise<Result<AxiosResponse<unknown>, FxError>> {
    try {
      const response = await this.http.get<unknown>(`${context.baseUrl}${path}`, {
        headers: {
          Authorization: `Bearer ${context.token}`,
          "Accept-Encoding": "identity",
        },
        signal,
        timeout: 30000,
        maxRedirects: 0,
        validateStatus: () => true,
      });
      if (response.status === 200 || (allowNotFound && response.status === 404))
        return ok(response);
      return err(this.rejection(response.status));
    } catch {
      return err(
        signal?.aborted
          ? new UserCancelError(publishPolicyActionName)
          : createPolicyError("WORKIQ_READ_FAILED", "driver.workiq.error.workIqRead", "system")
      );
    }
  }

  private rejection(status: number): FxError {
    if (status === 429 || status >= 500 || status < 400) {
      return createPolicyError("WORKIQ_READ_FAILED", "driver.workiq.error.workIqRead", "system");
    }
    return status === 401 || status === 403
      ? createPolicyError("WORKIQ_AUTH_REQUIRED", "driver.workiq.error.workIqAuthentication")
      : createPolicyError("WORKIQ_POLICY_REJECTED", "driver.workiq.error.workIqRejected");
  }
}

export const workIqPolicyClient = new WorkIqPolicyClient(
  "https://sdf-s01-106-nam-centralus.substrate.cosmic-ppe.office.net/workiq-mcp-server",
  WorkIqServiceScope()
);
