// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "reflect-metadata";
import { err, M365TokenProvider, ok, Platform, UserError } from "@microsoft/teamsfx-api";
import axios from "axios";
import MockAdapter from "axios-mock-adapter";
import { createHash } from "crypto";
import { mkdtemp, readFile, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WorkIqPolicyClient, workIqPolicyClient } from "../../../../src/client/workIqPolicyClient";
import {
  getResourceServiceEndpoint,
  MosServiceScope,
  ResourceServiceType,
} from "../../../../src/common/constants";
import { DriverContext } from "../../../../src/component/driver/interface/commonArgs";
import { PublishPolicyDriver } from "../../../../src/component/driver/workiq/publishPolicy";
import { PackageService } from "../../../../src/component/m365/packageService";
import { MockedAzureAccountProvider, MockedM365Provider } from "../../../core/utils";
import { MockedLogProvider } from "../../../plugins/solution/util";

const endpoint = getResourceServiceEndpoint(ResourceServiceType.MOS3);
const discoveryUrl = `${endpoint}/config/v1/environment`;
const serviceUrl = "https://regional.titles.example/avalon";
const titleId = "title/with-special?characters";
const callerId = "11111111-2222-4333-8444-555555555555";
const mosUrl = `${serviceUrl}/builder/v1/tenants/titles/${encodeURIComponent(titleId)}/policies/workiq`;
const workIqEndpoint = "https://workiq.example";
const workIqScope = "api://workiq-test/WorkIQSettings.ReadWrite.All";
const workIqUrl = `${workIqEndpoint}/api/policies/agents/${encodeURIComponent(titleId)}`;
const source = '{\r\n  "schemaVersion": 1,\r\n  "capabilities": []\r\n}\r\n';
const policy = { schemaVersion: 1, capabilities: [] };
const args = { titleId, callerId, policyPath: "./policy.json" };
const capabilityCountMessage = "Supply between 1 and 11 capabilities.";

describe("MOS Builder policy contract", () => {
  let projectPath: string;
  let adapter: MockAdapter;
  let provider: MockedM365Provider;
  const getToken = vi.fn<M365TokenProvider["getAccessToken"]>();
  const context = (signal?: AbortSignal): DriverContext => ({
    projectPath,
    platform: Platform.CLI,
    m365TokenProvider: provider,
    azureAccountProvider: new MockedAzureAccountProvider(),
    logProvider: new MockedLogProvider(),
    ui: undefined,
    progressBar: undefined,
    signal,
  });
  const execute = (signal?: AbortSignal) =>
    new PublishPolicyDriver().execute(
      args,
      context(signal),
      new Map([["report", "POLICY_REPORT"]])
    );
  const service = () => new PackageService(endpoint, undefined, provider);

  beforeEach(async () => {
    vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "mos");
    projectPath = await mkdtemp(path.join(os.tmpdir(), "atk-mos-policy-"));
    await writeFile(path.join(projectPath, "policy.json"), source);
    adapter = new MockAdapter(axios);
    adapter.onGet(discoveryUrl).reply(200, { titlesServiceUrl: serviceUrl });
    adapter.onGet(mosUrl).reply(404);
    adapter.onPut(mosUrl).reply(201, { policy }, { ETag: '"sdf-created"' });
    adapter.onGet(workIqUrl).reply(404);
    adapter.onPut(workIqUrl).reply(201, {}, { ETag: '"workiq-created"' });
    provider = new MockedM365Provider();
    getToken.mockReset();
    getToken.mockImplementation(async (request) =>
      ok(request.scopes.includes(workIqScope) ? "workiq-token" : "sdf-token")
    );
    provider.getAccessToken = getToken;
    const workIq = new WorkIqPolicyClient(workIqEndpoint, [workIqScope]);
    vi.spyOn(workIqPolicyClient, "prepare").mockImplementation((request, tokenProvider) =>
      workIq.prepare(request, tokenProvider)
    );
  });

  afterEach(async () => {
    adapter.restore();
    vi.restoreAllMocks();
    vi.unstubAllEnvs();
    await rm(projectPath, { recursive: true, force: true });
  });

  for (const bom of [false, true]) {
    it(`WIQ-POL-39: wraps MOS JSON with BOM=${bom} without changing file bytes or hash`, async () => {
      const bytes = Buffer.from(`${bom ? "\uFEFF" : ""}${source}`);
      await writeFile(path.join(projectPath, "policy.json"), bytes);

      const result = await execute();

      expect(result.result.isOk()).toBe(true);
      expect(result.report).toMatchObject({
        target: "mos",
        outcome: "published",
        policySha256: createHash("sha256").update(bytes).digest("hex"),
        mos: { state: "confirmed", etag: '"sdf-created"' },
        workIq: { state: "not-started" },
      });
      expect(getToken).toHaveBeenCalledExactlyOnceWith({ scopes: MosServiceScope() });
      expect(adapter.history.get.map((request) => axios.getUri(request))).toEqual([
        discoveryUrl,
        mosUrl,
      ]);
      expect(adapter.history.put).toHaveLength(1);
      expect(axios.getUri(adapter.history.put[0])).toBe(mosUrl);
      expect(adapter.history.put[0].data).toBe(`{"policy":${source}}`);
      expect(JSON.parse(adapter.history.put[0].data)).toEqual({ policy });
      expect(adapter.history.put[0].headers).toHaveProperty("Accept-Encoding", "identity");
      expect(adapter.history.put[0].headers).not.toHaveProperty("If-Match");
      expect(adapter.history.put[0].headers).toHaveProperty("If-None-Match", "*");
      expect(await readFile(path.join(projectPath, "policy.json"))).toEqual(bytes);
    });
  }

  it("WIQ-POL-39: wrapping preserves exact JSON numbers instead of reserializing parsed values", async () => {
    const exact =
      '{"schemaVersion":1,"capabilities":[{"name":"Email","contractVersion":"v1","configuration":{"value":9007199254740993}}]}';
    await writeFile(path.join(projectPath, "policy.json"), exact);
    expect((await execute()).result.isOk()).toBe(true);
    expect(adapter.history.put[0].data).toBe(`{"policy":${exact}}`);
  });

  for (const replacement of [false, true]) {
    for (const policyType of [undefined, "other"]) {
      it(`WIQ-POL-46: MOS ${replacement ? "replacement" : "creation"} ignores response policyType=${policyType}`, async () => {
        const envelope = {
          policy,
          ...(policyType === undefined ? {} : { policyType }),
        };
        adapter.onGet(mosUrl).reply(replacement ? 200 : 404, envelope, { ETag: '"stored-policy"' });
        adapter
          .onPut(mosUrl)
          .reply(replacement ? 200 : 201, envelope, { ETag: '"committed-policy"' });

        const result = await execute();

        expect(result.result.isOk()).toBe(true);
        expect(result.report).toMatchObject({
          target: "mos",
          outcome: "published",
          nextAction: "none",
          mos: { state: "confirmed", etag: '"committed-policy"' },
          workIq: { state: "not-started" },
        });
        expect(getToken).toHaveBeenCalledExactlyOnceWith({ scopes: MosServiceScope() });
        expect(adapter.history.put).toHaveLength(1);
        expect(adapter.history.put[0].headers).toHaveProperty(
          replacement ? "If-Match" : "If-None-Match",
          replacement ? '"stored-policy"' : "*"
        );
        if (result.result.isOk()) {
          expect(result.result.value.get("POLICY_REPORT")).toContain('"outcome":"published"');
        }
      });
    }
  }

  for (const stored of [null, [], false, 0, "opaque"]) {
    it(`WIQ-POL-40: stored ${JSON.stringify(stored)} fails preflight and cannot select creation or replacement`, async () => {
      vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "both");
      adapter.onGet(mosUrl).reply(200, { policy: stored }, { ETag: '"existing"' });

      const result = await execute();

      expect(result.result.isErr()).toBe(true);
      expect(result.report).toMatchObject({
        target: "both",
        outcome: "failed",
        nextAction: "review-error",
        mos: { state: "not-started" },
        workIq: { state: "not-started" },
      });
      if (result.result.isErr()) {
        expect(result.result.error.name).toBe("POLICY_PREFLIGHT_FAILED");
      }
      expect(result.summaries).toEqual([]);
      expect(adapter.history.put).toHaveLength(0);
      expect(adapter.history.delete).toHaveLength(0);

      const direct = await service().publishWorkIqPolicy({
        titleId,
        policy: { bytes: Buffer.from(source), sha256: "digest" },
      });
      expect(direct.isErr()).toBe(true);
      if (direct.isErr()) expect(direct.error.state).toBe("not-started");
      expect(adapter.history.put).toHaveLength(0);
    });
  }

  for (const headers of [{}, { ETag: 'W/"weak"' }]) {
    it(`WIQ-POL-41: existing null with invalid ETag ${JSON.stringify(headers)} fails before writes`, async () => {
      adapter.onGet(mosUrl).reply(200, { policy: null }, headers);
      const result = await execute();
      expect(result.result.isErr()).toBe(true);
      expect(result.report?.mos.state).toBe("not-started");
      expect(adapter.history.put).toHaveLength(0);
    });
  }

  for (const body of [null, {}, { policy: undefined }]) {
    it(`WIQ-POL-41: malformed GET envelope ${JSON.stringify(body)} cannot select creation`, async () => {
      adapter.onGet(mosUrl).reply(200, body, { ETag: '"stored"' });
      expect((await execute()).result.isErr()).toBe(true);
      expect(adapter.history.put).toHaveLength(0);
    });
  }

  for (const { status, state } of [
    { status: 409, state: "conflict" },
    { status: 413, state: "failed" },
    { status: 503, state: "unknown" },
  ]) {
    it(`WIQ-POL-41/WIQ-POL-43: MOS PUT ${status} retains ${state} and stops Work IQ writes`, async () => {
      vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "both");
      adapter.onPut(mosUrl).reply(status, { Error: { Message: "private-service-content" } });

      const result = await execute();

      expect(result.result.isErr()).toBe(true);
      expect(result.report).toMatchObject({
        mos: { state },
        workIq: { state: "not-started" },
      });
      expect(JSON.stringify(result)).not.toContain("private-service-content");
      expect(adapter.history.put.map((request) => axios.getUri(request))).toEqual([mosUrl]);
      expect(adapter.history.delete).toHaveLength(0);
    });
  }

  for (const { status, body, headers } of [
    { status: 201, body: {}, headers: { ETag: '"created"' } },
    { status: 201, body: { policy }, headers: {} },
    { status: 201, body: { policy }, headers: { ETag: 'W/"weak"' } },
    { status: 200, body: { policy }, headers: { ETag: '"created"' } },
    { status: 202, body: { policy }, headers: { ETag: '"created"' } },
  ]) {
    it(`WIQ-POL-41: invalid create acknowledgement ${status}/${JSON.stringify(body)}/${JSON.stringify(headers)} remains unknown`, async () => {
      adapter.onPut(mosUrl).reply(status, body, headers);
      const result = await execute();
      expect(result.result.isErr()).toBe(true);
      expect(result.report?.mos.state).toBe("unknown");
      expect(adapter.history.put).toHaveLength(1);
      expect(adapter.history.delete).toHaveLength(0);
    });
  }

  it("WIQ-POL-42: standard MOS auth and wire contract remain captured after retired env and caller mutations", async () => {
    getToken.mockImplementationOnce(async () => {
      vi.stubEnv("TEAMSFX_WORKIQ_MOS_ENV", "prod");
      return ok("sdf-token");
    });
    const input = {
      titleId,
      policy: { bytes: Buffer.from(source), sha256: "digest" },
    };
    const prepared = await service().prepareWorkIqPolicy(input);
    expect(prepared.isOk()).toBe(true);
    input.policy.bytes.fill(0);
    vi.stubEnv("TEAMSFX_WORKIQ_MOS_ENV", "msit");
    if (prepared.isOk()) expect((await prepared.value.publish()).isOk()).toBe(true);
    expect(adapter.history.put[0].data).toBe(`{"policy":${source}}`);
    expect(adapter.history.put[0].headers).not.toHaveProperty("If-Match");
    expect(adapter.history.put[0].headers).toHaveProperty("If-None-Match", "*");
    expect(getToken).toHaveBeenCalledExactlyOnceWith({ scopes: MosServiceScope() });
  });

  it("WIQ-POL-39: invalid MOS JSON from a direct client fails before authentication", async () => {
    const result = await service().prepareWorkIqPolicy({
      titleId,
      policy: { bytes: Buffer.from("{"), sha256: "digest" },
    });
    expect(result.isErr()).toBe(true);
    expect(getToken).not.toHaveBeenCalled();
    expect(adapter.history.get).toHaveLength(0);
  });

  it("WIQ-POL-43: both mode keeps MOS envelope, Work IQ App ID and source digest independent", async () => {
    vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "both");
    const result = await execute();
    expect(result.result.isOk()).toBe(true);
    expect(result.report).toMatchObject({
      outcome: "published",
      mos: { state: "confirmed", etag: '"sdf-created"' },
      workIq: { state: "confirmed", etag: '"workiq-created"' },
    });
    expect(adapter.history.put.map((request) => axios.getUri(request))).toEqual([
      mosUrl,
      workIqUrl,
    ]);
    expect(JSON.parse(adapter.history.put[0].data)).toEqual({ policy });
    expect(JSON.parse(adapter.history.put[1].data)).toEqual({ ...policy, appId: callerId });
    expect(await readFile(path.join(projectPath, "policy.json"), "utf8")).toBe(source);
  });

  it("WIQ-POL-43: cancellation during MOS PUT retains unknown receipt and stops Work IQ", async () => {
    vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "both");
    const controller = new AbortController();
    adapter.onPut(mosUrl).reply(() => {
      controller.abort();
      return [201, { policy }, { ETag: '"created"' }];
    });
    const result = await execute(controller.signal);
    expect(result.result.isErr()).toBe(true);
    expect(result.report).toMatchObject({
      outcome: "cancelled",
      nextAction: "reconcile",
      mos: { state: "unknown" },
      workIq: { state: "not-started" },
    });
    expect(adapter.history.put.map((request) => axios.getUri(request))).toEqual([mosUrl]);
  });

  for (const message of [
    capabilityCountMessage,
    `${capabilityCountMessage} (Parameter 'input')`,
    "The capability name is not supported.",
  ]) {
    it(`WIQ-POL-45: structured service message reaches the native user error and log for ${message}`, async () => {
      vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "both");
      adapter.onPut(mosUrl).reply(400, {
        Error: { Code: "BadRequest", Message: message },
        privateDetails: "private-policy-details",
      });
      const ctx = context();
      const logError = vi.spyOn(ctx.logProvider, "error");

      const result = await new PublishPolicyDriver().execute(
        args,
        ctx,
        new Map([["report", "POLICY_REPORT"]])
      );

      expect(result.result.isErr()).toBe(true);
      if (result.result.isErr()) {
        expect(result.result.error).toBeInstanceOf(UserError);
        expect(result.result.error.name).toBe("MosPolicyRejected");
        expect(result.result.error.message).toContain(message);
        expect(result.result.error.message).toContain("HTTP 400");
        expect(result.result.error.message).toContain("BadRequest");
        expect(result.result.error).toHaveProperty(
          "displayMessage",
          expect.stringContaining(message)
        );
        expect(result.result.error).toHaveProperty("report", result.report);
      }
      expect(logError).toHaveBeenCalledWith(expect.stringContaining(message));
      expect(logError).toHaveBeenCalledWith(
        expect.stringContaining("Service error code: BadRequest")
      );
      expect(result.report).toMatchObject({
        target: "both",
        outcome: "failed",
        nextAction: "review-error",
        mos: { state: "failed" },
        workIq: { state: "not-started" },
      });
      expect(result.summaries).toEqual([]);
      expect(adapter.history.put.map((request) => axios.getUri(request))).toEqual([mosUrl]);
      expect(adapter.history.delete).toHaveLength(0);
      expect(JSON.stringify(result)).not.toMatch(/private-policy-details|sdf-token|workiq-token/);
    });
  }

  for (const body of [
    null,
    { Error: "private-service-details" },
    { Error: [] },
    { Error: { Code: "BadRequest", Message: 11 } },
    { Error: { Code: 400, Message: "private-service-details" } },
    { Error: { Message: "private-service-details" } },
    { Error: { Code: "BadRequest\nforged-code", Message: "private-service-details" } },
    { Error: { Code: "BadRequest", Message: "https://private.example/private-service-details" } },
    { Error: { Code: "BadRequest", Message: "C:\\private-service-details\\policy.json" } },
    { Error: { Code: "BadRequest", Message: "/internal/private-service-details/policy.json" } },
    { Error: { Code: "BadRequest", Message: "<html>private-service-details</html>" } },
    { Error: { Code: "BadRequest", Message: "at Service.validate (private-service-details:1:2)" } },
  ]) {
    it(`WIQ-POL-45: unknown or malformed MOS rejection ${JSON.stringify(body)} retains private-safe generic guidance`, async () => {
      vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "both");
      adapter.onPut(mosUrl).reply(400, body);

      const result = await execute();

      expect(result.result.isErr()).toBe(true);
      if (result.result.isErr()) {
        expect(result.result.error.name).toBe("MosPolicyRejected");
        expect(result.result.error.message).toContain("MOS rejected the policy publication");
        expect(result.result.error.message).toContain("HTTP 400");
        expect(result.result.error.message).not.toContain("between 1 and 11 capabilities");
      }
      expect(result.report).toMatchObject({
        mos: { state: "failed" },
        workIq: { state: "not-started" },
      });
      expect(adapter.history.put.map((request) => axios.getUri(request))).toEqual([mosUrl]);
      expect(JSON.stringify(result)).not.toContain("private-service-details");
    });
  }

  for (const { status, name, state } of [
    { status: 201, name: "POLICY_PUBLISH_UNKNOWN", state: "unknown" },
    { status: 400, name: "MosPolicyRejected", state: "failed" },
    { status: 401, name: "POLICY_MOS_TOKEN_REJECTED", state: "failed" },
    { status: 403, name: "POLICY_MOS_ACCESS_DENIED", state: "failed" },
    { status: 409, name: "POLICY_REVISION_CONFLICT", state: "conflict" },
    { status: 412, name: "POLICY_REVISION_CONFLICT", state: "conflict" },
    { status: 422, name: "POLICY_PUBLISH_UNKNOWN", state: "unknown" },
    { status: 429, name: "POLICY_PUBLISH_UNKNOWN", state: "unknown" },
    { status: 500, name: "POLICY_PUBLISH_UNKNOWN", state: "unknown" },
    { status: 503, name: "POLICY_PUBLISH_UNKNOWN", state: "unknown" },
  ]) {
    it(`WIQ-POL-45: service text with MOS HTTP ${status} retains the existing ${state} category`, async () => {
      vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "both");
      adapter.onPut(mosUrl).reply(status, {
        Error: { Code: "BadRequest", Message: capabilityCountMessage },
      });

      const result = await execute();

      expect(result.result.isErr()).toBe(true);
      if (result.result.isErr()) {
        expect(result.result.error.name).toBe(name);
        if (status >= 400 && status < 500) {
          expect(result.result.error.message).toContain(capabilityCountMessage);
        } else {
          expect(result.result.error.message).not.toContain(capabilityCountMessage);
        }
        if (status >= 400) expect(result.result.error.message).toContain(`HTTP ${status}`);
      }
      expect(result.report).toMatchObject({
        mos: { state },
        workIq: { state: "not-started" },
      });
      expect(adapter.history.put.map((request) => axios.getUri(request))).toEqual([mosUrl]);
    });
  }

  it("WIQ-POL-45: 4xx message masking removes acquired tokens, env secrets and terminal controls before logging", async () => {
    vi.stubEnv("SECRET_MOS_DIAGNOSTIC_TEST", "unit-test-sensitive-value");
    adapter.onPut(mosUrl).reply(400, {
      Error: {
        Code: "ValidationError",
        Message:
          "\u001b[31mCheck sdf-token\r\nunit-test-sensitive-value\u0000and capability.\u001b[0m",
      },
    });
    const ctx = context();
    const logError = vi.spyOn(ctx.logProvider, "error");

    const result = await new PublishPolicyDriver().execute(args, ctx);

    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) {
      expect(result.result.error.message).toContain("Service message: Check");
      expect(result.result.error.message).toContain("and capability.");
      expect(result.result.error.message).not.toMatch(
        /sdf-token|unit-test-sensitive-value|[\u0000\u001b\r]/
      );
    }
    expect(JSON.stringify(logError.mock.calls)).not.toMatch(/sdf-token|unit-test-sensitive-value/);
  });

  it("WIQ-POL-45: acquired tokens are also masked if echoed inside a service error code", async () => {
    adapter.onPut(mosUrl).reply(500, {
      Error: { Code: "InternalError_sdf-token", Message: "private-service-details" },
    });

    const result = await execute();

    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) {
      expect(result.result.error.message).toContain("HTTP 500");
      expect(result.result.error.message).not.toContain("sdf-token");
    }
    expect(JSON.stringify(result)).not.toContain("private-service-details");
  });

  it("WIQ-POL-45: public identifier whitelisting never bypasses explicit registered-secret masking", async () => {
    const requestId = "11111111-2222-4333-8444-555555555555";
    vi.stubEnv("SECRET_MOS_TEST_CODE", "BadRequest");
    vi.stubEnv("SECRET_MOS_TEST_REQUEST_ID", requestId);
    adapter
      .onPut(mosUrl)
      .reply(
        400,
        { Error: { Code: "BadRequest", Message: "Check the capabilities array." } },
        { "request-id": requestId }
      );
    const ctx = context();
    const logError = vi.spyOn(ctx.logProvider, "error");

    const result = await new PublishPolicyDriver().execute(args, ctx);

    expect(result.result.isErr()).toBe(true);
    expect(JSON.stringify(result)).not.toContain("BadRequest");
    expect(JSON.stringify(result)).not.toContain(requestId);
    expect(JSON.stringify(logError.mock.calls)).not.toContain("BadRequest");
    expect(JSON.stringify(logError.mock.calls)).not.toContain(requestId);
  });

  it("WIQ-POL-45: service messages are bounded to 1024 characters after secret masking", async () => {
    adapter.onPut(mosUrl).reply(400, {
      Error: { Code: "BadRequest", Message: `${"a".repeat(1500)} sdf-token` },
    });

    const result = await execute();

    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) {
      const diagnostic = result.result.error.message.split("Service message: ")[1];
      expect(diagnostic).toBe(`${"a".repeat(1021)}...`);
      expect(diagnostic).toHaveLength(1024);
      expect(result.result.error.message).not.toContain("sdf-token");
    }
  });

  for (const method of ["get", "put"]) {
    for (const [header, requestId] of [
      ["x-ms-request-id", "11111111-2222-4333-8444-555555555555"],
      ["request-id", "11111111-2222-4333-8444-555555555555"],
      ["traceresponse", "00-da08223b2900b0861c020b0aa0a3bdfb-b1fcc7dc3b63e9eb-01"],
    ]) {
      it(`WIQ-POL-45: MOS ${method} HTTP 503 retains ${header} for support without exposing internal text`, async () => {
        vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "both");
        const response = {
          Error: { Code: "ServiceUnavailable", Message: "private-service-internal-error" },
        };
        if (method === "get") adapter.onGet(mosUrl).reply(503, response, { [header]: requestId });
        else adapter.onPut(mosUrl).reply(503, response, { [header]: requestId });

        const ctx = context();
        const logError = vi.spyOn(ctx.logProvider, "error");
        const result = await new PublishPolicyDriver().execute(args, ctx);

        expect(result.result.isErr()).toBe(true);
        if (result.result.isErr()) {
          expect(result.result.error.message).toContain("HTTP 503");
          expect(result.result.error.message).toContain("ServiceUnavailable");
          expect(result.result.error.message).toContain(`Request ID: ${requestId}`);
          expect(result.result.error.message).not.toContain("private-service-internal-error");
        }
        expect(logError).toHaveBeenCalledWith(
          expect.stringContaining("Service error code: ServiceUnavailable")
        );
        expect(logError).toHaveBeenCalledWith(expect.stringContaining(`Request ID: ${requestId}`));
        expect(result.report).toMatchObject({
          mos: { state: method === "get" ? "not-started" : "unknown" },
          workIq: { state: "not-started" },
          nextAction: method === "get" ? "review-error" : "reconcile",
        });
        expect(adapter.history.put).toHaveLength(method === "get" ? 0 : 1);
      });
    }
  }

  it("WIQ-POL-45: unsafe request identifiers are ignored instead of being displayed", async () => {
    adapter
      .onPut(mosUrl)
      .reply(
        500,
        { Error: { Code: "InternalError", Message: "private-service-details" } },
        { "x-ms-request-id": "private-header-details\nforged-log" }
      );

    const result = await execute();

    expect(result.result.isErr()).toBe(true);
    expect(JSON.stringify(result)).not.toMatch(/private-service-details|private-header-details/);
  });

  it("WIQ-POL-45: GET 400 shows validation guidance without starting either destination write", async () => {
    vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "both");
    adapter.onGet(mosUrl).reply(400, {
      Error: { Code: "InvalidTitle", Message: "Check the supplied Title ID." },
    });

    const result = await execute();

    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) {
      expect(result.result.error.message).toContain("HTTP 400");
      expect(result.result.error.message).toContain("Check the supplied Title ID.");
    }
    expect(result.report).toMatchObject({
      mos: { state: "not-started" },
      workIq: { state: "not-started" },
    });
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-45: a valid secondary request ID is used when the preferred header is malformed", async () => {
    const requestId = "11111111-2222-4333-8444-555555555555";
    adapter
      .onPut(mosUrl)
      .reply(
        500,
        { Error: { Code: "InternalError", Message: "private-service-details" } },
        { "x-ms-request-id": "invalid", "request-id": requestId }
      );

    const result = await execute();

    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) {
      expect(result.result.error.message).toContain(`Request ID: ${requestId}`);
    }
    expect(JSON.stringify(result)).not.toContain("private-service-details");
  });

  it("WIQ-POL-44: provider failure remains token acquisition failure without HTTP calls", async () => {
    getToken.mockResolvedValue(
      err(
        new UserError({ source: "provider", name: "Denied", message: "private-provider-details" })
      )
    );

    const result = await execute();

    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) {
      expect(result.result.error.name).toBe("POLICY_AUTH_REQUIRED");
      expect(result.result.error.message).toContain("Unable to acquire the delegated MOS token");
    }
    expect(result.report?.mos.state).toBe("not-started");
    expect(JSON.stringify(result)).not.toContain("private-provider-details");
    expect(adapter.history.get).toHaveLength(0);
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-44: discovery rejection is distinct from token acquisition and policy access errors", async () => {
    adapter.onGet(discoveryUrl).reply(403, { detail: "private-discovery-details" });

    const result = await execute();

    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) {
      expect(result.result.error.name).toBe("MosDiscoveryError");
      expect(result.result.error.message).toContain("discover");
      expect(result.result.error.message).not.toContain("Unable to acquire");
    }
    expect(result.report?.mos.state).toBe("not-started");
    expect(JSON.stringify(result)).not.toContain("private-discovery-details");
    expect(getToken).toHaveBeenCalledOnce();
    expect(adapter.history.get.map((request) => axios.getUri(request))).toEqual([discoveryUrl]);
    expect(adapter.history.put).toHaveLength(0);
  });

  for (const method of ["get", "put"]) {
    for (const { status, name, guidance } of [
      { status: 401, name: "POLICY_MOS_TOKEN_REJECTED", guidance: "rejected the acquired token" },
      { status: 403, name: "POLICY_MOS_ACCESS_DENIED", guidance: "denied access" },
    ]) {
      it(`WIQ-POL-44: MOS policy ${method} HTTP ${status} is not a token acquisition failure`, async () => {
        vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "both");
        if (method === "get")
          adapter.onGet(mosUrl).reply(status, { Error: { Message: "private-api-details" } });
        else adapter.onPut(mosUrl).reply(status, { Error: { Message: "private-api-details" } });

        const result = await execute();

        expect(result.result.isErr()).toBe(true);
        if (result.result.isErr()) {
          expect(result.result.error.name).toBe(name);
          expect(result.result.error.message).toContain(guidance);
          expect(result.result.error.message).toContain(`HTTP ${status}`);
          expect(result.result.error).toHaveProperty(
            "displayMessage",
            expect.stringContaining(guidance)
          );
          expect(result.result.error.message).not.toContain("Unable to acquire");
          expect(result.result.error).toHaveProperty("report", result.report);
        }
        expect(result.report).toMatchObject({
          mos: { state: method === "get" ? "not-started" : "failed" },
          workIq: { state: "not-started" },
        });
        expect(getToken.mock.calls.map(([request]) => request.scopes)).toEqual([
          [workIqScope],
          MosServiceScope(),
        ]);
        expect(adapter.history.get.map((request) => axios.getUri(request))).toEqual([
          workIqUrl,
          discoveryUrl,
          mosUrl,
        ]);
        expect(adapter.history.put.map((request) => axios.getUri(request))).toEqual(
          method === "get" ? [] : [mosUrl]
        );
        expect(JSON.stringify(result)).not.toMatch(
          /private-api-details|sdf-token|workiq-token|regional\.titles\.example/
        );
      });
    }
  }
});
