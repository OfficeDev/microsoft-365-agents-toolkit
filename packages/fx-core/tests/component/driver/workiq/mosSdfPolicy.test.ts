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
import { DriverContext } from "../../../../src/component/driver/interface/commonArgs";
import { PublishPolicyDriver } from "../../../../src/component/driver/workiq/publishPolicy";
import { PackageService } from "../../../../src/component/m365/packageService";
import { MockedAzureAccountProvider, MockedM365Provider } from "../../../core/utils";
import { MockedLogProvider } from "../../../plugins/solution/util";

const endpoint = "https://titles.sdf.mos.microsoft.com";
const discoveryUrl = `${endpoint}/config/v1/environment`;
const serviceUrl = "https://regional.titles.example/avalon";
const titleId = "sdf-title/with-special?characters";
const callerId = "11111111-2222-4333-8444-555555555555";
const mosUrl = `${serviceUrl}/builder/v1/tenants/titles/${encodeURIComponent(titleId)}/policies/workiq`;
const workIqEndpoint = "https://workiq.example";
const workIqScope = "api://workiq-test/WorkIQSettings.ReadWrite.All";
const workIqUrl = `${workIqEndpoint}/api/policies/agents/${encodeURIComponent(titleId)}`;
const source = '{\r\n  "schemaVersion": 1,\r\n  "capabilities": []\r\n}\r\n';
const policy = { schemaVersion: 1, capabilities: [] };
const args = { titleId, callerId, policyPath: "./policy.json" };

describe("MOS SDF Builder policy contract", () => {
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
    vi.stubEnv("TEAMSFX_WORKIQ_MOS_ENV", "sdf");
    vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "mos");
    projectPath = await mkdtemp(path.join(os.tmpdir(), "atk-mos-sdf-policy-"));
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
    it(`WIQ-POL-39: wraps SDF JSON with BOM=${bom} without changing file bytes or hash`, async () => {
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
      expect(getToken).toHaveBeenCalledExactlyOnceWith({ scopes: [`${endpoint}/.default`] });
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
      expect(adapter.history.put[0].headers).not.toHaveProperty("If-None-Match");
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

  for (const stored of [null, {}, [], false, 0, "opaque"]) {
    it(`WIQ-POL-40: stored ${JSON.stringify(stored)} is an existing-policy conflict, not a replacement`, async () => {
      vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "both");
      adapter.onGet(mosUrl).reply(200, { policy: stored }, { ETag: '"existing"' });

      const result = await execute();

      expect(result.result.isErr()).toBe(true);
      expect(result.report).toMatchObject({
        target: "both",
        outcome: "conflict",
        nextAction: "resolve-conflict",
        mos: { state: "conflict" },
        workIq: { state: "not-started" },
      });
      if (result.result.isErr()) {
        expect(result.result.error.name).toBe("POLICY_MOS_POLICY_EXISTS");
        expect(result.result.error.message).toContain("already exists");
      }
      expect(result.summaries).toEqual([]);
      expect(adapter.history.put).toHaveLength(0);
      expect(adapter.history.delete).toHaveLength(0);

      const direct = await service().publishWorkIqPolicy({
        titleId,
        policy: { bytes: Buffer.from(source), sha256: "digest" },
      });
      expect(direct.isErr()).toBe(true);
      if (direct.isErr()) expect(direct.error.state).toBe("conflict");
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
    it(`WIQ-POL-41/WIQ-POL-43: SDF PUT ${status} retains ${state} and stops Work IQ writes`, async () => {
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

  it("WIQ-POL-42: SDF auth and wire contract remain captured after env and caller mutations", async () => {
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
    expect(adapter.history.put[0].headers).not.toHaveProperty("If-None-Match");
    expect(getToken).toHaveBeenCalledExactlyOnceWith({ scopes: [`${endpoint}/.default`] });
  });

  it("WIQ-POL-39: invalid SDF JSON from a direct client fails before authentication", async () => {
    const result = await service().prepareWorkIqPolicy({
      titleId,
      policy: { bytes: Buffer.from("{"), sha256: "digest" },
    });
    expect(result.isErr()).toBe(true);
    expect(getToken).not.toHaveBeenCalled();
    expect(adapter.history.get).toHaveLength(0);
  });

  it("WIQ-POL-43: both mode keeps SDF envelope, Work IQ App ID and source digest independent", async () => {
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

  it("WIQ-POL-43: cancellation during SDF PUT retains unknown receipt and stops Work IQ", async () => {
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

  for (const environment of ["sdf", "prod"]) {
    for (const method of ["get", "put"]) {
      for (const { status, name, guidance } of [
        { status: 401, name: "POLICY_MOS_TOKEN_REJECTED", guidance: "rejected the acquired token" },
        { status: 403, name: "POLICY_MOS_ACCESS_DENIED", guidance: "denied access" },
      ]) {
        it(`WIQ-POL-44: ${environment} policy ${method} HTTP ${status} is not a token acquisition failure`, async () => {
          vi.stubEnv("TEAMSFX_WORKIQ_MOS_ENV", environment);
          vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "both");
          const entry = environment === "sdf" ? endpoint : "https://titles.prod.mos.microsoft.com";
          const discovery = `${entry}/config/v1/environment`;
          adapter.onGet(discovery).reply(200, { titlesServiceUrl: serviceUrl });
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
            [`${entry}/.default`],
          ]);
          expect(adapter.history.get.map((request) => axios.getUri(request))).toEqual([
            workIqUrl,
            discovery,
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
  }
});
