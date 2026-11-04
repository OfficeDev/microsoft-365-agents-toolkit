// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, ok, UserError } from "@microsoft/teamsfx-api";
import axios from "axios";
import MockAdapter from "axios-mock-adapter";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WorkIqPolicyClient, workIqPolicyClient } from "../../src/client/workIqPolicyClient";
import { MockedM365Provider } from "../core/utils";

const endpoint = "https://workiq.example/environment";
const scope = "api://workiq-test/WorkIQSettings.ReadWrite.All";
const callerId = "11111111-2222-4333-8444-555555555555";
const titleId = "P_title/with?characters";
const url = `${endpoint}/api/policies/agents/${encodeURIComponent(titleId)}`;
const bytes = Buffer.from('{\r\n "schemaVersion": 1, "capabilities": []\r\n}\r\n');
const request = () => ({
  titleId,
  callerId,
  policy: { bytes: Buffer.from(bytes), sha256: "captured-hash" },
});

describe("Work IQ Agent Policy API", () => {
  let adapter: MockAdapter;
  let client: WorkIqPolicyClient;
  let provider: MockedM365Provider;

  beforeEach(() => {
    const http = axios.create();
    adapter = new MockAdapter(http);
    adapter.onGet(`${endpoint}/api/policies/contracts`).reply(200, {});
    adapter.onGet(url).reply(404);
    client = new WorkIqPolicyClient(endpoint, [scope], http);
    provider = new MockedM365Provider();
    vi.spyOn(provider, "getAccessToken").mockResolvedValue(ok("workiq-only-token"));
  });

  afterEach(() => {
    adapter.restore();
    vi.restoreAllMocks();
  });

  it("WIQ-POL-24: the exported PPE endpoint preserves its service prefix", () => {
    expect(workIqPolicyClient.endpoint).toBe(
      "https://sdf-s01-106-nam-centralus.substrate.cosmic-ppe.office.net/workiq-mcp-server"
    );
  });

  it("WIQ-POL-26: shared MOS default scopes are configured but a real token provider is still required", async () => {
    const result = await workIqPolicyClient.prepare(request());
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("WORKIQ_AUTH_REQUIRED");
    expect(provider.getAccessToken).not.toHaveBeenCalled();
  });

  it("WIQ-POL-24: configured PPE requests retain the /workiq-mcp-server API prefix", async () => {
    const http = axios.create();
    const ppeAdapter = new MockAdapter(http);
    const ppe = new WorkIqPolicyClient(workIqPolicyClient.endpoint, [scope], http);
    const expectedUrl =
      "https://sdf-s01-106-nam-centralus.substrate.cosmic-ppe.office.net/workiq-mcp-server/api/policies/agents/" +
      encodeURIComponent(titleId);
    try {
      ppeAdapter.onPut(expectedUrl).reply(201, {}, { ETag: '"ppe"' });
      ppeAdapter.onGet(`${workIqPolicyClient.endpoint}/api/policies/contracts`).reply(200, {});
      ppeAdapter.onGet(expectedUrl).reply(404);
      expect((await ppe.publish(request(), provider)).isOk()).toBe(true);
      expect(ppeAdapter.history.put).toHaveLength(1);
      expect(ppeAdapter.history.put[0].url).toBe(expectedUrl);
    } finally {
      ppeAdapter.restore();
    }
  });

  it("WIQ-POL-25: Contract call precedes PUT, reuses the write token and does not validate its mocked body", async () => {
    const calls: string[] = [];
    adapter.onGet(`${endpoint}/api/policies/contracts`).reply(() => {
      calls.push("contract");
      return [200, "mocked-opaque-contract-response"];
    });
    adapter.onGet(url).reply(() => {
      calls.push("get-policy");
      return [404];
    });
    adapter.onPut(url).reply(() => {
      calls.push("put");
      return [201, {}, { ETag: '"published"' }];
    });
    const prepared = await client.prepare(request(), provider);
    expect(prepared.isOk()).toBe(true);
    expect(calls).toEqual(["contract", "get-policy"]);
    expect(adapter.history.put).toHaveLength(0);
    if (prepared.isOk()) expect((await prepared.value.publish()).isOk()).toBe(true);
    expect(calls).toEqual(["contract", "get-policy", "put"]);
    expect(provider.getAccessToken).toHaveBeenCalledExactlyOnceWith({ scopes: [scope] });
  });

  for (const status of [401, 503]) {
    it(`WIQ-POL-25: Contract ${status} fails preflight without a policy write`, async () => {
      adapter
        .onGet(`${endpoint}/api/policies/contracts`)
        .reply(status, { error: "private-contract-error" });
      const result = await client.prepare(request(), provider);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.message).not.toContain("private-contract-error");
      expect(adapter.history.put).toHaveLength(0);
    });
  }

  it("WIQ-POL-25: Contract timeout is a preflight read error, not an ambiguous policy write", async () => {
    adapter.onGet(`${endpoint}/api/policies/contracts`).timeout();
    const result = await client.prepare(request(), provider);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("WORKIQ_READ_FAILED");
    expect(adapter.history.put).toHaveLength(0);
    expect(provider.getAccessToken).toHaveBeenCalledOnce();
  });

  it("WIQ-POL-20: creates with a Work IQ token, appId transport binding, exact route and independent receipt", async () => {
    adapter.onPut(url).reply(201, { schemaVersion: 1, capabilities: [] }, { ETag: '"workiq-new"' });
    const input = request();
    const prepared = await client.prepare(input, provider);
    expect(prepared.isOk()).toBe(true);
    expect(adapter.history.put).toHaveLength(0);
    input.callerId = "changed";
    input.policy.bytes.fill(0);
    if (!prepared.isOk()) return;
    const result = await prepared.value.publish();
    expect(result.isOk()).toBe(true);
    if (result.isOk()) expect(result.value).toEqual({ state: "confirmed", etag: '"workiq-new"' });
    expect(JSON.parse(adapter.history.put[0].data)).toEqual({
      schemaVersion: 1,
      capabilities: [],
      appId: callerId,
    });
    expect(adapter.history.put[0].headers).toMatchObject({
      Authorization: "Bearer workiq-only-token",
      "If-None-Match": "*",
      "Content-Type": "application/json",
    });
    expect(adapter.history.put[0].headers).not.toHaveProperty("If-Match");
    expect(adapter.history.put[0].maxRedirects).toBe(0);
    expect(adapter.history.put[0].timeout).toBe(30000);
    expect(provider.getAccessToken).toHaveBeenCalledExactlyOnceWith({ scopes: [scope] });
  });

  it("WIQ-POL-20: replacement uses its GET ETag and returns the new ETag without a follow-up GET", async () => {
    adapter.onGet(url).reply(200, {}, { ETag: '"workiq-old"' });
    adapter.onPut(url).reply(200, {}, { ETag: '"replacement"' });
    const result = await client.publish(request(), provider);
    expect(result.isOk()).toBe(true);
    expect(adapter.history.put[0].headers).toHaveProperty("If-Match", '"workiq-old"');
    expect(adapter.history.put[0].headers).not.toHaveProperty("If-None-Match");
    expect(adapter.history.get.map((read) => read.url)).toEqual([
      `${endpoint}/api/policies/contracts`,
      url,
    ]);
    expect(provider.getAccessToken).toHaveBeenCalledOnce();
  });

  for (const [status, body, headers] of [
    [200, {}, {}],
    [200, {}, { ETag: 'W/"weak"' }],
    [200, [], { ETag: '"stored"' }],
    [304, undefined, { ETag: '"stored"' }],
    [401, { private: "read-error" }, {}],
    [503, { private: "read-error" }, {}],
  ] as const) {
    it(`WIQ-POL-27: invalid policy GET ${status} stops preparation without writing`, async () => {
      adapter.onGet(url).reply(status, body, headers);
      const result = await client.prepare(request(), provider);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(JSON.stringify(result.error)).not.toContain("read-error");
      expect(adapter.history.put).toHaveLength(0);
      expect(adapter.history.get.filter((read) => read.url === url)).toHaveLength(1);
    });
  }

  it("WIQ-POL-27: policy GET timeout stops preparation without retry or writing", async () => {
    adapter.onGet(url).timeout();
    const result = await client.prepare(request(), provider);
    expect(result.isErr()).toBe(true);
    expect(adapter.history.put).toHaveLength(0);
    expect(adapter.history.get.filter((read) => read.url === url)).toHaveLength(1);
  });

  it("WIQ-POL-27: cancellation during policy GET prevents publication", async () => {
    const controller = new AbortController();
    adapter.onGet(url).reply(() => {
      controller.abort();
      return [200, {}, { ETag: '"stored"' }];
    });
    const result = await client.prepare({ ...request(), signal: controller.signal }, provider);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("UserCancel");
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-22: stale GET ETag conflicts without re-reading or retrying", async () => {
    adapter.onGet(url).reply(200, {}, { ETag: '"stored"' });
    adapter.onPut(url).reply(412, {}, { ETag: '"newer"' });
    const result = await client.publish(request(), provider);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.state).toBe("conflict");
    expect(adapter.history.put).toHaveLength(1);
    expect(adapter.history.get.filter((read) => read.url === url)).toHaveLength(1);
    expect(adapter.history.put[0].headers).toHaveProperty("If-Match", '"stored"');
  });
  for (const appId of ["", "not-a-guid", "00000000-0000-0000-0000-000000000000"]) {
    it(`WIQ-POL-20: invalid runtime app ID '${appId}' fails before auth or HTTP`, async () => {
      const result = await client.prepare({ ...request(), callerId: appId }, provider);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.name).toBe("WORKIQ_APP_BINDING_INVALID");
      expect(provider.getAccessToken).not.toHaveBeenCalled();
      expect(adapter.history.put).toHaveLength(0);
    });
  }

  it("WIQ-POL-20: conflicting appId in the captured policy cannot override caller identity", async () => {
    const input = request();
    input.policy.bytes = Buffer.from(
      '{"schemaVersion":1,"capabilities":[],"appId":"22222222-2222-4333-8444-555555555555"}'
    );
    const result = await client.prepare(input, provider);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("WORKIQ_APP_BINDING_INVALID");
    expect(provider.getAccessToken).not.toHaveBeenCalled();
  });

  it("WIQ-POL-20: an explicit audience is required and is never inferred from the MOS token or host", async () => {
    const result = await new WorkIqPolicyClient(endpoint).prepare(request(), provider);
    expect(result.isErr()).toBe(true);
    expect(provider.getAccessToken).not.toHaveBeenCalled();
  });

  it("WIQ-POL-22: auth errors are safe and never start a request", async () => {
    vi.mocked(provider.getAccessToken).mockResolvedValue(
      err(
        new UserError({
          source: "auth",
          name: "Denied",
          message: "private-token https://private/account",
        })
      )
    );
    const result = await client.prepare(request(), provider);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.message).not.toContain("private");
    expect(adapter.history.put).toHaveLength(0);
  });

  for (const [status, state] of [
    [400, "failed"],
    [401, "failed"],
    [403, "failed"],
    [404, "failed"],
    [412, "conflict"],
    [413, "failed"],
    [422, "failed"],
    [428, "failed"],
    [429, "unknown"],
    [503, "unknown"],
    [409, "unknown"],
    [202, "unknown"],
  ]) {
    it(`WIQ-POL-22: PUT ${status} is ${state} without retry or raw diagnostics`, async () => {
      adapter.onPut(url).reply(Number(status), { error: "private-service-diagnostics" });
      const result = await client.publish(request(), provider);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error.state).toBe(state);
        expect(JSON.stringify(result.error)).not.toContain("private-service-diagnostics");
      }
      expect(adapter.history.put).toHaveLength(1);
    });
  }

  for (const headers of [{}, { ETag: 'W/"weak"' }]) {
    it("WIQ-POL-22: missing or malformed write acknowledgement is unknown", async () => {
      adapter.onPut(url).reply(201, {}, headers);
      const result = await client.publish(request(), provider);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.state).toBe("unknown");
    });
  }

  for (const replacement of [false, true]) {
    it("WIQ-POL-22: a success status inconsistent with the request condition is unknown", async () => {
      if (replacement) adapter.onGet(url).reply(200, {}, { ETag: '"previous"' });
      adapter.onPut(url).reply(replacement ? 201 : 200, {}, { ETag: '"unexpected-success"' });
      const result = await client.publish(request(), provider);
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.state).toBe("unknown");
      expect(adapter.history.put).toHaveLength(1);
    });
  }

  it("WIQ-POL-22: network timeout is unknown and does not retry", async () => {
    adapter.onPut(url).timeout();
    const result = await client.publish(request(), provider);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.state).toBe("unknown");
    expect(adapter.history.put).toHaveLength(1);
  });

  it("WIQ-POL-21: GET independently returns policy and ETag without expecting echoed appId", async () => {
    adapter.onGet(url).reply(200, { schemaVersion: 1, capabilities: [] }, { ETag: '"stored"' });
    const result = await client.getPolicy(titleId, provider);
    expect(result.isOk()).toBe(true);
    if (result.isOk())
      expect(result.value).toEqual({
        state: "found",
        etag: '"stored"',
        policy: { schemaVersion: 1, capabilities: [] },
      });
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-21: GET 404 is explicit not-found, not proof of full mapping reconciliation", async () => {
    adapter.onGet(url).reply(404);
    const result = await client.getPolicy(titleId, provider);
    expect(result.isOk()).toBe(true);
    if (result.isOk()) expect(result.value).toEqual({ state: "not-found" });
  });

  it("WIQ-POL-21: DELETE sends only the same appId and its original ETag", async () => {
    adapter.onDelete(url).reply(204);
    const result = await client.deletePolicy({ titleId, callerId, etag: '"original"' }, provider);
    expect(result.isOk()).toBe(true);
    expect(JSON.parse(adapter.history.delete[0].data)).toEqual({ appId: callerId });
    expect(adapter.history.delete[0].headers).toHaveProperty("If-Match", '"original"');
  });

  it("WIQ-POL-22: DELETE 503 remains unknown and never replays automatically", async () => {
    adapter.onDelete(url).reply(503, {}, { "Retry-After": "30" });
    const result = await client.deletePolicy({ titleId, callerId, etag: '"original"' }, provider);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.state).toBe("unknown");
    expect(adapter.history.delete).toHaveLength(1);
  });

  for (const invalidEndpoint of [
    "http://workiq.example",
    "https://user:password@workiq.example",
    `${endpoint}?`,
    `${endpoint}#`,
    "not-a-url",
  ]) {
    it("WIQ-POL-20: unsafe endpoints fail before token acquisition", async () => {
      const result = await new WorkIqPolicyClient(invalidEndpoint, [scope]).prepare(
        request(),
        provider
      );
      expect(result.isErr()).toBe(true);
      expect(provider.getAccessToken).not.toHaveBeenCalled();
    });
  }

  it("WIQ-POL-21: read-only scope permits inspection but not policy writes", async () => {
    const http = axios.create();
    const readAdapter = new MockAdapter(http);
    const readScope = "api://workiq-test/WorkIQSettings.Read.All";
    const readClient = new WorkIqPolicyClient(endpoint, [readScope], http);
    try {
      readAdapter.onGet(url).reply(200, { schemaVersion: 1, capabilities: [] }, { ETag: '"read"' });
      expect((await readClient.getPolicy(titleId, provider)).isOk()).toBe(true);
      expect((await readClient.prepare(request(), provider)).isErr()).toBe(true);
      expect(provider.getAccessToken).toHaveBeenCalledExactlyOnceWith({ scopes: [readScope] });
      expect(readAdapter.history.put).toHaveLength(0);
    } finally {
      readAdapter.restore();
    }
  });

  it("WIQ-POL-20: an unqualified permission name cannot accidentally request a Graph token", async () => {
    const result = await new WorkIqPolicyClient(endpoint, ["WorkIQSettings.ReadWrite.All"]).prepare(
      request(),
      provider
    );
    expect(result.isErr()).toBe(true);
    expect(provider.getAccessToken).not.toHaveBeenCalled();
  });

  it("WIQ-POL-20: mixed or partially unqualified audiences fail before token acquisition", async () => {
    for (const scopes of [
      [scope, "https://mos.example/.default"],
      [scope, "User.Read"],
    ]) {
      const result = await new WorkIqPolicyClient(endpoint, scopes).prepare(request(), provider);
      expect(result.isErr()).toBe(true);
    }
    expect(provider.getAccessToken).not.toHaveBeenCalled();
  });

  it("WIQ-POL-22: cancellation before auth performs no request", async () => {
    const controller = new AbortController();
    controller.abort();
    const result = await client.prepare({ ...request(), signal: controller.signal }, provider);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("UserCancel");
    expect(provider.getAccessToken).not.toHaveBeenCalled();
  });

  it("WIQ-POL-22: cancellation after auth prevents a write", async () => {
    const controller = new AbortController();
    vi.mocked(provider.getAccessToken).mockImplementation(async () => {
      controller.abort();
      return ok("token");
    });
    const result = await client.prepare({ ...request(), signal: controller.signal }, provider);
    expect(result.isErr()).toBe(true);
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-22: cancellation during PUT is unknown, not failed", async () => {
    const controller = new AbortController();
    adapter.onPut(url).reply(() => {
      controller.abort();
      return [201, {}, { ETag: '"saved"' }];
    });
    const result = await client.publish({ ...request(), signal: controller.signal }, provider);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.state).toBe("unknown");
  });

  it("WIQ-POL-22: thrown authentication diagnostics never escape", async () => {
    vi.mocked(provider.getAccessToken).mockRejectedValue(
      new Error("private-auth https://private/account")
    );
    const result = await client.prepare(request(), provider);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(JSON.stringify(result.error)).not.toContain("private");
  });

  it("WIQ-POL-21: malformed reads fail as read errors, not ambiguous writes", async () => {
    adapter.onGet(url).reply(200, { schemaVersion: 1 });
    const result = await client.getPolicy(titleId, provider);
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.name).toBe("WORKIQ_RESPONSE_INVALID");
      expect(result.error.message).not.toContain("remote write");
    }
  });

  it("WIQ-POL-21: read timeout and read rejection are safe errors", async () => {
    adapter.onGet(url).timeoutOnce().onGet(url).reply(403, { detail: "private-service" });
    const timeout = await client.getPolicy(titleId, provider);
    expect(timeout.isErr()).toBe(true);
    const denied = await client.getPolicy(titleId, provider);
    expect(denied.isErr()).toBe(true);
    if (denied.isErr()) expect(denied.error.name).toBe("WORKIQ_AUTH_REQUIRED");
  });

  it("WIQ-POL-21: weak deletion conditions fail before authentication", async () => {
    const result = await client.deletePolicy({ titleId, callerId, etag: 'W/"weak"' }, provider);
    expect(result.isErr()).toBe(true);
    expect(provider.getAccessToken).not.toHaveBeenCalled();
  });

  it("WIQ-POL-21: invalid agent or contract path segments do not escape their routes", async () => {
    expect((await client.getPolicy("..", provider)).isErr()).toBe(true);
    expect((await client.getCapabilityContract(".", "v1", provider)).isErr()).toBe(true);
    expect((await client.getCapabilityContract("Email", "..", provider)).isErr()).toBe(true);
    expect(provider.getAccessToken).not.toHaveBeenCalled();
  });

  it("WIQ-POL-21: malformed contract responses are not replaced with successful defaults", async () => {
    adapter.onGet(`${endpoint}/api/policies/contracts`).reply(200, "not-json");
    adapter.onGet(`${endpoint}/api/policies/contracts/Email/v1`).reply(200, []);
    expect((await client.getCapabilityContracts(provider)).isErr()).toBe(true);
    expect((await client.getCapabilityContract("Email", "v1", provider)).isErr()).toBe(true);
  });

  it("WIQ-POL-21: discovers tenant contracts and a specific configuration schema", async () => {
    adapter.onGet(`${endpoint}/api/policies/contracts`).reply(200, []);
    adapter
      .onGet(`${endpoint}/api/policies/contracts/Email/v1`)
      .reply(200, { configurationSchema: {} });
    expect((await client.getCapabilityContracts(provider)).isOk()).toBe(true);
    expect((await client.getCapabilityContract("Email", "v1", provider)).isOk()).toBe(true);
    expect(adapter.history.get).toHaveLength(2);
  });
});
