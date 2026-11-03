// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "reflect-metadata";
import {
  err,
  M365TokenProvider,
  ok,
  Platform,
  SystemError,
  UserError,
} from "@microsoft/teamsfx-api";
import axios from "axios";
import MockAdapter from "axios-mock-adapter";
import { createHash } from "crypto";
import { parse } from "dotenv";
import { mkdir, mkdtemp, readFile, rm, symlink, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import { Container } from "typedi";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { Lifecycle } from "../../../../src/component/configManager/lifecycle";
import { DriverContext } from "../../../../src/component/driver/interface/commonArgs";
import { WorkIqPolicyClient, workIqPolicyClient } from "../../../../src/client/workIqPolicyClient";
import {
  getResourceServiceEndpoint,
  MosServiceScope,
  ResourceServiceType,
} from "../../../../src/common/constants";
import {
  PolicyPublicationOperation,
  WorkIqPolicyPublisher,
} from "../../../../src/component/workiq/policyPublication";
import { PublishPolicyDriver } from "../../../../src/component/driver/workiq/publishPolicy";
import { PackageService } from "../../../../src/component/m365/packageService";
import { MockedAzureAccountProvider, MockedM365Provider, MockTools } from "../../../core/utils";
import { FxCoreClient } from "../../../../src/core/FxCoreClient";
import { MockedLogProvider } from "../../../plugins/solution/util";
import { UserCancelError } from "../../../../src/error/common";
import "../../../../src/component/driver";

const origin = getResourceServiceEndpoint(ResourceServiceType.MOS3);
const serviceOrigin = "https://regional.titles.example";
const baseUrl = `${serviceOrigin}/builder/v1`;
const scope = MosServiceScope()[0];
const discoveryUrl = `${origin}/config/v1/environment`;
const titleId = "title/with-special?characters";
const callerId = "11111111-2222-4333-8444-555555555555";
const url = `${baseUrl}/tenants/titles/${encodeURIComponent(titleId)}/policies/workiq`;
const policyBytes = Buffer.from('{\r\n  "rule": "${{DO_NOT_EXPAND}}", "allow": []\r\n}\r\n');

describe("workiq/publishPolicy Alpha", () => {
  let projectPath: string;
  let adapter: MockAdapter;
  let workIq: WorkIqPolicyClient;
  let tokenProvider: MockedM365Provider;
  const getToken = vi.fn<M365TokenProvider["getAccessToken"]>();

  const inputs = () => ({ titleId, callerId, policyPath: "./policy.json" });
  const mosService = () => new PackageService(origin, new MockedLogProvider(), tokenProvider);
  const operation = (client: WorkIqPolicyPublisher = workIq) =>
    new PolicyPublicationOperation({
      mos: { prepare: (request) => mosService().prepareWorkIqPolicy(request) },
      workIq: client,
    });

  beforeEach(async () => {
    projectPath = await mkdtemp(path.join(os.tmpdir(), "atk-workiq-policy-"));
    await writeFile(path.join(projectPath, "policy.json"), policyBytes);
    adapter = new MockAdapter(axios);
    const transport = axios.getAdapter(axios.defaults.adapter);
    axios.defaults.adapter = async (config) => ({
      ...(await transport(config)),
      request: { method: config.method?.toUpperCase() },
    });
    adapter.onGet(discoveryUrl).reply(200, { titlesServiceUrl: serviceOrigin });
    workIq = new WorkIqPolicyClient();
    vi.spyOn(workIqPolicyClient, "prepare").mockImplementation((request) =>
      workIq.prepare(request)
    );
    tokenProvider = new MockedM365Provider();
    tokenProvider.getAccessToken = getToken;
    getToken.mockClear();
    getToken.mockResolvedValue(ok("delegated-maker-token"));
  });

  afterEach(async () => {
    adapter.restore();
    vi.restoreAllMocks();
    vi.unstubAllEnvs();
    await rm(projectPath, { recursive: true, force: true });
  });

  it("WIQ-POL-03: sends captured bytes in one create-only MOS request and identifies the Work IQ mock", async () => {
    const publishWorkIq = vi.spyOn(workIq, "prepare");
    adapter.onPut(url).reply(async () => {
      await writeFile(path.join(projectPath, "policy.json"), '{"changed":true}');
      return [201, {}, { ETag: '"mos-version-1"' }];
    });

    const result = await operation().execute(inputs(), { projectPath });

    expect(result.isOk()).toBe(true);
    if (!result.isOk()) return;
    expect(result.value.policySha256).toBe(createHash("sha256").update(policyBytes).digest("hex"));
    expect(result.value.operationId).toMatch(/^[a-f0-9-]{36}$/);
    expect(result.value.mos).toEqual({ state: "confirmed", etag: '"mos-version-1"' });
    expect(result.value.workIq).toEqual({ state: "mocked" });
    expect(adapter.history.put).toHaveLength(1);
    expect(adapter.history.put[0].data).toEqual(policyBytes);
    expect(adapter.history.put[0].headers).toMatchObject({
      Authorization: "Bearer delegated-maker-token",
      "Content-Type": "application/json",
      "If-None-Match": "*",
    });
    expect(adapter.history.put[0].headers).not.toHaveProperty("If-Match");
    expect(adapter.history.put[0].maxRedirects).toBe(0);
    expect(adapter.history.put[0].timeout).toBe(30000);
    expect(getToken).toHaveBeenCalledExactlyOnceWith({ scopes: MosServiceScope() });
    expect(adapter.history.get).toHaveLength(1);
    expect(publishWorkIq).toHaveBeenCalledOnce();
    const request = publishWorkIq.mock.calls[0][0];
    expect(request.titleId).toBe(titleId);
    expect(request.callerId).toBe(callerId);
    expect(request.policy.bytes).toEqual(policyBytes);
    expect(request).not.toHaveProperty("token");
    expect(request.etag).toBeUndefined();
  });

  it("WIQ-POL-04: replaces only with the caller-supplied strong ETag", async () => {
    adapter.onPut(url).reply(204);
    const result = await operation().execute({ ...inputs(), etag: '"previous"' }, { projectPath });
    expect(result.isOk()).toBe(true);
    expect(adapter.history.put[0].headers).toHaveProperty("If-Match", '"previous"');
    expect(adapter.history.put[0].headers).not.toHaveProperty("If-None-Match");
  });

  for (const value of [
    null,
    {},
    { titleId, policyPath: "./policy.json" },
    { ...inputs(), callerId: " " },
    { ...inputs(), titleId: 123 },
    { ...inputs(), policyPath: "" },
    { ...inputs(), etag: "*" },
    { ...inputs(), etag: 'W/"weak"' },
    { ...inputs(), etag: '"injected"\r\nX-Test: true' },
    { ...inputs(), etag: '"trailing-newline"\n' },
    { ...inputs(), unexpected: true },
  ]) {
    it(`WIQ-POL-02: rejects invalid inputs ${JSON.stringify(value)}`, async () => {
      const publishWorkIq = vi.spyOn(workIq, "prepare");
      const result = await operation().execute(value, { projectPath });
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error).toBeInstanceOf(UserError);
        expect(result.error.report.mos.state).toBe("not-started");
        expect(result.error.report.workIq.state).toBe("not-started");
      }
      expect(getToken).not.toHaveBeenCalled();
      expect(adapter.history.put).toHaveLength(0);
      expect(publishWorkIq).not.toHaveBeenCalled();
    });
  }

  for (const content of ["not-json", "null", "[]", '"string"', Buffer.from([0x7b, 0x22, 0xff])]) {
    it(`WIQ-POL-02: rejects invalid JSON objects or encoding ${String(content)}`, async () => {
      await writeFile(path.join(projectPath, "policy.json"), content);
      const result = await operation().execute(inputs(), { projectPath });
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.name).toBe("InvalidWorkIqPolicy");
      expect(getToken).not.toHaveBeenCalled();
      expect(adapter.history.put).toHaveLength(0);
    });
  }

  it("WIQ-POL-02: reports unreadable policy files without destination calls", async () => {
    const result = await operation().execute(
      { ...inputs(), policyPath: "./missing.json" },
      { projectPath }
    );
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.name).toBe("WorkIqPolicyReadFailed");
    expect(getToken).not.toHaveBeenCalled();
  });

  it("WIQ-POL-02: rejects lexical and symlink escapes outside the project", async () => {
    const outside = await mkdtemp(path.join(os.tmpdir(), "atk-workiq-outside-"));
    try {
      await writeFile(path.join(outside, "policy.json"), policyBytes);
      await symlink(outside, path.join(projectPath, "external"), "junction");
      for (const policyPath of [
        path.join(outside, "policy.json"),
        path.join("external", "policy.json"),
      ]) {
        const result = await operation().execute({ ...inputs(), policyPath }, { projectPath });
        expect(result.isErr()).toBe(true);
        if (result.isErr()) expect(result.error.name).toBe("PolicyPathError");
      }
      expect(getToken).not.toHaveBeenCalled();
    } finally {
      await rm(path.join(projectPath, "external"), { recursive: true, force: true });
      await rm(outside, { recursive: true, force: true });
    }
  });

  for (const [status, state] of [
    [400, "failed"],
    [401, "failed"],
    [403, "failed"],
    [404, "failed"],
    [408, "unknown"],
    [418, "unknown"],
    [429, "unknown"],
    [409, "conflict"],
    [412, "conflict"],
    [202, "unknown"],
    [302, "unknown"],
    [500, "unknown"],
    [503, "unknown"],
  ]) {
    it(`WIQ-POL-05: preserves HTTP ${status} as ${state} without calling Work IQ or retrying`, async () => {
      const publishWorkIq = vi.spyOn(workIq, "prepare");
      adapter.onPut(url).reply(Number(status), { sensitiveBody: "do-not-log" });
      const result = await operation().execute(inputs(), { projectPath });
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error.report.mos.state).toBe(state);
        expect(result.error.report.workIq.state).toBe("not-started");
        expect(result.error.message).not.toContain("do-not-log");
      }
      expect(adapter.history.put).toHaveLength(1);
      expect(publishWorkIq).toHaveBeenCalledOnce();
    });
  }

  it("WIQ-POL-05: treats a MOS timeout as unknown, not as proof that no write occurred", async () => {
    adapter.onPut(url).timeout();
    const result = await operation().execute(inputs(), { projectPath });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.report.mos.state).toBe("unknown");
    expect(adapter.history.put).toHaveLength(1);
  });

  for (const state of ["failed", "unknown"] as const) {
    it(`WIQ-POL-06: retains MOS confirmation when Work IQ is ${state}`, async () => {
      adapter.onPut(url).reply(200, {}, { ETag: '"saved"' });
      const cause = new SystemError({ source: "workiq", name: "TestFailure", message: "test" });
      const result = await operation({
        prepare: async () => ok({ publish: async () => err({ state, error: cause }) }),
      }).execute(inputs(), { projectPath });
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error).not.toBe(cause);
        expect(cause).not.toHaveProperty("report");
        expect(result.error.report.outcome).toBe(state === "unknown" ? "unknown" : "incomplete");
        expect(result.error.report.mos).toEqual({ state: "confirmed", etag: '"saved"' });
        expect(result.error.report.workIq.state).toBe(state);
      }
    });
  }

  it("WIQ-POL-06: classifies an unexpected Work IQ transport exception as unknown", async () => {
    adapter.onPut(url).reply(200);
    const result = await operation({
      prepare: async () =>
        ok({
          publish: async () => {
            throw new Error("transport interrupted");
          },
        }),
    }).execute(inputs(), { projectPath });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.report.mos.state).toBe("confirmed");
      expect(result.error.report.workIq.state).toBe("unknown");
    }
  });

  it("WIQ-POL-07: pre-cancelled publication never acquires a token or writes", async () => {
    const controller = new AbortController();
    controller.abort();
    const result = await operation().execute(inputs(), { projectPath, signal: controller.signal });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.name).toBe("UserCancel");
      expect(result.error.report.mos.state).toBe("not-started");
    }
    expect(getToken).not.toHaveBeenCalled();
  });

  it("WIQ-POL-07: cancellation during a MOS request retains an unknown outcome and skips Work IQ", async () => {
    const controller = new AbortController();
    const publishWorkIq = vi.spyOn(workIq, "prepare");
    adapter.onPut(url).reply(() => {
      controller.abort();
      return [200, {}];
    });
    const result = await operation().execute(inputs(), { projectPath, signal: controller.signal });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.report.mos.state).toBe("unknown");
      expect(result.error.report.workIq.state).toBe("not-started");
    }
    expect(adapter.history.put[0].signal).toBe(controller.signal);
    expect(publishWorkIq).toHaveBeenCalledOnce();
  });

  it("WIQ-POL-07: cancellation after a received MOS acknowledgement preserves confirmation", async () => {
    const controller = new AbortController();
    const publishWorkIq = vi.spyOn(workIq, "prepare");
    const mos = mosService();
    adapter.onPut(url).reply(200);
    const result = await new PolicyPublicationOperation({
      mos: {
        prepare: async (request) => {
          const prepared = await mos.prepareWorkIqPolicy(request);
          if (prepared.isErr()) return prepared;
          return ok({
            publish: async () => {
              const acknowledged = await prepared.value.publish();
              controller.abort();
              return acknowledged;
            },
          });
        },
      },
      workIq,
    }).execute(inputs(), { projectPath, signal: controller.signal });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.report.mos.state).toBe("confirmed");
      expect(result.error.report.workIq.state).toBe("not-started");
    }
    expect(publishWorkIq).toHaveBeenCalledOnce();
  });

  for (const discovered of [
    undefined,
    "",
    "http://regional.titles.example",
    "https://user:password@regional.titles.example",
    `${serviceOrigin}?token=unsafe`,
    `${serviceOrigin}?`,
    `${serviceOrigin}#`,
    "not-a-url",
  ]) {
    it(`WIQ-POL-08: rejects an invalid discovered MOS service ${discovered}`, async () => {
      adapter.onGet(discoveryUrl).reply(200, { titlesServiceUrl: discovered });
      const result = await operation().execute(inputs(), { projectPath });
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error).toBeInstanceOf(SystemError);
        expect(result.error.report.mos.state).toBe("not-started");
      }
      expect(adapter.history.put).toHaveLength(0);
    });
  }

  it("WIQ-POL-08: discovery errors prevent policy writes and Work IQ publication", async () => {
    adapter.onGet(discoveryUrl).reply(403);
    const publishWorkIq = vi.spyOn(workIq, "prepare");
    const result = await operation().execute(inputs(), { projectPath });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.report.mos.state).toBe("not-started");
    expect(adapter.history.put).toHaveLength(0);
    expect(publishWorkIq).toHaveBeenCalledOnce();
  });

  it("WIQ-POL-08: PackageService without a delegated token provider cannot publish a policy", async () => {
    const result = await new PackageService(origin).publishWorkIqPolicy({
      titleId,
      policy: { bytes: policyBytes, sha256: "captured-digest" },
    });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.state).toBe("not-started");
    expect(adapter.history.get).toHaveLength(0);
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-07: cancellation during service discovery prevents the policy write", async () => {
    const controller = new AbortController();
    adapter.onGet(discoveryUrl).reply(() => {
      controller.abort();
      return [200, { titlesServiceUrl: serviceOrigin }];
    });
    const result = await operation().execute(inputs(), { projectPath, signal: controller.signal });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.name).toBe("UserCancel");
      expect(result.error.report.mos.state).toBe("not-started");
    }
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-08: preserves login errors without sending an HTTP request", async () => {
    const failure = new UserError({ source: "auth", name: "LoginDenied", message: "Login denied" });
    getToken.mockResolvedValueOnce(err(failure));
    const result = await operation().execute(inputs(), { projectPath });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error).not.toBe(failure);
      expect(result.error.name).toBe("POLICY_AUTH_REQUIRED");
      expect(failure).not.toHaveProperty("report");
      expect(result.error.report.mos.state).toBe("not-started");
    }
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-08: rejects an empty token without attempting a MOS write", async () => {
    getToken.mockResolvedValueOnce(ok(""));
    const result = await operation().execute(inputs(), { projectPath });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.report.mos.state).toBe("not-started");
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-04: PackageService cannot bypass strong-ETag validation", async () => {
    const result = await mosService().publishWorkIqPolicy({
      titleId,
      policy: { bytes: policyBytes, sha256: "captured-digest" },
      etag: "*",
    });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.state).toBe("not-started");
    expect(getToken).not.toHaveBeenCalled();
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-08: cancellation during login prevents a subsequent MOS write", async () => {
    const controller = new AbortController();
    getToken.mockImplementationOnce(async () => {
      controller.abort();
      return ok("maker");
    });
    const result = await operation().execute(inputs(), { projectPath, signal: controller.signal });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.report.mos.state).toBe("not-started");
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-08: authentication exceptions do not leak their original error message", async () => {
    getToken.mockRejectedValueOnce(new Error("sensitive-authentication-content"));
    const result = await operation().execute(inputs(), { projectPath });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.report.mos.state).toBe("not-started");
      expect(JSON.stringify(result.error)).not.toContain("sensitive-authentication-content");
    }
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-05: malformed MOS acknowledgement ETags cannot report confirmed publication", async () => {
    adapter.onPut(url).reply(200, {}, { ETag: 'W/"weak"' });
    const result = await operation().execute(inputs(), { projectPath });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.report.mos.state).toBe("unknown");
  });

  it("WIQ-POL-07: the mock Work IQ client respects a pre-cancelled request", async () => {
    const controller = new AbortController();
    controller.abort();
    const result = await workIq.publish({
      titleId,
      callerId,
      policy: { bytes: policyBytes, sha256: "captured-digest" },
      signal: controller.signal,
    });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.state).toBe("not-started");
  });

  it("WIQ-POL-13: Work IQ has an empty code-level endpoint and an explicitly mocked acknowledgement", async () => {
    expect(workIq.endpoint).toBe("");
    const result = await workIq.publish({
      titleId,
      callerId,
      policy: { bytes: policyBytes, sha256: "captured-digest" },
    });
    expect(result.isOk()).toBe(true);
    if (result.isOk()) expect(result.value.state).toBe("mocked");
    expect(adapter.history.get).toHaveLength(0);
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-13: an unsupported nonempty Work IQ endpoint cannot silently simulate success", async () => {
    const result = await new WorkIqPolicyClient("https://workiq.example").publish({
      titleId,
      callerId,
      policy: { bytes: policyBytes, sha256: "captured-digest" },
    });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) expect(result.error.state).toBe("not-started");
    expect(adapter.history.put).toHaveLength(0);
  });

  const context = (): DriverContext => ({
    projectPath,
    platform: Platform.CLI,
    azureAccountProvider: new MockedAzureAccountProvider(),
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    ui: undefined,
    progressBar: undefined,
  });

  it("WIQ-POL-14: a known unsupported Work IQ endpoint prevents every destination write", async () => {
    adapter.onPut(url).reply(201);
    const result = await operation(new WorkIqPolicyClient("https://unimplemented.example")).execute(
      inputs(),
      { projectPath }
    );
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.report.mos.state).toBe("not-started");
      expect(result.error.report.workIq.state).toBe("not-started");
    }
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-17: duplicate decoded JSON keys fail before auth or publication", async () => {
    await writeFile(path.join(projectPath, "policy.json"), '{"rule":1,"\\u0072ule":2}');
    adapter.onPut(url).reply(201);
    const result = await operation().execute(inputs(), { projectPath });
    expect(result.isErr()).toBe(true);
    expect(getToken).not.toHaveBeenCalled();
    expect(adapter.history.put).toHaveLength(0);
  });

  for (const content of [
    '{"rules":[{"a":1,"a":2}]}',
    '{"nested":{"a":1,"\\u0061":2}}',
    '{"__proto__":{},"__proto__":{}}',
  ]) {
    it(`WIQ-POL-17: rejects nested duplicate keys in ${content}`, async () => {
      await writeFile(path.join(projectPath, "policy.json"), content);
      const result = await operation().execute(inputs(), { projectPath });
      expect(result.isErr()).toBe(true);
      if (result.isErr()) expect(result.error.report.mos.state).toBe("not-started");
      expect(getToken).not.toHaveBeenCalled();
    });
  }

  for (const content of [
    '{"a":{"x":1},"b":{"x":2}}',
    '{"rules":[{"a":1},{"a":2}]}',
    '{"a\u2028b":1,"a b":2,"braces":"{}[]","quote":"\\""}',
  ]) {
    it("WIQ-POL-17: accepts distinct JSON keys and Unicode without changing bytes", async () => {
      await writeFile(path.join(projectPath, "policy.json"), content);
      adapter.onPut(url).reply(201);
      const result = await operation().execute(inputs(), { projectPath });
      expect(result.isOk()).toBe(true);
      expect(adapter.history.put[0].data).toEqual(Buffer.from(content));
    });
  }

  it("WIQ-POL-14: a thrown Work IQ preflight failure cannot start a MOS write", async () => {
    const result = await operation({
      prepare: async () => {
        throw new Error("private-auth-context");
      },
    }).execute(inputs(), { projectPath });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.report.mos.state).toBe("not-started");
      expect(result.error.report.outcome).toBe("failed");
      expect(JSON.stringify(result.error)).not.toContain("private-auth-context");
    }
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-14: an auth-dialog cancellation thrown during preflight remains cancelled", async () => {
    const result = await operation({
      prepare: async () => {
        throw new UserCancelError();
      },
    }).execute(inputs(), { projectPath });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.name).toBe("UserCancel");
      expect(result.error.report.outcome).toBe("cancelled");
      expect(result.error.report.mos.state).toBe("not-started");
    }
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-16: a thrown cancellation during Work IQ submission remains unknown and requires reconciliation", async () => {
    adapter.onPut(url).reply(201);
    const result = await operation({
      prepare: async () =>
        ok({
          publish: async () => {
            throw new UserCancelError();
          },
        }),
    }).execute(inputs(), { projectPath });
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.report.outcome).toBe("cancelled");
      expect(result.error.report.workIq.state).toBe("unknown");
      expect(result.error.report.nextAction).toBe("reconcile");
    }
  });

  it("WIQ-POL-19: prepared MOS publication retains auth and performs no write until invoked", async () => {
    adapter.onPut(url).reply(201);
    const prepared = await mosService().prepareWorkIqPolicy({
      titleId,
      policy: { bytes: policyBytes, sha256: "captured-digest" },
    });
    expect(prepared.isOk()).toBe(true);
    expect(adapter.history.put).toHaveLength(0);
    getToken.mockResolvedValue(
      err(new UserError({ source: "auth", name: "Expired", message: "private" }))
    );
    if (prepared.isOk()) expect((await prepared.value.publish()).isOk()).toBe(true);
    expect(getToken).toHaveBeenCalledOnce();
    expect(adapter.history.get).toHaveLength(1);
    expect(adapter.history.put).toHaveLength(1);
  });

  it("WIQ-POL-19: prepared MOS contexts snapshot the validated identity, condition and policy bytes", async () => {
    const request = {
      titleId,
      etag: '"validated-etag"',
      policy: { bytes: Buffer.from(policyBytes), sha256: "captured-digest" },
    };
    adapter.onPut(url).reply(201);
    const prepared = await mosService().prepareWorkIqPolicy(request);
    expect(prepared.isOk()).toBe(true);
    request.titleId = "different-title";
    request.etag = "*";
    request.policy.bytes.fill(0);
    if (prepared.isOk()) expect((await prepared.value.publish()).isOk()).toBe(true);
    expect(adapter.history.put[0].url).toContain(encodeURIComponent(titleId));
    expect(adapter.history.put[0].headers).toHaveProperty("If-Match", '"validated-etag"');
    expect(adapter.history.put[0].data).toEqual(policyBytes);
  });

  it("WIQ-POL-16: lifecycle cancellation preserves the MOS receipt and successful substep", async () => {
    const controller = new AbortController();
    const publish = vi
      .spyOn(PackageService.prototype, "prepareWorkIqPolicy")
      .mockImplementation(async () => {
        return ok({
          publish: async () => {
            controller.abort();
            return ok({ state: "confirmed", etag: '"saved-before-cancel"' });
          },
        });
      });
    const ctx = { ...context(), signal: controller.signal };
    const result = await new Lifecycle(
      "publish",
      [{ uses: "workiq/publishPolicy", with: inputs() }],
      "v1.13"
    ).execute(ctx);
    expect(publish).toHaveBeenCalledOnce();
    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) {
      expect(result.result.error).toHaveProperty("reason.error.report.mos", {
        state: "confirmed",
        etag: '"saved-before-cancel"',
      });
    }
    expect(result.summaries.flat().join(" ")).toContain("MOS policy publication was confirmed");
  });

  it("WIQ-POL-16: cancellation after a successful driver return still carries its receipt", async () => {
    adapter.onPut(url).reply(201);
    const controller = new AbortController();
    const ctx = { ...context(), signal: controller.signal };
    vi.spyOn(ctx.m365TokenProvider, "getAccessToken").mockResolvedValue(ok("maker"));
    const driver = new PublishPolicyDriver();
    const execute = driver.execute.bind(driver);
    vi.spyOn(driver, "execute").mockImplementation(async (args, context, outputs) => {
      const result = await execute(args, context, outputs);
      controller.abort();
      return result;
    });
    const name = "test/policy-late-cancel";
    Container.set(name, driver);
    try {
      const result = await new Lifecycle(
        "publish",
        [{ uses: name, with: inputs() }],
        "v1.13"
      ).execute(ctx);
      expect(result.result.isErr()).toBe(true);
      if (result.result.isErr()) {
        expect(result.result.error).toHaveProperty(
          "reason.error.report",
          expect.objectContaining({
            outcome: "cancelled",
            mos: { state: "confirmed" },
            workIq: { state: "mocked" },
          })
        );
      }
    } finally {
      Container.remove(name);
    }
  });

  it("WIQ-POL-16: lifecycle cancellation preserves an in-flight unknown outcome and reconcile guidance", async () => {
    const controller = new AbortController();
    const ctx = { ...context(), signal: controller.signal };
    vi.spyOn(ctx.m365TokenProvider, "getAccessToken").mockResolvedValue(ok("maker"));
    adapter.onPut(url).reply(() => {
      controller.abort();
      return [200, {}];
    });

    const result = await new Lifecycle(
      "publish",
      [{ uses: "workiq/publishPolicy", with: inputs() }],
      "v1.13"
    ).execute(ctx);
    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) {
      expect(result.result.error).toHaveProperty(
        "reason.error.report",
        expect.objectContaining({
          outcome: "cancelled",
          nextAction: "reconcile",
          mos: { state: "unknown" },
          workIq: { state: "not-started" },
        })
      );
    }
  });

  it("WIQ-POL-16: the public in-process client preserves cancelled publication receipts", async () => {
    vi.stubEnv("TEAMSFX_ENV", undefined);
    await mkdir(path.join(projectPath, "env"));
    await writeFile(path.join(projectPath, "env", ".env.dev"), "TEAMSFX_ENV=dev\n");
    await writeFile(
      path.join(projectPath, "m365agents.yml"),
      [
        "version: v1.13",
        "environmentFolderPath: ./env",
        "publish:",
        "  - uses: workiq/publishPolicy",
        "    with:",
        "      titleId: title-id",
        `      callerId: ${callerId}`,
        "      policyPath: ./policy.json",
        "",
      ].join("\n")
    );
    const controller = new AbortController();
    vi.spyOn(PackageService.prototype, "prepareWorkIqPolicy").mockResolvedValue(
      ok({
        publish: async () => {
          controller.abort();
          return ok({ state: "confirmed", etag: '"public-client-receipt"' });
        },
      })
    );
    const result = await new FxCoreClient(new MockTools()).publish(
      { platform: Platform.CLI, projectPath, env: "dev", nonInteractive: true },
      { signal: controller.signal }
    );
    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error).toHaveProperty(
        "report",
        expect.objectContaining({
          outcome: "cancelled",
          mos: { state: "confirmed", etag: '"public-client-receipt"' },
          workIq: { state: "not-started" },
        })
      );
    }
  });

  it("WIQ-POL-08: the native action reuses PackageService discovery without policy-specific environment configuration", async () => {
    const globalAdapter = adapter;
    try {
      const discovered = "https://regional.titles.example";
      const discovery = vi
        .spyOn(PackageService.prototype, "getTitleServiceUrl")
        .mockResolvedValue(discovered);
      const ctx = context();
      const token = vi
        .spyOn(ctx.m365TokenProvider, "getAccessToken")
        .mockResolvedValue(ok("maker"));
      globalAdapter
        .onPut(
          `${discovered}/builder/v1/tenants/titles/${encodeURIComponent(titleId)}/policies/workiq`
        )
        .reply(201);
      const result = await new PublishPolicyDriver().execute(inputs(), ctx);
      expect(result.result.isOk()).toBe(true);
      expect(discovery).toHaveBeenCalledExactlyOnceWith("maker", undefined);
      expect(token).toHaveBeenCalledExactlyOnceWith({ scopes: [scope] });
    } finally {
      globalAdapter.restore();
    }
  });

  it("WIQ-POL-09: runs the native publish lifecycle with resolved inputs and an explicit mock report", async () => {
    const globalAdapter = adapter;
    try {
      vi.stubEnv("M365_TITLE_ID", titleId);
      vi.stubEnv("WORKIQ_CALLER_ID", callerId);
      vi.stubEnv("TEAMSFX_ENV", "dev");
      vi.stubEnv("WORKIQ_POLICY_PUBLICATION_REPORT", undefined);
      await mkdir(path.join(projectPath, "env"));
      await writeFile(
        path.join(projectPath, "m365agents.yml"),
        "version: v1.13\nenvironmentFolderPath: ./env\n"
      );
      await writeFile(path.join(projectPath, "env", ".env.dev"), "TEAMSFX_ENV=dev\n");
      const ctx = context();
      const token = vi
        .spyOn(ctx.m365TokenProvider, "getAccessToken")
        .mockResolvedValue(ok("maker"));
      const warning = vi.spyOn(ctx.logProvider, "warning");
      const publishWorkIq = vi.mocked(workIqPolicyClient.prepare);
      globalAdapter.onPut(url).reply(201);
      const lifecycle = new Lifecycle(
        "publish",
        [
          {
            uses: "workiq/publishPolicy",
            with: {
              titleId: "${{M365_TITLE_ID}}",
              policyPath: "./policy.json",
              callerId: "${{WORKIQ_CALLER_ID}}",
            },
            writeToEnvironmentFile: { report: "WORKIQ_POLICY_PUBLICATION_REPORT" },
          },
        ],
        "v1.13"
      );
      const result = await lifecycle.execute(ctx);
      expect(result.result.isOk()).toBe(true);
      if (result.result.isOk()) {
        const report: unknown = JSON.parse(
          result.result.value.get("WORKIQ_POLICY_PUBLICATION_REPORT") ?? "null"
        );
        expect(report).toMatchObject({ mos: { state: "confirmed" }, workIq: { state: "mocked" } });
        const persisted = parse(await readFile(path.join(projectPath, "env", ".env.dev")));
        expect(persisted.WORKIQ_POLICY_PUBLICATION_REPORT).toBe(
          result.result.value.get("WORKIQ_POLICY_PUBLICATION_REPORT")
        );
      }
      expect(warning).not.toHaveBeenCalled();
      expect(token).toHaveBeenCalledExactlyOnceWith({ scopes: [scope] });
      expect(publishWorkIq).toHaveBeenCalledOnce();
      expect(await readFile(path.join(projectPath, "policy.json"))).toEqual(policyBytes);
    } finally {
      globalAdapter.restore();
    }
  });
  it("WIQ-POL-10: returns failed destination summaries without leaking request or response content", async () => {
    const globalAdapter = adapter;
    try {
      const ctx = context();
      vi.spyOn(ctx.m365TokenProvider, "getAccessToken").mockResolvedValue(ok("sensitive-token"));
      const logError = vi.spyOn(ctx.logProvider, "error");
      globalAdapter.onPut(url).reply(503, { body: "sensitive-response" });
      const result = await new PublishPolicyDriver().execute(inputs(), ctx);
      expect(result.result.isErr()).toBe(true);
      expect(result.report?.mos.state).toBe("unknown");
      expect(result.summaries).toEqual([]);
      if (result.result.isErr()) {
        expect(result.result.error).toHaveProperty(
          "displayMessage",
          expect.stringContaining("MOS=unknown")
        );
      }
      expect(logError).toHaveBeenCalled();
      const messages = JSON.stringify([result.summaries, logError.mock.calls]);
      expect(messages).not.toContain("sensitive-token");
      expect(messages).not.toContain("sensitive-response");
      expect(messages).not.toContain("DO_NOT_EXPAND");
    } finally {
      globalAdapter.restore();
    }
  });

  it("WIQ-POL-23: native publication keeps MOS and runtime app bindings, tokens and ETags independent", async () => {
    const calls: string[] = [];
    const workIqEndpoint = "https://workiq.example/environment";
    const workIqScope = "api://workiq-test/WorkIQSettings.ReadWrite.All";
    adapter.onGet(`${workIqEndpoint}/api/policies/contracts`).reply(() => {
      calls.push("contract");
      return [200, "mocked-opaque-response"];
    });
    adapter.onGet(discoveryUrl).reply(() => {
      calls.push("mos-prepare");
      return [200, { titlesServiceUrl: serviceOrigin }];
    });
    const liveClient = new WorkIqPolicyClient(workIqEndpoint, [workIqScope]);
    vi.mocked(workIqPolicyClient.prepare).mockImplementation((request, provider) =>
      liveClient.prepare(request, provider)
    );
    const ctx = context();
    const token = vi
      .spyOn(ctx.m365TokenProvider, "getAccessToken")
      .mockImplementation(async (request) =>
        ok(request.scopes.includes(workIqScope) ? "workiq-only" : "mos-only")
      );
    const content = Buffer.from('{\r\n "schemaVersion":1, "capabilities":[]\r\n}');
    await writeFile(path.join(projectPath, "policy.json"), content);
    adapter.onPut(url).reply(() => {
      calls.push("mos");
      return [200, {}, { ETag: '"mos-new"' }];
    });
    const workIqUrl = `${workIqEndpoint}/api/policies/agents/${encodeURIComponent(titleId)}`;
    adapter.onPut(workIqUrl).reply(() => {
      calls.push("workiq");
      return [200, {}, { ETag: '"workiq-new"' }];
    });
    const result = await new PublishPolicyDriver().execute(
      {
        ...inputs(),
        etag: '"mos-old"',
        workIqEtag: '"workiq-old"',
      },
      ctx
    );
    expect(result.result.isOk()).toBe(true);
    expect(result.report).toMatchObject({
      outcome: "published",
      enforcement: "not-verified",
      mos: { state: "confirmed", etag: '"mos-new"' },
      workIq: { state: "confirmed", etag: '"workiq-new"' },
    });
    const mosWrite = adapter.history.put[0];
    const runtimeWrite = adapter.history.put[1];
    expect(mosWrite.data).toEqual(content);
    expect(mosWrite.headers).toMatchObject({
      Authorization: "Bearer mos-only",
      "If-Match": '"mos-old"',
    });
    expect(runtimeWrite.headers).toMatchObject({
      Authorization: "Bearer workiq-only",
      "If-Match": '"workiq-old"',
    });
    expect(JSON.parse(runtimeWrite.data)).toEqual({
      schemaVersion: 1,
      capabilities: [],
      appId: callerId,
    });
    expect(await readFile(path.join(projectPath, "policy.json"))).toEqual(content);
    expect(token).toHaveBeenCalledTimes(2);
    expect(calls).toEqual(["contract", "mos-prepare", "mos", "workiq"]);
  });

  it("WIQ-POL-26: shared MOS scopes and one provider token work across Contract, MOS and Work IQ requests", async () => {
    const endpoint = "https://workiq.example/environment";
    const sharedScopes = MosServiceScope();
    const client = new WorkIqPolicyClient(endpoint, sharedScopes);
    vi.mocked(workIqPolicyClient.prepare).mockImplementation((request, provider) =>
      client.prepare(request, provider)
    );
    const ctx = context();
    const token = vi
      .spyOn(ctx.m365TokenProvider, "getAccessToken")
      .mockResolvedValue(ok("shared-mos-token"));
    adapter.onGet(`${endpoint}/api/policies/contracts`).reply(200, {});
    adapter.onPut(url).reply(201, {}, { ETag: '"mos-shared"' });
    adapter
      .onPut(`${endpoint}/api/policies/agents/${encodeURIComponent(titleId)}`)
      .reply(201, {}, { ETag: '"runtime-shared"' });
    const result = await new PublishPolicyDriver().execute(inputs(), ctx);
    expect(result.result.isOk()).toBe(true);
    expect(result.report).toMatchObject({
      outcome: "published",
      mos: { state: "confirmed", etag: '"mos-shared"' },
      workIq: { state: "confirmed", etag: '"runtime-shared"' },
    });
    expect(token.mock.calls.map(([request]) => request.scopes)).toEqual([
      sharedScopes,
      sharedScopes,
    ]);
    for (const request of [...adapter.history.get, ...adapter.history.put]) {
      expect(request.headers).toMatchObject({ Authorization: "Bearer shared-mos-token" });
    }
  });

  it("WIQ-POL-25: Contract failure stops the native action before MOS authentication or writes", async () => {
    const endpoint = "https://workiq.example/environment";
    const scope = "api://workiq-test/WorkIQSettings.ReadWrite.All";
    const client = new WorkIqPolicyClient(endpoint, [scope]);
    vi.mocked(workIqPolicyClient.prepare).mockImplementation((request, provider) =>
      client.prepare(request, provider)
    );
    const ctx = context();
    const token = vi
      .spyOn(ctx.m365TokenProvider, "getAccessToken")
      .mockResolvedValue(ok("workiq-token"));
    adapter.onGet(`${endpoint}/api/policies/contracts`).reply(503, { detail: "private-service" });
    const result = await new PublishPolicyDriver().execute(inputs(), ctx);
    expect(result.result.isErr()).toBe(true);
    expect(result.report).toMatchObject({
      outcome: "failed",
      mos: { state: "not-started" },
      workIq: { state: "not-started" },
    });
    expect(token).toHaveBeenCalledExactlyOnceWith({ scopes: [scope] });
    expect(adapter.history.get.map((request) => request.url)).toEqual([
      `${endpoint}/api/policies/contracts`,
    ]);
    expect(adapter.history.put).toHaveLength(0);
  });

  it("WIQ-POL-01: the native action runs without an environment output mapping", async () => {
    const globalAdapter = adapter;
    try {
      const ctx = context();
      vi.spyOn(ctx.m365TokenProvider, "getAccessToken").mockResolvedValue(ok("maker"));
      globalAdapter.onPut(url).reply(201);
      const result = await new PublishPolicyDriver().execute(inputs(), ctx);
      expect(result.result.isOk()).toBe(true);
      if (result.result.isOk()) expect(result.result.value.size).toBe(0);
      expect(result.report?.workIq.state).toBe("mocked");
    } finally {
      globalAdapter.restore();
    }
  });

  it("WIQ-POL-10: only confirmed MOS writes become successful substep summaries after Work IQ failure", async () => {
    const globalAdapter = adapter;
    try {
      const ctx = context();
      vi.spyOn(ctx.m365TokenProvider, "getAccessToken").mockResolvedValue(ok("maker"));
      vi.spyOn(WorkIqPolicyClient.prototype, "prepare").mockResolvedValue(
        ok({
          publish: async () =>
            err({
              state: "unknown",
              error: new SystemError({ source: "workiq", name: "TransportError", message: "test" }),
            }),
        })
      );
      globalAdapter.onPut(url).reply(201);
      const result = await new PublishPolicyDriver().execute(inputs(), ctx);
      expect(result.result.isErr()).toBe(true);
      expect(result.report?.mos.state).toBe("confirmed");
      expect(result.report?.workIq.state).toBe("unknown");
      expect(result.summaries).toEqual(["MOS policy publication was confirmed."]);
      if (result.result.isErr()) {
        expect(result.result.error).toHaveProperty(
          "displayMessage",
          expect.stringContaining("Work IQ=unknown")
        );
      }
    } finally {
      globalAdapter.restore();
    }
  });

  it("WIQ-POL-10: invalid output mappings fail before token acquisition or destination writes", async () => {
    const ctx = context();
    const token = vi.spyOn(ctx.m365TokenProvider, "getAccessToken");
    const result = await new PublishPolicyDriver().execute(
      inputs(),
      ctx,
      new Map([["success", "PUBLISHED"]])
    );
    expect(result.result.isErr()).toBe(true);
    expect(token).not.toHaveBeenCalled();
    expect(result.summaries).toEqual([]);
  });

  it("WIQ-POL-10: output variable names cannot contain a trailing newline", async () => {
    const ctx = context();
    const token = vi.spyOn(ctx.m365TokenProvider, "getAccessToken");
    const result = await new PublishPolicyDriver().execute(
      inputs(),
      ctx,
      new Map([["report", "POLICY_REPORT\n"]])
    );
    expect(result.result.isErr()).toBe(true);
    if (result.result.isErr()) expect(result.result.error.name).toBe("InvalidActionInputError");
    expect(token).not.toHaveBeenCalled();
  });
});
