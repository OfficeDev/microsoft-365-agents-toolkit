// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "reflect-metadata";
import { err, M365TokenProvider, ok, Platform, UserError } from "@microsoft/teamsfx-api";
import axios from "axios";
import MockAdapter from "axios-mock-adapter";
import { mkdtemp, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { WorkIqPolicyClient, workIqPolicyClient } from "../../../../src/client/workIqPolicyClient";
import { getResourceServiceEndpoint, ResourceServiceType } from "../../../../src/common/constants";
import { DriverContext } from "../../../../src/component/driver/interface/commonArgs";
import { PublishPolicyDriver } from "../../../../src/component/driver/workiq/publishPolicy";
import { PackageService } from "../../../../src/component/m365/packageService";
import { PolicyPublicationOperation } from "../../../../src/component/workiq/policyPublication";
import { UserCancelError } from "../../../../src/error/common";
import { MockedAzureAccountProvider, MockedM365Provider } from "../../../core/utils";
import { MockedLogProvider } from "../../../plugins/solution/util";

const flag = "TEAMSFX_WORKIQ_PUBLISH_TARGET";
const mosEndpoint = getResourceServiceEndpoint(ResourceServiceType.MOS3);
const discoveryUrl = `${mosEndpoint}/config/v1/environment`;
const mosServiceUrl = "https://regional.titles.example";
const workIqEndpoint = "https://workiq.example";
const workIqScope = "api://workiq-test/WorkIQSettings.ReadWrite.All";
const titleId = "test-title";
const callerId = "11111111-2222-4333-8444-555555555555";
const mosUrl = `${mosServiceUrl}/builder/v1/tenants/titles/${titleId}/policies/workiq`;
const workIqUrl = `${workIqEndpoint}/api/policies/agents/${titleId}`;
const contractUrl = `${workIqEndpoint}/api/policies/contracts`;
const inputs = { titleId, callerId, policyPath: "./policy.json" };

describe("policy publication targets", () => {
  let projectPath: string;
  let adapter: MockAdapter;
  let client: WorkIqPolicyClient;
  let provider: MockedM365Provider;
  const getToken = vi.fn<M365TokenProvider["getAccessToken"]>();
  const execute = (signal?: AbortSignal, args: unknown = inputs) =>
    new PolicyPublicationOperation({
      mos: {
        prepare: (request) =>
          new PackageService(mosEndpoint, undefined, provider).prepareWorkIqPolicy(request),
      },
      workIq: {
        prepare: (request) => client.prepare(request, provider),
      },
    }).execute(args, { projectPath, signal });
  const writes = (url: string) =>
    adapter.history.put.filter((request) => axios.getUri(request) === url);
  const context = (): DriverContext => ({
    projectPath,
    platform: Platform.CLI,
    azureAccountProvider: new MockedAzureAccountProvider(),
    m365TokenProvider: provider,
    logProvider: new MockedLogProvider(),
    ui: undefined,
    progressBar: undefined,
  });

  beforeEach(async () => {
    vi.stubEnv(flag, "both");
    vi.stubEnv("TEAMSFX_WORKIQ_ALLOW_MOS_FAILURE", undefined);
    vi.stubEnv("TEAMSFX_WORKIQ_MOS_ENV", "prod");
    projectPath = await mkdtemp(path.join(os.tmpdir(), "atk-workiq-target-"));
    await writeFile(path.join(projectPath, "policy.json"), '{"schemaVersion":1,"capabilities":[]}');
    adapter = new MockAdapter(axios);
    adapter.onGet(contractUrl).reply(200, {});
    adapter.onGet(workIqUrl).reply(404);
    adapter.onGet(discoveryUrl).reply(200, { titlesServiceUrl: mosServiceUrl });
    adapter.onGet(mosUrl).reply(404);
    adapter.onPut(mosUrl).reply(201, { policyType: "workiq", policy: {} }, { ETag: '"mos-saved"' });
    adapter.onPut(workIqUrl).reply(201, {}, { ETag: '"workiq-saved"' });
    provider = new MockedM365Provider();
    getToken.mockReset();
    getToken.mockImplementation(async (request) =>
      ok(request.scopes.includes(workIqScope) ? "workiq-token" : "mos-token")
    );
    provider.getAccessToken = getToken;
    client = new WorkIqPolicyClient(workIqEndpoint, [workIqScope]);
    vi.spyOn(workIqPolicyClient, "prepare").mockImplementation((request, tokenProvider) =>
      client.prepare(request, tokenProvider)
    );
  });

  afterEach(async () => {
    adapter.restore();
    vi.restoreAllMocks();
    vi.unstubAllEnvs();
    await rm(projectPath, { recursive: true, force: true });
  });

  for (const value of [undefined, "mos", "workiq", "both"]) {
    it(`WIQ-POL-34/WIQ-POL-35/WIQ-POL-37: ${value ?? "unset"} contacts exactly its selected targets`, async () => {
      vi.stubEnv(flag, value);
      const target = value ?? "workiq";
      const result = await execute();

      expect(result.isOk()).toBe(true);
      if (result.isOk())
        expect(result.value).toMatchObject({
          target,
          outcome: "published",
          nextAction: "none",
          mos: { state: target === "workiq" ? "not-started" : "confirmed" },
          workIq: { state: target === "mos" ? "not-started" : "confirmed" },
        });
      expect(getToken.mock.calls.map(([request]) => request.scopes)).toEqual([
        ...(target !== "mos" ? [[workIqScope]] : []),
        ...(target !== "workiq" ? [[`${mosEndpoint}/.default`]] : []),
      ]);
      expect(adapter.history.get.map((request) => axios.getUri(request))).toEqual([
        ...(target !== "mos" ? [workIqUrl] : []),
        ...(target !== "workiq" ? [discoveryUrl, mosUrl] : []),
      ]);
      expect(adapter.history.put.map((request) => axios.getUri(request))).toEqual([
        ...(target !== "workiq" ? [mosUrl] : []),
        ...(target !== "mos" ? [workIqUrl] : []),
      ]);
    });
  }

  it("WIQ-POL-34: MOS-only ignores unavailable Work IQ prerequisites", async () => {
    vi.stubEnv(flag, "mos");
    const prepare = vi.spyOn(client, "prepare").mockRejectedValue(new Error("private-error"));
    adapter.onGet(contractUrl).reply(403);
    const result = await execute();
    expect(result.isOk()).toBe(true);
    expect(prepare).not.toHaveBeenCalled();
    expect(writes(mosUrl)).toHaveLength(1);
    expect(writes(workIqUrl)).toHaveLength(0);
  });

  it("WIQ-POL-35: Work IQ-only ignores invalid MOS configuration and unavailable preparation", async () => {
    vi.stubEnv(flag, "workiq");
    vi.stubEnv("TEAMSFX_WORKIQ_MOS_ENV", "invalid");
    const prepare = vi
      .spyOn(PackageService.prototype, "prepareWorkIqPolicy")
      .mockRejectedValue(new Error("private-error"));
    const result = await execute();
    expect(result.isOk()).toBe(true);
    expect(prepare).not.toHaveBeenCalled();
    expect(writes(mosUrl)).toHaveLength(0);
    expect(writes(workIqUrl)).toHaveLength(1);
  });

  for (const target of ["mos", "workiq"]) {
    for (const { status, state, outcome, nextAction } of [
      { status: 400, state: "failed", outcome: "failed", nextAction: "review-error" },
      { status: 412, state: "conflict", outcome: "conflict", nextAction: "resolve-conflict" },
      { status: 503, state: "unknown", outcome: "unknown", nextAction: "reconcile" },
    ]) {
      it(`WIQ-POL-34/WIQ-POL-35: ${target}-only preserves selected ${state} without a false partial outcome`, async () => {
        vi.stubEnv(flag, target);
        const selectedUrl = target === "mos" ? mosUrl : workIqUrl;
        adapter.onPut(selectedUrl).reply(status, { detail: "private-error" });
        const result = await execute();
        expect(result.isErr()).toBe(true);
        if (result.isErr()) {
          expect(result.error.report).toMatchObject({
            target,
            outcome,
            nextAction,
            [target === "mos" ? "mos" : "workIq"]: { state },
            [target === "mos" ? "workIq" : "mos"]: { state: "not-started" },
          });
          expect(JSON.stringify(result.error)).not.toContain("private-error");
        }
        expect(adapter.history.put).toHaveLength(1);
      });
    }
  }

  for (const stage of ["workiq-get", "mos-auth", "mos-discovery", "mos-get", "mos-throw"]) {
    it(`WIQ-POL-36: both mode stops before all writes on ${stage} preparation failure`, async () => {
      vi.stubEnv(flag, "both");
      if (stage === "workiq-get") adapter.onGet(workIqUrl).reply(403);
      else if (stage === "mos-auth")
        getToken.mockImplementation(async (request) =>
          request.scopes.includes(workIqScope)
            ? ok("workiq-token")
            : err(new UserError({ source: "auth", name: "Denied", message: "private-error" }))
        );
      else if (stage === "mos-discovery") adapter.onGet(discoveryUrl).reply(403);
      else if (stage === "mos-get") adapter.onGet(mosUrl).reply(403);
      else
        vi.spyOn(PackageService.prototype, "prepareWorkIqPolicy").mockRejectedValue(
          new Error("private-error")
        );
      const result = await execute();
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error.report).toMatchObject({
          target: "both",
          mos: { state: "not-started" },
          workIq: { state: "not-started" },
        });
        expect(JSON.stringify(result.error)).not.toContain("private-error");
      }
      expect(adapter.history.put).toHaveLength(0);
    });
  }

  for (const status of [400, 412, 503]) {
    it(`WIQ-POL-36: both mode never continues after MOS PUT ${status}`, async () => {
      vi.stubEnv(flag, "both");
      adapter.onPut(mosUrl).reply(status);
      expect((await execute()).isErr()).toBe(true);
      expect(writes(mosUrl)).toHaveLength(1);
      expect(writes(workIqUrl)).toHaveLength(0);
    });
  }

  for (const value of ["", "yes", "no", "true", "MOS", "both ", "unknown"]) {
    it(`WIQ-POL-37: invalid target ${JSON.stringify(value)} fails before auth or HTTP`, async () => {
      vi.stubEnv(flag, value);
      const result = await execute();
      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error.name).toBe("POLICY_PUBLISH_TARGET_INVALID");
        expect(result.error.message).toContain(flag);
        expect(result.error.message).toContain("Leave it unset to publish to Work IQ only.");
      }
      expect(getToken).not.toHaveBeenCalled();
      expect(adapter.history.get).toHaveLength(0);
      expect(adapter.history.put).toHaveLength(0);
    });
  }

  for (const value of [undefined, "mos", "workiq", "both"]) {
    for (const stage of ["pre-aborted", "invalid-input"]) {
      it(`WIQ-POL-37: ${value ?? "unset"} target survives ${stage} before preparation`, async () => {
        vi.stubEnv(flag, value);
        const controller = new AbortController();
        if (stage === "pre-aborted") controller.abort();

        const result = await execute(controller.signal, stage === "invalid-input" ? {} : inputs);

        expect(result.isErr()).toBe(true);
        if (result.isErr())
          expect(result.error.report).toMatchObject({
            target: value ?? "workiq",
            outcome: stage === "pre-aborted" ? "cancelled" : "failed",
            mos: { state: "not-started" },
            workIq: { state: "not-started" },
          });
        expect(getToken).not.toHaveBeenCalled();
        expect(adapter.history.get).toHaveLength(0);
        expect(adapter.history.put).toHaveLength(0);
      });
    }
  }

  it("WIQ-POL-37: the former yes/no switch cannot enable continuation", async () => {
    vi.stubEnv("TEAMSFX_WORKIQ_ALLOW_MOS_FAILURE", "yes");
    adapter.onPut(mosUrl).reply(400);
    expect((await execute()).isErr()).toBe(true);
    expect(writes(mosUrl)).toHaveLength(1);
    expect(writes(workIqUrl)).toHaveLength(0);
  });

  for (const value of [undefined, "mos", "workiq", "both"]) {
    it(`WIQ-POL-37: ${value ?? "unset"} selection is captured before authentication`, async () => {
      vi.stubEnv(flag, value);
      const target = value ?? "workiq";
      getToken.mockImplementation(async () => {
        vi.stubEnv(flag, target === "mos" ? "workiq" : "mos");
        return ok("token");
      });
      const result = await execute();
      expect(result.isOk()).toBe(true);
      if (result.isOk()) expect(result.value).toHaveProperty("target", target);
      expect(writes(mosUrl)).toHaveLength(target === "workiq" ? 0 : 1);
      expect(writes(workIqUrl)).toHaveLength(target === "mos" ? 0 : 1);
    });
  }

  for (const target of ["mos", "workiq"]) {
    for (const stage of ["prepare", "write"]) {
      for (const thrown of [false, true]) {
        it(`WIQ-POL-36: ${target}-only retains ${thrown ? "thrown" : "returned"} ${stage} cancellation`, async () => {
          vi.stubEnv(flag, target);
          const prepare =
            target === "mos"
              ? vi.spyOn(PackageService.prototype, "prepareWorkIqPolicy")
              : vi.spyOn(client, "prepare");
          prepare.mockImplementation(async () => {
            const cancellation = new UserCancelError();
            if (stage === "prepare") {
              if (thrown) throw cancellation;
              return err(cancellation);
            }
            return ok({
              publish: async () => {
                if (thrown) throw cancellation;
                return err({ state: "unknown", error: cancellation });
              },
            });
          });
          const result = await execute();
          expect(result.isErr()).toBe(true);
          if (result.isErr()) {
            expect(result.error.report).toMatchObject({
              target,
              outcome: "cancelled",
              nextAction: stage === "write" ? "reconcile" : "none",
            });
          }
          expect(adapter.history.put).toHaveLength(0);
        });
      }
    }
  }

  for (const method of ["get", "put"]) {
    it(`WIQ-POL-36: MOS HTTP ${method} cancellation without an aborted signal stops both mode`, async () => {
      const cancel = () => {
        throw new UserCancelError();
      };
      if (method === "get") adapter.onGet(mosUrl).reply(cancel);
      else adapter.onPut(mosUrl).reply(cancel);
      const result = await execute();
      expect(result.isErr()).toBe(true);
      if (result.isErr())
        expect(result.error.report).toMatchObject({
          outcome: "cancelled",
          mos: { state: method === "get" ? "not-started" : "unknown" },
          workIq: { state: "not-started" },
        });
      expect(writes(workIqUrl)).toHaveLength(0);
    });
  }

  for (const value of [undefined, "mos", "workiq", "both"]) {
    it(`WIQ-POL-38: native ${value ?? "unset"} success persists the target-aware report`, async () => {
      vi.stubEnv(flag, value);
      const target = value ?? "workiq";
      const result = await new PublishPolicyDriver().execute(
        inputs,
        context(),
        new Map([["report", "POLICY_REPORT"]])
      );
      expect(result.result.isOk()).toBe(true);
      expect(result.report).toMatchObject({ target, outcome: "published", nextAction: "none" });
      if (result.result.isOk())
        expect(result.result.value.get("POLICY_REPORT")).toBe(JSON.stringify(result.report));
      expect(result.summaries.join(" ")).toContain(`target=${target}`);
    });
  }

  for (const target of ["mos", "workiq"]) {
    it(`WIQ-POL-38: native ${target} failure exposes receipts instead of success outputs`, async () => {
      vi.stubEnv(flag, target);
      adapter.onPut(target === "mos" ? mosUrl : workIqUrl).reply(503);
      const result = await new PublishPolicyDriver().execute(
        inputs,
        context(),
        new Map([["report", "POLICY_REPORT"]])
      );
      expect(result.result.isErr()).toBe(true);
      expect(result.report).toMatchObject({ target, outcome: "unknown", nextAction: "reconcile" });
      if (result.result.isErr())
        expect(result.result.error).toHaveProperty("report", result.report);
      expect(result.summaries).toEqual([]);
      expect(adapter.history.put).toHaveLength(1);
    });
  }

  for (const target of ["mos", "workiq"]) {
    it(`WIQ-POL-38: ${target} emits a mock warning only for an actually selected Work IQ mock`, async () => {
      vi.stubEnv(flag, target);
      client = new WorkIqPolicyClient();
      const ctx = context();
      const warning = vi.spyOn(ctx.logProvider, "warning");
      const result = await new PublishPolicyDriver().execute(inputs, ctx);
      expect(result.result.isOk()).toBe(true);
      expect(result.report).toMatchObject({
        target,
        outcome: target === "mos" ? "published" : "simulated",
        nextAction: target === "mos" ? "none" : "verify-live-publication",
      });
      expect(warning).toHaveBeenCalledTimes(target === "mos" ? 0 : 1);
      expect(adapter.history.get).toHaveLength(target === "mos" ? 2 : 0);
    });
  }
});
