// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError, err, ok } from "@microsoft/teamsfx-api";
import { createHash } from "node:crypto";
import { mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { ActionNames } from "../../src/actions";
import { workIqPolicyClient } from "../../src/client/workIqPolicyClient";
import { WorkIqServiceScope } from "../../src/common/constants";
import { PublishPolicyDriver } from "../../src/component/driver/workiq/publishPolicy";
import type { ExecutionResult } from "../../src/component/driver/interface/stepDriver";
import { PackageService } from "../../src/component/m365/packageService";
import type { PreparedPolicyPublication } from "../../src/component/workiq/policyPublication";
import { FxCoreClient } from "../../src/core/FxCoreClient";
import { MockTools } from "./utils";

const callerId = "11111111-2222-4333-8444-555555555555";
const policyBytes = Buffer.from('{\r\n  "rule": "${{DO_NOT_EXPAND}}", "allow": []\r\n}\r\n');

function fakeDestinations() {
  const mosPublish = vi
    .fn<PreparedPolicyPublication["publish"]>()
    .mockResolvedValue(ok({ state: "confirmed", etag: '"mos-committed"' }));
  const workIqPublish = vi
    .fn<PreparedPolicyPublication["publish"]>()
    .mockResolvedValue(ok({ state: "confirmed", etag: '"workiq-committed"' }));
  const mosPrepare = vi
    .spyOn(PackageService.prototype, "prepareWorkIqPolicy")
    .mockResolvedValue(ok({ publish: mosPublish }));
  const workIqPrepare = vi
    .spyOn(workIqPolicyClient, "prepare")
    .mockResolvedValue(ok({ publish: workIqPublish }));
  return { mosPublish, workIqPublish, mosPrepare, workIqPrepare };
}

describe("public policy action", () => {
  let projectPath: string;
  let tools: MockTools;
  let client: FxCoreClient;

  beforeEach(async () => {
    vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "both");
    projectPath = await mkdtemp(join(tmpdir(), "atk-public-policy-"));
    await writeFile(join(projectPath, "policy.json"), policyBytes);
    tools = new MockTools();
    client = new FxCoreClient(tools);
  });

  afterEach(async () => {
    vi.restoreAllMocks();
    vi.unstubAllEnvs();
    await rm(projectPath, { recursive: true, force: true });
  });

  const args = () => ({
    titleId: "U_title",
    callerId,
    policyPath: join(projectPath, "policy.json"),
  });

  for (const value of [undefined, "mos", "workiq", "both"]) {
    it(`POLICY-ACTION-05: ${value ?? "unset"} preserves the native target selection`, async () => {
      vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", value);
      const target = value ?? "workiq";
      const destinations = fakeDestinations();

      const execution = await client.executeAction(ActionNames.PublishPolicy, args(), {
        projectPath,
      });

      expect(execution.result.isOk()).toBe(true);
      if (!("report" in execution)) throw new Error("Missing native policy report");
      expect(execution.report).toMatchObject({
        target,
        outcome: "published",
        nextAction: "none",
        mos: { state: target === "workiq" ? "not-started" : "confirmed" },
        workIq: { state: target === "mos" ? "not-started" : "confirmed" },
      });
      expect(destinations.mosPrepare).toHaveBeenCalledTimes(target === "workiq" ? 0 : 1);
      expect(destinations.mosPublish).toHaveBeenCalledTimes(target === "workiq" ? 0 : 1);
      expect(destinations.workIqPrepare).toHaveBeenCalledTimes(target === "mos" ? 0 : 1);
      expect(destinations.workIqPublish).toHaveBeenCalledTimes(target === "mos" ? 0 : 1);
    });
  }

  it("POLICY-ACTION-02: returns native receipts after one captured snapshot and MOS-first publication", async () => {
    const destinations = fakeDestinations();
    const token = vi.spyOn(tools.tokenProvider.m365TokenProvider, "getAccessToken");
    const driver = vi.spyOn(PublishPolicyDriver.prototype, "execute");
    const signal = new AbortController().signal;
    destinations.workIqPrepare.mockImplementation(async (_request, provider) => {
      if (!provider) throw new Error("Missing injected token provider");
      const auth = await provider.getAccessToken({ scopes: WorkIqServiceScope() });
      if (auth.isErr()) return err(auth.error);
      await writeFile(join(projectPath, "policy.json"), '{"edited":true}');
      return ok({ publish: destinations.workIqPublish });
    });

    const execution = await client.executeAction(ActionNames.PublishPolicy, args(), {
      projectPath,
      signal,
    });
    expect(execution.result.isOk()).toBe(true);
    expect("report" in execution).toBe(true);
    if (!("report" in execution)) throw new Error("Missing native policy report");
    expect(execution.report).toMatchObject({
      reportVersion: 1,
      outcome: "published",
      enforcement: "not-verified",
      policySha256: createHash("sha256").update(policyBytes).digest("hex"),
      mos: { state: "confirmed", etag: '"mos-committed"' },
      workIq: { state: "confirmed", etag: '"workiq-committed"' },
    });
    if (execution.result.isOk()) expect(execution.result.value.size).toBe(0);
    expect(token).toHaveBeenCalledWith({ scopes: WorkIqServiceScope() });
    expect(driver).toHaveBeenCalledWith(
      args(),
      expect.objectContaining({ projectPath, signal, nonInteractive: true, ui: undefined })
    );
    expect(destinations.mosPrepare.mock.calls[0][0].policy.bytes).toEqual(policyBytes);
    expect(destinations.workIqPrepare.mock.calls[0][0]).toMatchObject({
      titleId: "U_title",
      callerId,
      signal,
    });
    expect(destinations.mosPrepare.mock.invocationCallOrder[0]).toBeLessThan(
      destinations.mosPublish.mock.invocationCallOrder[0]
    );
    expect(destinations.workIqPrepare.mock.invocationCallOrder[0]).toBeLessThan(
      destinations.mosPublish.mock.invocationCallOrder[0]
    );
    expect(destinations.mosPublish.mock.invocationCallOrder[0]).toBeLessThan(
      destinations.workIqPublish.mock.invocationCallOrder[0]
    );
    expect(destinations.mosPublish).toHaveBeenCalledOnce();
    expect(destinations.workIqPublish).toHaveBeenCalledOnce();
    expect(await readdir(projectPath)).toEqual(["policy.json"]);
    expect(await readFile(join(projectPath, "policy.json"), "utf8")).toBe('{"edited":true}');
  });

  for (const showDialog of [undefined, true, false]) {
    it(`POLICY-ACTION-06: preserves the provider and showDialog=${showDialog} authentication request`, async () => {
      vi.stubEnv("TEAMSFX_WORKIQ_PUBLISH_TARGET", "workiq");
      const destinations = fakeDestinations();
      const provider = tools.tokenProvider.m365TokenProvider;
      const token = vi.spyOn(provider, "getAccessToken").mockResolvedValue(ok("provider-token"));
      const request = {
        scopes: WorkIqServiceScope(),
        ...(showDialog !== undefined ? { showDialog } : {}),
      };
      destinations.workIqPrepare.mockImplementation(async (_request, tokenProvider) => {
        if (!tokenProvider) throw new Error("Missing injected token provider");
        const auth = await tokenProvider.getAccessToken(request);
        if (auth.isErr()) return err(auth.error);
        return ok({ publish: destinations.workIqPublish });
      });

      const execution = await client.executeAction(ActionNames.PublishPolicy, args(), {
        projectPath,
      });

      expect(execution.result.isOk()).toBe(true);
      expect(destinations.workIqPrepare.mock.calls[0][1]).toBe(provider);
      expect(token).toHaveBeenCalledExactlyOnceWith(request);
      expect(destinations.workIqPublish).toHaveBeenCalledOnce();
      expect(destinations.mosPrepare).not.toHaveBeenCalled();
      if (!("report" in execution)) throw new Error("Missing native policy report");
      expect(execution.report).toMatchObject({ target: "workiq", outcome: "published" });
    });
  }

  it("POLICY-ACTION-03: preserves partial receipts and domain errors without retry", async () => {
    const destinations = fakeDestinations();
    destinations.workIqPublish.mockResolvedValue(
      err({
        state: "failed",
        error: new UserError({ source: "test", name: "RejectedPolicy", message: "rejected" }),
      })
    );
    const execution = await client.executeAction(ActionNames.PublishPolicy, args(), {
      projectPath,
    });
    expect(execution.result.isErr()).toBe(true);
    if (!("report" in execution)) throw new Error("Missing failure report");
    expect(execution.report).toMatchObject({
      outcome: "incomplete",
      nextAction: "complete-publication",
      mos: { state: "confirmed" },
      workIq: { state: "failed" },
    });
    if (execution.result.isErr())
      expect(execution.result.error).toHaveProperty("report", execution.report);
    expect(destinations.mosPublish).toHaveBeenCalledOnce();
    expect(destinations.workIqPublish).toHaveBeenCalledOnce();
  });

  it("POLICY-ACTION-03: awaits the native driver and preserves confirmed MOS receipts on cancellation", async () => {
    const destinations = fakeDestinations();
    const controller = new AbortController();
    destinations.mosPublish.mockImplementation(async () => {
      controller.abort();
      return ok({ state: "confirmed", etag: '"mos-committed"' });
    });
    const execution = await client.executeAction(ActionNames.PublishPolicy, args(), {
      projectPath,
      signal: controller.signal,
    });
    expect(execution.result.isErr()).toBe(true);
    if (!("report" in execution)) throw new Error("Missing cancellation report");
    expect(execution.report).toMatchObject({
      outcome: "cancelled",
      nextAction: "complete-publication",
      mos: { state: "confirmed", etag: '"mos-committed"' },
      workIq: { state: "not-started" },
    });
    if (execution.result.isErr()) {
      expect(execution.result.error.name).toBe("UserCancel");
      expect(execution.result.error).toHaveProperty("report", execution.report);
    }
    expect(destinations.workIqPublish).not.toHaveBeenCalled();
  });

  it("POLICY-ACTION-04: keeps the native pre-aborted report without authentication or preparation", async () => {
    const destinations = fakeDestinations();
    const controller = new AbortController();
    controller.abort();
    const token = vi.spyOn(tools.tokenProvider.m365TokenProvider, "getAccessToken");
    const execution = await client.executeAction(ActionNames.PublishPolicy, args(), {
      projectPath,
      signal: controller.signal,
    });
    expect(execution.result.isErr()).toBe(true);
    if (!("report" in execution)) throw new Error("Missing pre-aborted report");
    expect(execution.report).toMatchObject({
      outcome: "cancelled",
      mos: { state: "not-started" },
      workIq: { state: "not-started" },
    });
    expect(token).not.toHaveBeenCalled();
    expect(destinations.mosPrepare).not.toHaveBeenCalled();
    expect(destinations.workIqPrepare).not.toHaveBeenCalled();
  });

  it("POLICY-ACTION-04: rejects invalid policy args before entering the driver", async () => {
    const driver = vi.spyOn(PublishPolicyDriver.prototype, "execute");
    const token = vi.spyOn(tools.tokenProvider.m365TokenProvider, "getAccessToken");
    const execution: ExecutionResult = await Reflect.apply(client.executeAction, client, [
      ActionNames.PublishPolicy,
      { titleId: "U_title", policyPath: "policy.json" },
      { projectPath },
    ]);
    expect(execution.result.isErr()).toBe(true);
    if (execution.result.isErr())
      expect(execution.result.error.name).toBe("ActionArgsValidationError");
    expect(driver).not.toHaveBeenCalled();
    expect(token).not.toHaveBeenCalled();
  });
});
