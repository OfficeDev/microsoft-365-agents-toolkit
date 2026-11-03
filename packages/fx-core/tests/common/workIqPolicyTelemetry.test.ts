// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ok } from "@microsoft/teamsfx-api";
import axios from "axios";
import MockAdapter from "axios-mock-adapter";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { setTools, TOOLS } from "../../src/common/globalVars";
import { getResourceServiceEndpoint, ResourceServiceType } from "../../src/common/constants";
import { TelemetryProperty, TelemetrySuccess } from "../../src/common/telemetry";
import { WrappedAxiosClient } from "../../src/common/wrappedAxiosClient";
import { PackageService } from "../../src/component/m365/packageService";
import { MockedM365Provider, MockTools } from "../core/utils";
import { MockedLogProvider } from "../plugins/solution/util";
import { WorkIqPolicyClient } from "../../src/client/workIqPolicyClient";

describe("Work IQ policy HTTP telemetry", () => {
  let previousTools: typeof TOOLS;
  let mockTools: MockTools;

  beforeEach(() => {
    previousTools = TOOLS;
    mockTools = new MockTools();
    setTools(mockTools);
  });

  afterEach(() => {
    setTools(previousTools);
    vi.restoreAllMocks();
  });

  for (const status of [201, 400]) {
    it(`WIQ-POL-10: HTTP ${status} telemetry does not capture policy, identity, token or response content`, async () => {
      const send = vi.spyOn(mockTools.telemetryReporter, "sendTelemetryEvent");
      const sendError = vi.spyOn(mockTools.telemetryReporter, "sendTelemetryErrorEvent");
      const http = WrappedAxiosClient.create();
      const adapter = new MockAdapter(http);
      const url =
        "https://regional.titles.example/builder/v1/tenants/titles/private-title/policies/workiq";
      try {
        adapter.onPut(url).reply(status, { error: "private-service-response" });
        const request = http.put(url, Buffer.from('{"rule":"private-policy-content"}'), {
          headers: { Authorization: "Bearer private-maker-token" },
        });
        if (status === 201) await expect(request).resolves.toHaveProperty("status", 201);
        else await expect(request).rejects.toHaveProperty("response.status", 400);
        expect(WrappedAxiosClient.convertUrlToApiName(url, "PUT")).toBe("mos_put_workiq_policy");
        const telemetry = JSON.stringify([send.mock.calls, sendError.mock.calls]);
        for (const secret of [
          "private-title",
          "private-policy-content",
          "private-service-response",
          "private-maker-token",
        ]) {
          expect(telemetry).not.toContain(secret);
        }
        expect(telemetry).toContain("mos_put_workiq_policy");
      } finally {
        adapter.restore();
      }
    });
  }

  for (const failure of ["response", "retry"]) {
    it(`WIQ-POL-18: shared discovery ${failure} diagnostics exclude prerequisite messages, URLs and response data`, async () => {
      const send = vi.spyOn(mockTools.telemetryReporter, "sendTelemetryEvent");
      const sendError = vi.spyOn(mockTools.telemetryReporter, "sendTelemetryErrorEvent");
      const logger = new MockedLogProvider();
      const logError = vi.spyOn(logger, "error");
      const logWarning = vi.spyOn(logger, "warning");
      const logDebug = vi.spyOn(logger, "debug");
      const provider = new MockedM365Provider();
      vi.spyOn(provider, "getAccessToken").mockResolvedValue(ok("private-maker-token"));
      const adapter = new MockAdapter(axios);
      const transport = axios.getAdapter(axios.defaults.adapter);
      axios.defaults.adapter = async (config) => ({
        ...(await transport(config)),
        request: { method: config.method?.toUpperCase() },
      });
      const origin = getResourceServiceEndpoint(ResourceServiceType.MOS3);
      let attempts = 0;
      adapter.onGet(`${origin}/config/v1/environment`).reply(async (config) => {
        attempts++;
        if (failure === "retry" && attempts > 1) {
          return [200, { titlesServiceUrl: "https://regional.example/private-region" }];
        }
        const error = new Error("private-message https://private.example/account");
        throw failure === "response"
          ? Object.assign(error, {
              config,
              response: { status: 403, data: "private-response", config },
            })
          : error;
      });
      try {
        const result = await new PackageService(origin, logger, provider).prepareWorkIqPolicy({
          titleId: "title",
          policy: { bytes: Buffer.from("{}"), sha256: "digest" },
        });
        expect(result.isOk()).toBe(failure === "retry");
        expect(adapter.history.put).toHaveLength(0);
        const diagnostics = JSON.stringify([
          logError.mock.calls,
          logWarning.mock.calls,
          logDebug.mock.calls,
          send.mock.calls,
          sendError.mock.calls,
        ]);
        expect(diagnostics).not.toContain("private-");
        expect(diagnostics).not.toContain("https://private.example");
      } finally {
        adapter.restore();
      }
    });
  }

  for (const status of [201, 422, 503]) {
    it(`WIQ-POL-22: Work IQ ${status} telemetry never exposes runtime app binding, title or policy data`, async () => {
      const send = vi.spyOn(mockTools.telemetryReporter, "sendTelemetryEvent");
      const sendError = vi.spyOn(mockTools.telemetryReporter, "sendTelemetryErrorEvent");
      const http = WrappedAxiosClient.create();
      const adapter = new MockAdapter(http);
      const endpoint = "https://workiq.example/private-environment";
      const callerId = "11111111-2222-4333-8444-555555555555";
      const client = new WorkIqPolicyClient(
        endpoint,
        ["api://workiq-test/WorkIQSettings.ReadWrite.All"],
        http
      );
      adapter
        .onGet(`${endpoint}/api/policies/contracts`)
        .reply(200, { opaque: "private-contract-response" });
      const provider = new MockedM365Provider();
      vi.spyOn(provider, "getAccessToken").mockResolvedValue(ok("private-runtime-token"));
      adapter
        .onPut(`${endpoint}/api/policies/agents/private-title`)
        .reply(status, { error: "private-service-diagnostics" }, { ETag: '"runtime-etag"' });
      try {
        await client.publish(
          {
            titleId: "private-title",
            callerId,
            policy: {
              bytes: Buffer.from('{"schemaVersion":1,"capabilities":[],"private-policy":true}'),
              sha256: "hash",
            },
          },
          provider
        );
        const telemetry = JSON.stringify([send.mock.calls, sendError.mock.calls]);
        expect(telemetry).toContain("workiq_put_agent_policy");
        for (const value of [
          "private-title",
          "private-policy",
          "private-runtime-token",
          "private-service-diagnostics",
          "private-environment",
          callerId,
        ]) {
          expect(telemetry).not.toContain(value);
        }
      } finally {
        adapter.restore();
      }
    });
  }

  for (const prefix of ["", "/private-region"]) {
    for (const status of [201, 400, 503]) {
      it(`WIQ-POL-10: PackageService HTTP ${status} with discovery prefix '${prefix}' records safe, correctly classified telemetry`, async () => {
        const send = vi.spyOn(mockTools.telemetryReporter, "sendTelemetryEvent");
        const sendError = vi.spyOn(mockTools.telemetryReporter, "sendTelemetryErrorEvent");
        const adapter = new MockAdapter(axios);
        const transport = axios.getAdapter(axios.defaults.adapter);
        axios.defaults.adapter = async (config) => ({
          ...(await transport(config)),
          request: { method: config.method?.toUpperCase() },
        });
        const origin = getResourceServiceEndpoint(ResourceServiceType.MOS3);
        const serviceUrl = `https://regional.titles.example${prefix}`;
        const url = `${serviceUrl}/builder/v1/tenants/titles/private-title/policies/workiq`;
        try {
          adapter.onGet(`${origin}/config/v1/environment`).reply(200, {
            titlesServiceUrl: serviceUrl,
          });
          adapter.onPut(url).reply(status, { error: "private-service-response" });
          const provider = new MockedM365Provider();
          vi.spyOn(provider, "getAccessToken").mockResolvedValue(ok("private-maker-token"));
          const result = await new PackageService(origin, undefined, provider).publishWorkIqPolicy({
            titleId: "private-title",
            policy: { bytes: Buffer.from('{"rule":"private-policy-content"}'), sha256: "digest" },
          });
          expect(result.isOk()).toBe(status === 201);
          const expected = {
            url: "<mos_put_workiq_policy-url>",
            [TelemetryProperty.Success]:
              status === 201 ? TelemetrySuccess.Yes : TelemetrySuccess.No,
            "status-code": String(status),
          };
          if (status === 201) {
            expect(send).toHaveBeenCalledWith(
              expect.any(String),
              expect.objectContaining(expected)
            );
          } else {
            expect(sendError).toHaveBeenCalledWith(
              expect.any(String),
              expect.objectContaining(expected)
            );
          }
          const telemetry = JSON.stringify([send.mock.calls, sendError.mock.calls]);
          for (const secret of [
            "private-title",
            "private-region",
            "private-policy-content",
            "private-service-response",
            "private-maker-token",
          ]) {
            expect(telemetry).not.toContain(secret);
          }
          expect(adapter.history.put).toHaveLength(1);
        } finally {
          adapter.restore();
        }
      });
    }
  }
});
