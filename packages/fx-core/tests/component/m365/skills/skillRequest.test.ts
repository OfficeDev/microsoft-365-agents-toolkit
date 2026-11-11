// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AxiosError, AxiosHeaders } from "axios";
import { SystemError } from "@microsoft/teamsfx-api";
import { expect, vi } from "vitest";
import { requestSkillOperation } from "../../../../src/component/m365/skills/skillRequest";

function httpFailure(status: number, retryAfter?: string): AxiosError {
  const config = { headers: new AxiosHeaders() };
  return new AxiosError("sensitive service message", undefined, config, undefined, {
    status,
    statusText: "",
    config,
    headers: retryAfter === undefined ? {} : { "retry-after": retryAfter },
    data: { downloadUrl: "https://private.example/?secret=never-log" },
  });
}

describe("skill requests", () => {
  afterEach(() => {
    vi.useRealTimers();
    vi.restoreAllMocks();
  });

  it("AC-08: honors Retry-After before retrying a throttled operation", async () => {
    vi.useFakeTimers();
    const request = vi.fn().mockRejectedValueOnce(httpFailure(429, "3")).mockResolvedValue("ok");
    const pending = requestSkillOperation(request);
    await vi.advanceTimersByTimeAsync(2999);
    expect(request).toHaveBeenCalledTimes(1);
    await vi.advanceTimersByTimeAsync(1);
    const result = await pending;
    expect(result.isOk() && result.value).toBe("ok");
    expect(request).toHaveBeenCalledTimes(2);
  });

  it("AC-08: honors an HTTP-date Retry-After", async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date("2026-10-10T00:00:00Z"));
    const request = vi
      .fn()
      .mockRejectedValueOnce(httpFailure(429, "Sat, 10 Oct 2026 00:00:04 GMT"))
      .mockResolvedValue("ok");
    const pending = requestSkillOperation(request);
    await vi.advanceTimersByTimeAsync(3999);
    expect(request).toHaveBeenCalledTimes(1);
    await vi.advanceTimersByTimeAsync(1);
    expect((await pending).isOk()).toBe(true);
  });

  it("AC-08: makes at most six attempts with exponential fallback", async () => {
    vi.useFakeTimers();
    vi.spyOn(Math, "random").mockReturnValue(0);
    const request = vi.fn().mockRejectedValue(httpFailure(503));
    const pending = requestSkillOperation(request);
    await vi.runAllTimersAsync();
    expect((await pending).isErr()).toBe(true);
    expect(request).toHaveBeenCalledTimes(6);
  });

  it("AC-08: does not shorten a server wait beyond the retry deadline", async () => {
    const request = vi.fn().mockRejectedValue(httpFailure(429, "3600"));
    const result = await requestSkillOperation(request);
    expect(result.isErr() && result.error.name).toBe("SkillUploadRetryExhausted");
    expect(request).toHaveBeenCalledTimes(1);
  });

  it.each([400, 401, 403, 404, 409, 412])(
    "AC-10/AC-17: does not retry permanent HTTP %s",
    async (status) => {
      const request = vi.fn().mockRejectedValue(httpFailure(status));
      const result = await requestSkillOperation(request);
      expect(result.isErr()).toBe(true);
      expect(request).toHaveBeenCalledTimes(1);
    }
  );

  it("AC-16: does not expose response bodies, URLs or bearer configuration", async () => {
    const result = await requestSkillOperation(() => Promise.reject(httpFailure(403)));
    expect(result.isErr()).toBe(true);
    expect(JSON.stringify(result)).not.toContain("secret");
    expect(JSON.stringify(result)).not.toContain("private.example");
    expect(JSON.stringify(result)).not.toContain("sensitive service message");
  });

  it("AC-10: unknown failures are explicit and not retried", async () => {
    const request = vi.fn().mockRejectedValue(new Error("untrusted error detail"));
    const result = await requestSkillOperation(request);
    expect(result.isErr()).toBe(true);
    expect(request).toHaveBeenCalledTimes(1);
  });

  it("AC-08: retries replay-safe transport failures only", async () => {
    vi.useFakeTimers();
    const request = vi
      .fn()
      .mockRejectedValueOnce(new AxiosError("TLS details", "ECONNRESET"))
      .mockResolvedValue("ok");
    const pending = requestSkillOperation(request);
    await vi.runAllTimersAsync();
    expect((await pending).isOk()).toBe(true);
    expect(request).toHaveBeenCalledTimes(2);
  });

  it.each([new AxiosError("lost response", "ECONNRESET"), httpFailure(503)])(
    "AC-09: does not replay an ambiguously completed mutation",
    async (failure) => {
      const request = vi.fn().mockRejectedValue(failure);
      expect((await requestSkillOperation(request, false)).isErr()).toBe(true);
      expect(request).toHaveBeenCalledTimes(1);
    }
  );
  it.each([429, 503])(
    "AC-09: never mode does not replay directory or package POST after HTTP %s",
    async (status) => {
      vi.useFakeTimers();
      const request = vi.fn().mockRejectedValue(httpFailure(status, "1"));
      const pending = requestSkillOperation(request, "never");
      await vi.runAllTimersAsync();
      const result = await pending;
      expect(result.isErr() && result.error.name).toBe("SkillUploadRequestFailed");
      expect(result.isErr() && result.error).toBeInstanceOf(SystemError);
      expect(request).toHaveBeenCalledTimes(1);
      expect(JSON.stringify(result)).not.toContain("sensitive service message");
    }
  );
});
