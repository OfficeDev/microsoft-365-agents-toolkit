// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { UserError } from "@microsoft/teamsfx-api";
import axios, { AxiosAdapter, AxiosError, AxiosHeaders } from "axios";
import { expect, vi } from "vitest";
import * as localizeUtils from "../../../../src/common/localizeUtils";
import { SkillDirectoryClient } from "../../../../src/component/m365/skills/skillDirectoryClient";

const folder = "11111111-1111-1111-1111-111111111111";
const siteUrl = "https://tenant.sharepoint.com/sites/skills";
const monitor = "https://tenant.sharepoint.com/_api/v2.1/monitor/copy?tempauth=secret";
const response = () => ({
  folder,
  sharePointIds: {
    siteId: "2c2c2c2c-1a1a-4b4b-8c8c-9d9d9d9d9d9d",
    webId: "3f3f3f3f-2b2b-4c4c-8d8d-0e0e0e0e0e0e",
    listId: "4a4a4a4a-3c3c-4d4d-8e8e-1f1f1f1f1f1f",
    uniqueId: "{11111111-1111-1111-1111-111111111111}",
    siteUrl,
    tenantId: folder,
  },
  copyStatusUrls: null,
});

function httpError(status: number): AxiosError {
  return new AxiosError("secret response", undefined, undefined, undefined, {
    status,
    statusText: "",
    data: { secret: "private" },
    headers: { "retry-after": "2" },
    config: { headers: new AxiosHeaders({ Authorization: "secret" }), url: monitor },
  });
}

describe("SkillDirectoryClient", () => {
  const request = vi.fn<AxiosAdapter>();
  let body: unknown;
  let status: number;

  beforeEach(() => {
    vi.useFakeTimers();
    body = response();
    status = 200;
    const create = axios.create.bind(axios);
    vi.spyOn(axios, "create").mockImplementation((config) =>
      create({ ...config, adapter: request })
    );
    request.mockImplementation(async (config) => ({
      status,
      statusText: "",
      data: config.method === "get" ? { status: "completed" } : body,
      headers: {},
      config,
    }));
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.useRealTimers();
    request.mockReset();
  });

  function expectSingleDirectoryRequest(): void {
    expect(request).toHaveBeenCalledTimes(1);
    expect(request.mock.calls[0][0]).toEqual(
      expect.objectContaining({ method: "post", url: "/builder/v1/directories" })
    );
    expect(vi.getTimerCount()).toBe(0);
  }

  async function createDirectory() {
    const pending = new SkillDirectoryClient(
      "https://mos.example/",
      "mos-token"
    ).createReadyDirectory("agent", "code-review");
    await vi.runAllTimersAsync();
    return pending;
  }

  it("AC-20: sends the exact skip-copy body once with a 30-second timeout and no redirects", async () => {
    const result = await createDirectory();
    expect(result.isOk()).toBe(true);
    expectSingleDirectoryRequest();
    const config = request.mock.calls[0][0];
    expect(config.baseURL).toBe("https://mos.example");
    expect(config.timeout).toBe(30000);
    expect(config.maxRedirects).toBe(0);
    expect(config.headers.get("Authorization")).toBe("Bearer mos-token");
    expect(JSON.parse(config.data)).toEqual({
      agentId: "agent",
      directoryType: "skill",
      name: "code-review",
      skipPublishedContentCopy: true,
    });
  });

  it.each([
    ["missing", {}],
    ["undefined", { copyStatusUrls: undefined }],
    ["null", { copyStatusUrls: null }],
    ["empty", { copyStatusUrls: [] }],
  ])(
    "AC-21: accepts %s monitors under the exclusive ATK skip-copy prerequisite",
    async (_label, monitorFields) => {
      body = {
        folder,
        sharePointIds: response().sharePointIds,
        ...monitorFields,
      };
      const result = await createDirectory();
      expectSingleDirectoryRequest();
      expect(result.isOk() && result.value).toEqual({
        folder,
        // eslint-disable-next-line no-secrets/no-secrets -- Published MOS GUID-encoding test vector, not a credential.
        driveId: "b!LCwsLBoaS0uMjJ2dnZ2dnT8_Pz8rK0xMjY0ODg4ODg5KSkpKPDxNTY6OHx8fHx8f",
        itemId: response().sharePointIds.uniqueId,
      });
    }
  );

  it.each([{ monitors: [monitor] }, { monitors: [monitor, `${monitor}2`] }])(
    "AC-07: rejects reported copies without requesting any monitor ($monitors)",
    async ({ monitors }) => {
      body = { ...response(), copyStatusUrls: monitors };
      const defaultString = vi
        .spyOn(localizeUtils, "getDefaultString")
        .mockImplementation((key) => key);
      const localizedString = vi
        .spyOn(localizeUtils, "getLocalizedString")
        .mockImplementation((key) => key);
      const result = await createDirectory();
      expect.soft(result.isErr() && result.error).toBeInstanceOf(UserError);
      expect.soft(result.isErr() && result.error.name).toBe("SkillUploadCopyNotSupported");
      expect.soft(defaultString).toHaveBeenCalledWith("error.m365.skills.copyNotSupported");
      expect.soft(localizedString).toHaveBeenCalledWith("error.m365.skills.copyNotSupported");
      expect(JSON.stringify(result)).not.toMatch(/tempauth|private|secret|mos-token/);
      expectSingleDirectoryRequest();
    }
  );

  it("AC-07/16: makes no monitor request even when Axios inherits authorization", async () => {
    const inheritedAuthorization = axios.defaults.headers.common.Authorization;
    axios.defaults.headers.common.Authorization = "Bearer inherited-token";
    try {
      body = { ...response(), copyStatusUrls: [monitor] };
      const result = await createDirectory();
      expect.soft(result.isErr() && result.error.name).toBe("SkillUploadCopyNotSupported");
      expect(JSON.stringify(result)).not.toMatch(
        /tempauth|private|secret|mos-token|inherited-token/
      );
      expectSingleDirectoryRequest();
    } finally {
      if (inheritedAuthorization === undefined) {
        delete axios.defaults.headers.common.Authorization;
      } else {
        axios.defaults.headers.common.Authorization = inheritedAuthorization;
      }
    }
  });

  it.each([401, 403, 409, 412, 429, 500, 502, 503, 504])(
    "AC-20: never replays directory POST after HTTP %s even with Retry-After",
    async (httpStatus) => {
      request.mockRejectedValue(httpError(httpStatus));
      const result = await createDirectory();
      expect(result.isErr()).toBe(true);
      expect(JSON.stringify(result)).not.toMatch(/tempauth|private|secret|mos-token/);
      expectSingleDirectoryRequest();
    }
  );

  it.each(["ECONNRESET", "ETIMEDOUT", "ECONNABORTED", "EAI_AGAIN"])(
    "AC-20: never replays a failed or lost directory response (%s)",
    async (code) => {
      request.mockRejectedValue(new AxiosError("lost secret", code));
      const result = await createDirectory();
      expect(result.isErr()).toBe(true);
      expectSingleDirectoryRequest();
    }
  );

  it.each([201, 202, 301, 302, 307, 308])(
    "AC-20: rejects directory HTTP %s even with a valid body without following redirects",
    async (httpStatus) => {
      status = httpStatus;
      const result = await createDirectory();
      expect(result.isErr() && result.error.name).toBe("SkillDirectoryFailed");
      expectSingleDirectoryRequest();
      expect(request.mock.calls[0][0].maxRedirects).toBe(0);
    }
  );

  it.each(
    ["not-an-array", {}, 0, false, [1], [null], [monitor, 1]].map((copyStatusUrls) => ({
      copyStatusUrls,
    }))
  )(
    "AC-07/21: rejects malformed monitor fields $copyStatusUrls as invalidResponse",
    async ({ copyStatusUrls }) => {
      body = { ...response(), copyStatusUrls };
      const result = await createDirectory();
      expect(result.isErr() && result.error.name).toBe("SkillInvalidResponse");
      expectSingleDirectoryRequest();
    }
  );

  it.each([
    null,
    {},
    { ...response(), folder: "bad" },
    { ...response(), folder: "../skill" },
    { ...response(), sharePointIds: {} },
    { ...response(), sharePointIds: { ...response().sharePointIds, siteId: "bad" } },
    { ...response(), sharePointIds: { ...response().sharePointIds, webId: "bad" } },
    { ...response(), sharePointIds: { ...response().sharePointIds, listId: "bad" } },
    { ...response(), sharePointIds: { ...response().sharePointIds, uniqueId: "" } },
    { ...response(), sharePointIds: { ...response().sharePointIds, uniqueId: ".." } },
    { ...response(), sharePointIds: { ...response().sharePointIds, siteUrl: undefined } },
    { ...response(), folder: "bad", copyStatusUrls: [monitor] },
  ])("AC-21: rejects malformed directory response %j as invalidResponse", async (directoryBody) => {
    body = directoryBody;
    const result = await createDirectory();
    expect(result.isErr() && result.error.name).toBe("SkillInvalidResponse");
    expectSingleDirectoryRequest();
  });

  it.each([
    "http://tenant.sharepoint.com/monitor",
    "https://attacker.example/monitor",
    "https://other.sharepoint.com/monitor",
    "https://user:password@tenant.sharepoint.com/monitor",
    "https://tenant.sharepoint.com/monitor#secret",
    "https://tenant.sharepoint.com/monitor#",
    "https://@tenant.sharepoint.com/monitor",
    "https://tenant.sharepoint.com\\monitor",
    "https://tenant.sharepoint.com:444/monitor",
    "https://tenant.sharepoint.com/monitor with-space",
    "https://",
    "/relative",
  ])("AC-07/21: rejects unsafe monitor URL %s before requesting any URL", async (url) => {
    body = { ...response(), copyStatusUrls: [monitor, url] };
    const result = await createDirectory();
    expect(result.isErr() && result.error.name).toBe("SkillInvalidResponse");
    expectSingleDirectoryRequest();
  });

  it.each([
    "https://sharepoint.com.attacker.example",
    "https://evil.example",
    "http://tenant.sharepoint.com",
    "https://tenant.sharepoint.com:444",
    "https://user:password@tenant.sharepoint.com",
    "https://tenant.sharepoint.com/sites/skills#secret",
  ])("AC-21: rejects unsafe siteUrl %s as invalidResponse", async (url) => {
    body = {
      ...response(),
      sharePointIds: { ...response().sharePointIds, siteUrl: url },
      copyStatusUrls: [`${url}/monitor`],
    };
    const result = await createDirectory();
    expect(result.isErr() && result.error.name).toBe("SkillInvalidResponse");
    expectSingleDirectoryRequest();
  });

  it.each(["siteId", "webId", "listId", "uniqueId", "siteUrl"])(
    "AC-21: rejects missing lower-camel SharePoint field %s",
    async (field) => {
      body = {
        ...response(),
        sharePointIds: { ...response().sharePointIds, [field]: undefined },
      };
      const result = await createDirectory();
      expect(result.isErr() && result.error.name).toBe("SkillInvalidResponse");
      expectSingleDirectoryRequest();
    }
  );

  it.each([
    "tenant.sharepoint.us",
    "tenant.sharepoint-mil.us",
    "tenant.sharepoint.cn",
    "tenant.sharepoint.de",
  ])("AC-07/21: validates recognized SharePoint cloud host %s without polling", async (host) => {
    const directory = {
      ...response(),
      sharePointIds: { ...response().sharePointIds, siteUrl: `https://${host}/sites/skills` },
    };
    body = directory;
    const client = new SkillDirectoryClient("https://mos.example", "mos-token");
    expect((await client.createReadyDirectory("agent", "code-review")).isOk()).toBe(true);
    expectSingleDirectoryRequest();
    request.mockClear();

    body = { ...directory, copyStatusUrls: [`https://${host}/monitor`] };
    const result = await client.createReadyDirectory("agent", "code-review");
    expect.soft(result.isErr() && result.error.name).toBe("SkillUploadCopyNotSupported");
    expectSingleDirectoryRequest();
  });
});
