// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios, { AxiosAdapter, AxiosError } from "axios";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { setTools, TOOLS } from "../../../../src/common/globalVars";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../../../../src/common/telemetry";
import { WrappedAxiosClient } from "../../../../src/common/wrappedAxiosClient";
import { GraphSkillFileStore } from "../../../../src/component/m365/skills/graphSkillFileStore";
import { SkillDirectoryClient } from "../../../../src/component/m365/skills/skillDirectoryClient";
import { MockTools } from "../../../core/utils";

const graphEndpoint = "https://graph.microsoft.com";
const mosEndpoint = "https://titles.prod.mos.microsoft.com";
const file = {
  id: "private-item",
  name: "private-file.txt",
  eTag: "private-etag",
  size: 3,
  file: { hashes: {} },
};
const folder = {
  id: "private-folder-id",
  name: "private-folder",
  eTag: "private-etag",
  folder: {},
};
const directory = {
  folder: "11111111-1111-4111-8111-111111111111",
  sharePointIds: {
    siteId: "22222222-2222-4222-8222-222222222222",
    webId: "33333333-3333-4333-8333-333333333333",
    listId: "44444444-4444-4444-8444-444444444444",
    uniqueId: "55555555-5555-4555-8555-555555555555",
    siteUrl: "https://private-tenant.sharepoint.com/sites/private-site",
  },
  copyStatusUrls: [],
};
const operations = [
  ["directory", "mos_create_skill_directory", TelemetryEvent.MOSApi, "post"],
  ["list", "graph-list-drive-items", TelemetryEvent.DependencyApi, "get"],
  ["folder", "graph-create-drive-folder", TelemetryEvent.DependencyApi, "post"],
  ["upload", "graph-upload-drive-file", TelemetryEvent.DependencyApi, "put"],
  ["update", "graph-update-drive-file", TelemetryEvent.DependencyApi, "put"],
  ["delete", "graph-delete-drive-item", TelemetryEvent.DependencyApi, "delete"],
];
const operationPaths: Record<string, string> = {
  directory: "/builder/v1/directories",
  list: "/drives/private-drive/items/private-parent/children",
  folder: "/drives/private-drive/items/private-parent/children",
  upload: "/drives/private-drive/items/private-parent:/private-file.txt:/content",
  update: "/drives/private-drive/items/private-item/content",
  delete: "/drives/private-drive/items/private-item",
};

async function invoke(operation: string) {
  if (operation === "directory") {
    return new SkillDirectoryClient(mosEndpoint, "private-mos-token").createReadyDirectory(
      "private-agent",
      "private-skill"
    );
  }
  const store = new GraphSkillFileStore(graphEndpoint, "private-graph-token", "private-drive");
  switch (operation) {
    case "list":
      return store.listChildren("private-parent");
    case "folder":
      return store.createFolder("private-parent", folder.name);
    case "upload":
      return store.uploadFile("private-parent", file.name, Buffer.from("private-content"));
    case "update":
      return store.uploadFile("private-parent", file.name, Buffer.from("private-content"), {
        id: file.id,
        name: file.name,
        eTag: file.eTag,
        size: file.size,
        kind: "file",
      });
    case "delete":
      return store.deleteItem({
        id: file.id,
        name: file.name,
        eTag: file.eTag,
        kind: "file",
        size: 3,
      });
    default:
      throw new Error("Unknown test operation");
  }
}

describe("skill API telemetry", () => {
  const previousTools = TOOLS;
  const tools = new MockTools();
  const adapter = vi.fn<AxiosAdapter>();
  let status: number;

  beforeEach(() => {
    setTools(tools);
    vi.spyOn(tools.telemetryReporter, "sendTelemetryEvent");
    vi.spyOn(tools.telemetryReporter, "sendTelemetryErrorEvent");
    status = 200;
    const create = axios.create.bind(axios);
    vi.spyOn(axios, "create").mockImplementation((config) => create({ ...config, adapter }));
    adapter.mockImplementation(async (config) => {
      const response = {
        config,
        status,
        statusText: "",
        headers: { "request-id": "private-header", traceresponse: "private-trace" },
        data:
          config.url === "/builder/v1/directories"
            ? directory
            : config.method === "get"
              ? {
                  value: [file],
                  "@microsoft.graph.downloadUrl": "https://private-download/?token=secret",
                }
              : config.method === "post"
                ? folder
                : file,
      };
      if (status >= 400) {
        throw new AxiosError("private-transport-message", undefined, config, undefined, {
          ...response,
          data: { error: { code: "private-code", message: "private-server-message" } },
        });
      }
      if (config.method === "delete") response.status = 204;
      else if (config.method === "post" && config.url !== "/builder/v1/directories")
        response.status = 201;
      return response;
    });
  });

  afterEach(() => {
    vi.restoreAllMocks();
    adapter.mockReset();
    setTools(previousTools);
  });

  function assertNoPrivateTelemetry(): void {
    const events = [
      ...vi.mocked(tools.telemetryReporter.sendTelemetryEvent).mock.calls,
      ...vi.mocked(tools.telemetryReporter.sendTelemetryErrorEvent).mock.calls,
    ];
    expect(JSON.stringify(events)).not.toContain("private");
    expect(JSON.stringify(events)).not.toContain(directory.folder);
    expect(JSON.stringify(events)).not.toContain("secret");
    for (const [, properties] of events) {
      expect(
        Object.keys(properties ?? {}).every((key) =>
          ["url", "method", TelemetryProperty.Success, "status-code"].includes(key)
        )
      ).toBe(true);
    }
  }

  it.each(operations)(
    "AC-27 AC-28: %s emits redacted start and success events through the real client",
    async (operation, name, event, method) => {
      const result = await invoke(operation);
      expect(result.isOk()).toBe(true);
      expect(adapter).toHaveBeenCalledTimes(1);
      expect(
        WrappedAxiosClient.convertUrlToApiName(axios.getUri(adapter.mock.calls[0][0]), method)
      ).toBe(name);
      expect(tools.telemetryReporter.sendTelemetryEvent).toHaveBeenCalledTimes(2);
      const expectedStatus =
        operation === "delete" ? "204" : operation === "folder" ? "201" : "200";
      expect(tools.telemetryReporter.sendTelemetryEvent).toHaveBeenNthCalledWith(
        1,
        `${event}-start`,
        {
          url: `<${name}-url>`,
          method,
        }
      );
      expect(tools.telemetryReporter.sendTelemetryEvent).toHaveBeenNthCalledWith(2, event, {
        url: `<${name}-url>`,
        method,
        [TelemetryProperty.Success]: TelemetrySuccess.Yes,
        "status-code": expectedStatus,
      });
      expect(tools.telemetryReporter.sendTelemetryErrorEvent).not.toHaveBeenCalled();
      assertNoPrivateTelemetry();
    }
  );

  it.each(operations)(
    "AC-27 AC-28: %s emits a sanitized failure event without server bodies or trace values",
    async (operation, name, event, method) => {
      status = 403;
      const result = await invoke(operation);
      expect(result.isErr()).toBe(true);
      expect(adapter).toHaveBeenCalledTimes(1);
      expect(tools.telemetryReporter.sendTelemetryEvent).toHaveBeenCalledExactlyOnceWith(
        `${event}-start`,
        {
          url: `<${name}-url>`,
          method,
        }
      );
      expect(tools.telemetryReporter.sendTelemetryErrorEvent).toHaveBeenCalledExactlyOnceWith(
        event,
        {
          url: `<${name}-url>`,
          method,
          [TelemetryProperty.Success]: TelemetrySuccess.No,
          "status-code": "403",
        }
      );
      assertNoPrivateTelemetry();
    }
  );

  it.each(["https://graph.microsoft.us", "https://dod-graph.microsoft.us"])(
    "AC-27: classifies Graph calls in %s",
    async (endpoint) => {
      const result = await new GraphSkillFileStore(
        endpoint,
        "private-token",
        "private-drive"
      ).listChildren("private-parent");
      expect(result.isOk()).toBe(true);
      expect(tools.telemetryReporter.sendTelemetryEvent).toHaveBeenCalledWith(
        `${TelemetryEvent.DependencyApi}-start`,
        { url: "<graph-list-drive-items-url>", method: "get" }
      );
      assertNoPrivateTelemetry();
    }
  );

  it("AC-28: absolute pagination URLs keep the same API name without leaking query values", async () => {
    adapter.mockImplementationOnce(async (config) => ({
      config,
      status: 200,
      statusText: "",
      headers: {},
      data: {
        value: [],
        "@odata.nextLink": `${graphEndpoint}/v1.0/drives/private-drive/items/private-parent/children?$skiptoken=private-page`,
      },
    }));
    const result = await invoke("list");
    expect(result.isOk()).toBe(true);
    expect(adapter).toHaveBeenCalledTimes(2);
    expect(tools.telemetryReporter.sendTelemetryEvent).toHaveBeenCalledTimes(4);
    for (const [, properties] of vi.mocked(tools.telemetryReporter.sendTelemetryEvent).mock.calls) {
      expect(properties?.url).toBe("<graph-list-drive-items-url>");
    }
    assertNoPrivateTelemetry();
  });

  it("AC-27 AC-29: each throttled Graph attempt has telemetry without changing retries", async () => {
    adapter.mockImplementationOnce(async (config) => {
      throw new AxiosError("private-rate-limit", undefined, config, undefined, {
        config,
        status: 429,
        statusText: "",
        headers: { "retry-after": "0" },
        data: "private-body",
      });
    });
    expect((await invoke("list")).isOk()).toBe(true);
    expect(adapter).toHaveBeenCalledTimes(2);
    expect(tools.telemetryReporter.sendTelemetryEvent).toHaveBeenCalledTimes(3);
    expect(tools.telemetryReporter.sendTelemetryErrorEvent).toHaveBeenCalledTimes(1);
    assertNoPrivateTelemetry();
  });

  it("AC-28: transport failures with no config never fall back to raw error messages", async () => {
    adapter.mockRejectedValue(new AxiosError("private-error-message"));
    expect((await invoke("directory")).isErr()).toBe(true);
    expect(tools.telemetryReporter.sendTelemetryErrorEvent).toHaveBeenCalledExactlyOnceWith(
      TelemetryEvent.DependencyApi,
      {
        url: "<unclassified-api-url>",
        method: "unknown",
        [TelemetryProperty.Success]: TelemetrySuccess.No,
        "status-code": "undefined",
      }
    );
    assertNoPrivateTelemetry();
  });

  it("AC-28 AC-29: response config identifies a failure without changing the rejected error", async () => {
    let originalError: AxiosError | undefined;
    adapter.mockImplementationOnce(async (config) => {
      originalError = new AxiosError("private-error", undefined, undefined, undefined, {
        config,
        status: 403,
        statusText: "",
        headers: {},
        data: "private-response",
      });
      throw originalError;
    });
    const client = WrappedAxiosClient.create({ baseURL: mosEndpoint });
    await expect(client.post("/builder/v1/directories")).rejects.toSatisfy(
      (error: unknown) => error === originalError
    );
    expect(tools.telemetryReporter.sendTelemetryErrorEvent).toHaveBeenCalledExactlyOnceWith(
      TelemetryEvent.MOSApi,
      {
        url: "<mos_create_skill_directory-url>",
        method: "post",
        [TelemetryProperty.Success]: TelemetrySuccess.No,
        "status-code": "403",
      }
    );
    assertNoPrivateTelemetry();
  });

  it.each(
    operations.flatMap(([operation, name, event, method]) =>
      [200, 403].map((responseStatus) => ({ operation, name, event, method, responseStatus }))
    )
  )(
    "AC-34: ordinary wrapper calls automatically protect $operation at HTTP $responseStatus",
    async ({ operation, name, event, method, responseStatus }) => {
      status = responseStatus;
      const client = WrappedAxiosClient.create({
        baseURL: operation === "directory" ? mosEndpoint : `${graphEndpoint}/v1.0`,
        headers: { Authorization: "private-token" },
      });
      const request = client.request({
        url: operationPaths[operation],
        method,
        params: { secret: "private-query" },
        data: "private-body",
      });
      if (responseStatus === 200) await expect(request).resolves.toBeDefined();
      else await expect(request).rejects.toBeInstanceOf(AxiosError);
      expect(adapter).toHaveBeenCalledTimes(1);
      expect(tools.telemetryReporter.sendTelemetryEvent).toHaveBeenCalledWith(`${event}-start`, {
        url: `<${name}-url>`,
        method,
      });
      const calls =
        responseStatus === 200
          ? vi.mocked(tools.telemetryReporter.sendTelemetryEvent).mock.calls
          : vi.mocked(tools.telemetryReporter.sendTelemetryErrorEvent).mock.calls;
      expect(calls.at(-1)?.[1]).toMatchObject({
        url: `<${name}-url>`,
        method,
        [TelemetryProperty.Success]:
          responseStatus === 200 ? TelemetrySuccess.Yes : TelemetrySuccess.No,
      });
      assertNoPrivateTelemetry();
    }
  );

  it.each(
    [
      {
        baseURL: graphEndpoint,
        url: "/v1.0/drives/private-drive/items/private-id/unknown?secret=private-query",
        event: TelemetryEvent.DependencyApi,
      },
      {
        baseURL: graphEndpoint,
        url: "/beta/drives/private-drive/items/private-id/children",
        event: TelemetryEvent.DependencyApi,
      },
      {
        baseURL: mosEndpoint,
        url: "/builder/v1/directories/private-id/unknown",
        event: TelemetryEvent.MOSApi,
      },
      { baseURL: mosEndpoint, url: "/builder/v1/directories", event: TelemetryEvent.MOSApi },
      { baseURL: undefined, url: "https://[private-invalid", event: TelemetryEvent.DependencyApi },
      { baseURL: undefined, url: "", event: TelemetryEvent.DependencyApi },
    ].flatMap((route) => [200, 403].map((responseStatus) => ({ ...route, responseStatus })))
  )(
    "AC-28 AC-34: unknown or malformed protected route $url uses a fixed fallback at HTTP $responseStatus",
    async ({ baseURL, url, event, responseStatus }) => {
      status = responseStatus;
      const client = WrappedAxiosClient.create({
        baseURL,
        headers: { Authorization: "private-token" },
      });
      const request = client.get(url, {
        params: { secret: "private-param" },
      });
      if (responseStatus === 200) await expect(request).resolves.toMatchObject({ status: 200 });
      else await expect(request).rejects.toBeInstanceOf(AxiosError);
      expect(tools.telemetryReporter.sendTelemetryEvent).toHaveBeenCalledWith(`${event}-start`, {
        url: "<unclassified-api-url>",
        method: "get",
      });
      assertNoPrivateTelemetry();
    }
  );

  it("AC-34: missing config failures are sanitized even outside the skill clients", async () => {
    const error = new AxiosError("private-error");
    adapter.mockRejectedValue(error);
    const client = WrappedAxiosClient.create({ baseURL: graphEndpoint });
    await expect(
      client.get("/v1.0/drives/private-drive/items/private-parent/children")
    ).rejects.toBe(error);
    expect(tools.telemetryReporter.sendTelemetryErrorEvent).toHaveBeenCalledExactlyOnceWith(
      TelemetryEvent.DependencyApi,
      {
        url: "<unclassified-api-url>",
        method: "unknown",
        success: "no",
        "status-code": "undefined",
      }
    );
    assertNoPrivateTelemetry();
  });

  it.each([
    {
      baseURL: "https://alternate.example",
      url: "/builder/v1/directories",
      method: "post",
      name: "mos_create_skill_directory",
    },
    {
      baseURL: "https://alternate.example/v1.0",
      url: "/drives/private-drive/items/private-parent/children",
      method: "get",
      name: "graph-list-drive-items",
    },
    {
      baseURL: "https://alternate.example",
      url: `${graphEndpoint}/v1.0/drives/private-drive/items/private-parent/children?secret=private-query`,
      method: "get",
      name: "graph-list-drive-items",
    },
  ])(
    "AC-34: route classification protects alternate endpoints and absolute URLs ($method $url)",
    async ({ baseURL, url, method, name }) => {
      const client = WrappedAxiosClient.create({ baseURL });
      await expect(client.request({ url, method })).resolves.toBeDefined();
      expect(tools.telemetryReporter.sendTelemetryEvent).toHaveBeenCalledTimes(2);
      for (const [, properties] of vi.mocked(tools.telemetryReporter.sendTelemetryEvent).mock
        .calls) {
        expect(properties?.url).toBe(`<${name}-url>`);
      }
      assertNoPrivateTelemetry();
    }
  );

  it.each(["absent", "throwing"])(
    "AC-29: %s telemetry reporters do not alter success or failure results",
    async (mode) => {
      if (mode === "absent") setTools({ ...tools, telemetryReporter: undefined });
      else {
        vi.mocked(tools.telemetryReporter.sendTelemetryEvent).mockImplementation(() => {
          throw new Error("reporter failure");
        });
        vi.mocked(tools.telemetryReporter.sendTelemetryErrorEvent).mockImplementation(() => {
          throw new Error("reporter failure");
        });
      }
      expect((await invoke("directory")).isOk()).toBe(true);
      status = 403;
      const failed = await invoke("directory");
      expect(failed.isErr() && failed.error.name).toBe("SkillUploadAuthorizationFailed");
      status = 200;
      adapter.mockImplementationOnce(async (config) => {
        throw new AxiosError("private-rate-limit", undefined, config, undefined, {
          config,
          status: 429,
          statusText: "",
          headers: { "retry-after": "0" },
          data: "private-body",
        });
      });
      expect((await invoke("list")).isOk()).toBe(true);
      expect(adapter).toHaveBeenCalledTimes(4);
    }
  );

  it.each([
    {
      baseURL: mosEndpoint,
      url: "/builder/v1/users/packages?next=/builder/v1/directories",
      event: TelemetryEvent.MOSApi,
      name: "mos_post_builder_package",
    },
    {
      baseURL: graphEndpoint,
      url: "/v1.0/appCatalogs/teamsApps?next=/v1.0/drives",
      event: TelemetryEvent.DependencyApi,
      name: "graph-publish-app",
    },
  ])(
    "AC-29 AC-34: unrelated API $name retains its telemetry even with protected paths in query values",
    async ({ baseURL, url, event, name }) => {
      let originalError: AxiosError | undefined;
      adapter.mockImplementationOnce(async (config) => {
        originalError = new AxiosError("legacy-failure", undefined, config, undefined, {
          config,
          status: 403,
          statusText: "",
          headers: {},
          data: { error: { code: "legacy-code", message: "legacy-details" } },
        });
        throw originalError;
      });
      const client = WrappedAxiosClient.create({ baseURL });
      await expect(client.post(url, undefined, { params: { legacy: "value" } })).rejects.toSatisfy(
        (error: unknown) => error === originalError
      );
      expect(tools.telemetryReporter.sendTelemetryEvent).toHaveBeenCalledExactlyOnceWith(
        `${event}-start`,
        { url: `<${name}-url>`, method: "post", params: "legacy:value" }
      );
      expect(tools.telemetryReporter.sendTelemetryErrorEvent).toHaveBeenCalledExactlyOnceWith(
        event,
        expect.objectContaining({
          url: `<${name}-url>`,
          params: "legacy:value",
          [TelemetryProperty.ErrorMessage]: expect.stringContaining("legacy-details"),
        })
      );
    }
  );
});
