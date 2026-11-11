// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios, { AxiosAdapter, AxiosError } from "axios";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { GraphClient } from "../../../../../src/client/graphClient";
import { RetryHandler } from "../../../../../src/common/retryHandler";
import { SPOClient } from "../../../../../src/component/driver/deploy/spfx/utility/spoClient";
import { MockedM365Provider } from "../../../../core/utils";
import { appId, productId, publicationTenantContext } from "./preparedPublicationFixtures";

describe("Prepared Graph catalog and submission boundary", () => {
  const app = {
    id: "catalog-app-id",
    externalId: appId,
    displayName: "Sample",
    appDefinitions: [{ version: "2.3.4", publishingState: "published" }],
  };

  beforeEach(() => {
    RetryHandler.RETRIES = 1;
  });

  afterEach(() => {
    RetryHandler.RETRIES = 6;
    vi.restoreAllMocks();
  });

  function requester() {
    const http = axios.create();
    vi.spyOn(axios, "create").mockReturnValue(http);
    return { http, client: new GraphClient(new MockedM365Provider()) };
  }

  it.each([
    { name: "no HTTP response", response: undefined, message: "did not return" },
    {
      name: "invalid root",
      response: { data: null },
      message: "invalid Teams app catalog response",
    },
    {
      name: "missing result array",
      response: { data: {} },
      message: "invalid Teams app catalog response",
    },
    {
      name: "non-object result",
      response: { data: { value: [null] } },
      message: "incomplete Teams app catalog result",
    },
    {
      name: "paginated result",
      response: { data: { value: [app], "@odata.nextLink": "next" } },
      message: "incomplete Teams app catalog result",
    },
    {
      name: "multiple matching apps",
      response: { data: { value: [app, app] } },
      message: "multiple Teams apps",
    },
    {
      name: "foreign external ID",
      response: { data: { value: [{ ...app, externalId: productId }] } },
      message: "incomplete Teams app catalog entry",
    },
    {
      name: "missing definitions",
      response: { data: { value: [{ ...app, appDefinitions: undefined }] } },
      message: "incomplete Teams app catalog entry",
    },
    {
      name: "paginated definitions",
      response: { data: { value: [{ ...app, "appDefinitions@odata.nextLink": "next" }] } },
      message: "incomplete Teams app catalog entry",
    },
    {
      name: "non-object definition",
      response: { data: { value: [{ ...app, appDefinitions: [null] }] } },
      message: "incomplete Teams app catalog entry",
    },
    {
      name: "incomplete definition",
      response: { data: { value: [{ ...app, appDefinitions: [{ version: "2.3.4" }] }] } },
      message: "incomplete Teams app definition",
    },
    {
      name: "no definitions",
      response: { data: { value: [{ ...app, appDefinitions: [] }] } },
      message: "without definitions",
    },
  ])(
    "SPFX-PREP-07: refuses $name before a prepared submission can be selected",
    async ({ response, message }) => {
      const { http, client } = requester();
      vi.spyOn(http, "get").mockResolvedValue(response);
      const post = vi.spyOn(http, "post");
      await expect(client.getPreparedStagedApp("token", appId)).rejects.toThrow(message);
      expect(post).not.toHaveBeenCalled();
    }
  );

  it("SPFX-PREP-07: returns only complete definitions and accepts an empty catalog", async () => {
    const { http, client } = requester();
    vi.spyOn(http, "get")
      .mockResolvedValueOnce({ data: { value: [{ ...app, externalId: appId.toUpperCase() }] } })
      .mockResolvedValueOnce({ data: { value: [] } });
    expect(await client.getPreparedStagedApp("token", appId)).toEqual({
      teamsAppId: app.id,
      displayName: app.displayName,
      definitions: app.appDefinitions,
    });
    expect(await client.getPreparedStagedApp("token", appId)).toBeUndefined();
  });

  it("SPFX-PREP-08: observes pre-cancellation before creating a Graph requester", async () => {
    const controller = new AbortController();
    controller.abort();
    const create = vi.spyOn(axios, "create");
    const client = new GraphClient(new MockedM365Provider());
    await expect(
      client.publishPreparedTeamsApp(
        "token",
        appId,
        undefined,
        Buffer.from("app"),
        controller.signal
      )
    ).rejects.toThrow("cancelled");
    expect(create).not.toHaveBeenCalled();
  });

  it("SPFX-PREP-07: submits an update exactly once and keeps the selected catalog ID", async () => {
    const { http, client } = requester();
    const post = vi.spyOn(http, "post").mockResolvedValue({ data: {} });
    const retry = vi.spyOn(RetryHandler, "Retry");
    const bytes = Buffer.from("prepared application");
    expect(await client.publishPreparedTeamsApp("token", appId, app.id, bytes)).toBe(app.id);
    expect(post).toHaveBeenCalledOnce();
    expect(post.mock.calls[0]).toEqual([
      `/appCatalogs/teamsApps/${app.id}/appDefinitions?requiresReview=true`,
      bytes,
      expect.objectContaining({ signal: expect.any(AbortSignal) }),
    ]);
    expect(retry).not.toHaveBeenCalled();
  });

  it.each([{ id: "missing-external-id" }, { id: "foreign", externalId: productId }, null])(
    "SPFX-PREP-08: does not infer a create success from an unbound Graph response (%j)",
    async (data) => {
      const { http, client } = requester();
      const post = vi.spyOn(http, "post").mockResolvedValue({ data });
      await expect(
        client.publishPreparedTeamsApp("token", appId, undefined, Buffer.from("app"))
      ).rejects.toThrow("without returning its app ID");
      expect(post).toHaveBeenCalledOnce();
    }
  );
});

describe("Prepared SharePoint HTTP response validation", () => {
  let originalAdapter: typeof axios.defaults.adapter;
  let data: unknown;
  let status: number;
  const context = { catalogUrl: publicationTenantContext.catalogUrl };

  beforeEach(() => {
    originalAdapter = axios.defaults.adapter;
    status = 200;
    const adapter: AxiosAdapter = async (config) => {
      const response = { data, status, statusText: String(status), headers: {}, config };
      if (status >= 400) {
        throw new AxiosError(
          "Catalog read failed",
          AxiosError.ERR_BAD_REQUEST,
          config,
          undefined,
          response
        );
      }
      return response;
    };
    axios.defaults.adapter = adapter;
  });

  afterEach(() => {
    axios.defaults.adapter = originalAdapter;
    vi.restoreAllMocks();
  });

  it.each([
    null,
    {},
    {
      ID: productId,
      AppCatalogVersion: "1.2.3",
      Deployed: true,
      CurrentVersionDeployed: true,
      IsPackageDefaultSkipFeatureDeployment: false,
    },
  ])("SPFX-PREP-06: refuses incomplete app metadata (%j)", async (value) => {
    data = value;
    await expect(SPOClient.getPreparedAvailableApp("token", productId, context)).rejects.toThrow(
      "incomplete app catalog metadata"
    );
  });

  it("SPFX-PREP-06: treats only 404 as absent and preserves other catalog errors", async () => {
    status = 404;
    expect(await SPOClient.getPreparedAvailableApp("token", productId, context)).toBeUndefined();
    status = 403;
    await expect(
      SPOClient.getPreparedAvailableApp("token", productId, context)
    ).rejects.toMatchObject({
      response: { status: 403 },
    });
  });

  it.each([
    { value: null, message: "invalid app catalog file listing" },
    { value: {}, message: "invalid app catalog file listing" },
    { value: { value: [{}, {}] }, message: "ambiguous app catalog package file" },
    { value: { value: [{}] }, message: "ambiguous app catalog package file" },
    { value: { value: [{ File: {} }] }, message: "invalid app catalog package URL" },
    {
      value: { value: [{ File: { ServerRelativeUrl: "https://foreign.test/package" } }] },
      message: "invalid app catalog package URL",
    },
    {
      value: { value: [{ File: { ServerRelativeUrl: "/bad\\path" } }] },
      message: "invalid app catalog package URL",
    },
    {
      value: { value: [{ File: { ServerRelativeUrl: "/bad\r\npath" } }] },
      message: "invalid app catalog package URL",
    },
  ])("SPFX-PREP-06: refuses unsafe package discovery ($message)", async ({ value, message }) => {
    data = value;
    await expect(
      SPOClient.readPreparedAppPackage("token", "solution.sppkg", context)
    ).rejects.toThrow(message);
  });

  it("SPFX-PREP-06: represents no matching package without fabricating bytes", async () => {
    data = { value: [] };
    expect(
      await SPOClient.readPreparedAppPackage("token", "solution.sppkg", context)
    ).toBeUndefined();
  });

  it.each(["../solution.sppkg", "quote's.sppkg", "solution.zip"])(
    "SPFX-PREP-06: refuses unsafe upload and read filenames (%s) before HTTP",
    async (name) => {
      const create = vi.spyOn(axios, "create");
      await expect(SPOClient.readPreparedAppPackage("token", name, context)).rejects.toThrow(
        "file name is invalid"
      );
      await expect(
        SPOClient.uploadPreparedAppPackage("token", name, Buffer.from("bytes"), false, context)
      ).rejects.toThrow("file name is invalid");
      expect(create).not.toHaveBeenCalled();
    }
  );
});
