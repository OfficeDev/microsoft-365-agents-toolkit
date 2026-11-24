// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios, { AxiosAdapter, AxiosError } from "axios";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { SPOClient } from "../../../../../src/component/driver/deploy/spfx/utility/spoClient";
import { productId, publicationTenantContext } from "./preparedPublicationFixtures";

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
