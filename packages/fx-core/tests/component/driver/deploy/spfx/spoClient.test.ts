// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import axios, { AxiosAdapter, AxiosRequestConfig, AxiosResponse } from "axios";
import { chai, vi } from "vitest";
import { SPOClient } from "../../../../../src/component/driver/deploy/spfx/utility/spoClient";

const expect = chai.expect;

describe("SPFx SPO Client", async () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("get app catalog site", async () => {
    vi.spyOn(axios, "create").mockReturnValue({
      defaults: { headers: { common: {} } },
      interceptors: {
        request: {
          use: vi.fn(),
        },
      },
      get: function <T = any, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig | undefined
      ): Promise<R> {
        return { data: { CorporateCatalogUrl: "fakeUrl" } } as any;
      },
    } as any);
    expect(await SPOClient.getAppCatalogSite("")).to.equal("fakeUrl");
  });

  it("get app catalog site - undefined", async () => {
    vi.spyOn(axios, "create").mockReturnValue({
      defaults: { headers: { common: {} } },
      interceptors: {
        request: {
          use: vi.fn(),
        },
      },
      get: function <T = any, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig | undefined
      ): Promise<R> {
        return { data: {} } as any;
      },
    } as any);
    expect(await SPOClient.getAppCatalogSite("")).to.equal(undefined);
  });

  it("upload app package", async () => {
    vi.spyOn(axios, "create").mockReturnValue({
      defaults: { headers: { common: {} } },
      interceptors: {
        request: {
          use: vi.fn(),
        },
      },
      post: function <T = any, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig | undefined
      ): Promise<R> {
        return { data: { CorporateCatalogUrl: "fakeUrl" } } as any;
      },
    } as any);
    await SPOClient.uploadAppPackage("", "", Buffer.from(""));
  });

  it("deploy app package", async () => {
    vi.spyOn(axios, "create").mockReturnValue({
      defaults: { headers: { common: {} } },
      interceptors: {
        request: {
          use: vi.fn(),
        },
      },
      post: function <T = any, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig | undefined
      ): Promise<R> {
        return { data: { CorporateCatalogUrl: "fakeUrl" } } as any;
      },
    } as any);
    await SPOClient.deployAppPackage("", "");
  });

  it("create app catelog", async () => {
    vi.spyOn(axios, "create").mockReturnValue({
      defaults: { headers: { common: {} } },
      interceptors: {
        request: {
          use: vi.fn(),
        },
      },
      post: function <T = any, R = AxiosResponse<T>>(
        url: string,
        config?: AxiosRequestConfig | undefined
      ): Promise<R> {
        return { data: { CorporateCatalogUrl: "fakeUrl" } } as any;
      },
    } as any);
    await SPOClient.createAppCatalog("");
  });
});

describe("SPFx SPO Client prepared catalog context", () => {
  let originalAdapter: typeof axios.defaults.adapter;

  beforeEach(() => {
    originalAdapter = axios.defaults.adapter;
  });

  afterEach(() => {
    axios.defaults.adapter = originalAdapter;
  });

  it("SPFX-PREP-06: interleaved prepared requests keep their own catalog while legacy calls use the module tenant", async () => {
    const seen: { baseURL?: string; url?: string; method?: string }[] = [];
    const pending: { url?: string; release: () => void }[] = [];
    const adapter: AxiosAdapter = (config) =>
      new Promise((resolve) => {
        seen.push({ baseURL: config.baseURL, url: config.url, method: config.method });
        const response = {
          data:
            config.method === "get" && config.url?.includes("AvailableApps")
              ? {
                  ID: "app",
                  AppCatalogVersion: "1.0.0.0",
                  Deployed: true,
                  CurrentVersionDeployed: true,
                  IsPackageDefaultSkipFeatureDeployment: false,
                }
              : { CorporateCatalogUrl: `${config.baseURL ?? ""}/sites/appcatalog` },
          status: 200,
          statusText: "OK",
          headers: {},
          config,
          request: {},
        };
        pending.push({ url: config.url, release: () => resolve(response) });
      });
    axios.defaults.adapter = adapter;
    const tenantA = "https://tenant-a.example.test/sites/appcatalog";
    const tenantB = "https://tenant-b.example.test/sites/appcatalog";

    SPOClient.setBaseUrl("https://legacy-1.example.test");
    const readA = SPOClient.getPreparedAvailableApp("token-a", "app", { catalogUrl: tenantA });
    SPOClient.setBaseUrl("https://legacy-2.example.test");
    const deployB = SPOClient.deployPreparedAppPackage("token-b", "app", false, {
      catalogUrl: tenantB,
    });
    const uploadA = SPOClient.uploadPreparedAppPackage(
      "token-a",
      "solution.sppkg",
      Buffer.from("package"),
      false,
      { catalogUrl: tenantA }
    );
    const legacy = SPOClient.getAppCatalogSite("token-legacy");
    await vi.waitFor(() => expect(pending).to.have.length(4));
    for (const request of [...pending].reverse()) {
      request.release();
    }
    await Promise.all([readA, deployB, uploadA, legacy]);

    expect(seen.map((request) => request.baseURL)).to.deep.equal([
      tenantA,
      tenantB,
      tenantA,
      "https://legacy-2.example.test",
    ]);
    expect(await legacy).to.equal("https://legacy-2.example.test/sites/appcatalog");
  });
});
