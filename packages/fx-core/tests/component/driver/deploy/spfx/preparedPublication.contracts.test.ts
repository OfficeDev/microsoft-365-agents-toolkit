// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { DriverContext } from "../../../../../src/component/driver/interface/commonArgs";
import { FxError, Platform, Result, UserError, err, ok } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import { createHash } from "crypto";
import axios, {
  AxiosAdapter,
  AxiosError,
  AxiosHeaders,
  AxiosResponse,
  InternalAxiosRequestConfig,
} from "axios";
import fs from "fs-extra";
import os from "os";
import path from "path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import * as Tools from "../../../../../src/common/tools";
import { SPFxDeployDriver } from "../../../../../src/component/driver/deploy/spfx/deployDriver";
import {
  loadPreparedPublication,
  preparePublication,
  PreparedPublicationJournal,
  PreparedPublicationOutput,
  verifyPreparedApplicationPackage,
  verifyPreparedArtifact,
  verifyPreparedSharePointPackage,
  verifyPreparedSources,
  withPreparedPublicationIdentityLock,
} from "../../../../../src/component/driver/deploy/spfx/utility/preparedPublication";
import { sleep } from "../../../../../src/component/driver/deploy/spfx/utility/sleep";
import { PreparePublicationDriver } from "../../../../../src/component/driver/deploy/spfx/preparePublicationDriver";
import { PublishAppPackageDriver } from "../../../../../src/component/driver/teamsApp/publishAppPackage";
import { UserCancelError } from "../../../../../src/error/common";
import { MockTools } from "../../../../core/utils";
import {
  createAppPackage,
  createSppkg,
  mockTokenClaims,
  appId,
  productId,
  publicationTenantContext,
  tenantId,
  TokenClaimsOverride,
  tenantUrl,
} from "./preparedPublicationFixtures";

vi.mock("../../../../../src/component/driver/deploy/spfx/utility/sleep", () => ({
  sleep: vi.fn(async () => undefined),
}));

const catalogUrl = `${tenantUrl}/sites/appcatalog`;
const otherTenantId = "dddddddd-eeee-4fff-8aaa-111111111111";

type Route = (config: InternalAxiosRequestConfig) => Promise<AxiosResponse>;

interface FakeCatalog {
  packageBytes?: Buffer;
  deployed: boolean;
  version?: string;
  ignoreDeploy?: boolean;
}

function reply(
  config: InternalAxiosRequestConfig,
  status: number,
  data?: unknown,
  responseHeaders?: Record<string, string>
): AxiosResponse {
  const response: AxiosResponse = {
    data,
    status,
    statusText: String(status),
    headers: new AxiosHeaders(responseHeaders),
    config,
    request: {},
  };
  if (status >= 400) {
    throw new AxiosError(
      `Request failed with status code ${status}`,
      AxiosError.ERR_BAD_REQUEST,
      config,
      {},
      response
    );
  }
  return response;
}

/** Serves the SharePoint REST surface used by prepared publication from in-memory catalog state. */
function catalogRoute(catalog: FakeCatalog): Route {
  return async (config) => {
    const url = config.url ?? "";
    if (config.method === "get" && url === "/_api/SP_TenantSettings_Current") {
      return reply(config, 200, { CorporateCatalogUrl: catalogUrl });
    }
    if (config.method === "get" && url.includes("getbytitle('Apps for SharePoint')")) {
      return reply(config, 200, {
        value: catalog.packageBytes
          ? [{ File: { ServerRelativeUrl: "/sites/appcatalog/solution.sppkg" } }]
          : [],
      });
    }
    if (config.method === "get" && url.includes("GetFileByServerRelativeUrl")) {
      return catalog.packageBytes ? reply(config, 200, catalog.packageBytes) : reply(config, 404);
    }
    if (config.method === "get" && url.includes("AvailableApps/GetById")) {
      return catalog.packageBytes
        ? reply(config, 200, {
            ID: productId,
            AppCatalogVersion: catalog.version ?? "1.2.3.4",
            Deployed: catalog.deployed,
            CurrentVersionDeployed: catalog.deployed,
            IsPackageDefaultSkipFeatureDeployment: false,
          })
        : reply(config, 404);
    }
    if (config.method === "post" && url.includes("tenantappcatalog/Add(")) {
      if (!Buffer.isBuffer(config.data)) {
        throw new Error("The fake catalog expected package bytes.");
      }
      catalog.packageBytes = config.data;
      catalog.version = undefined;
      catalog.deployed = false;
      return reply(config, 200, {});
    }
    if (config.method === "post" && url.endsWith("/Deploy")) {
      catalog.deployed = catalog.ignoreDeploy !== true;
      return reply(config, 200, {});
    }
    throw new Error(`Unexpected SharePoint request: ${config.method ?? ""} ${url}`);
  };
}

function hangUntilAborted(config: InternalAxiosRequestConfig): Promise<AxiosResponse> {
  return new Promise((_resolve, reject) => {
    const signal = config.signal;
    if (!(signal instanceof AbortSignal)) {
      reject(new Error("The request had no settlement signal."));
      return;
    }
    signal.addEventListener("abort", () => reject(new axios.CanceledError()), { once: true });
  });
}

describe("Prepared SharePoint publication contracts", () => {
  let projectPath: string | undefined;
  let originalAdapter: typeof axios.defaults.adapter;
  let requests: InternalAxiosRequestConfig[];
  let route: Route;

  beforeEach(() => {
    requests = [];
    originalAdapter = axios.defaults.adapter;
    const adapter: AxiosAdapter = async (config) => {
      requests.push(config);
      return route(config);
    };
    axios.defaults.adapter = adapter;
    vi.mocked(sleep).mockClear();
  });

  afterEach(async () => {
    axios.defaults.adapter = originalAdapter;
    vi.restoreAllMocks();
    if (projectPath !== undefined) {
      await fs.remove(projectPath);
      projectPath = undefined;
    }
  });

  async function prepareProject(
    sppkg = createSppkg()
  ): Promise<{ root: string; prepared: PreparedPublicationOutput }> {
    projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-sharepoint-contract-"));
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), sppkg);
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), createAppPackage());
    const prepared = await preparePublication({
      projectPath,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext: publicationTenantContext,
    });
    if (prepared.isErr()) {
      throw prepared.error;
    }
    return { root: await fs.realpath(projectPath), prepared: prepared.value };
  }

  function createContext(
    claims: TokenClaimsOverride = {},
    signal?: AbortSignal,
    onApproval?: () => void
  ): { context: DriverContext; approvals: () => number } {
    vi.spyOn(SPFxDeployDriver.prototype, "getTenant").mockResolvedValue(tenantUrl);
    vi.spyOn(Tools, "getSPFxToken").mockResolvedValue("sharepoint-token");
    const tools = new MockTools();
    mockTokenClaims(tools.tokenProvider.m365TokenProvider, claims);
    let approvalCount = 0;
    vi.spyOn(tools.ui, "showMessage").mockImplementation(async (_level, _message, modal) => {
      if (!modal) return ok("");
      approvalCount += 1;
      onApproval?.();
      return ok("Confirm");
    });
    return {
      context: {
        azureAccountProvider: tools.tokenProvider.azureAccountProvider,
        m365TokenProvider: tools.tokenProvider.m365TokenProvider,
        ui: tools.ui,
        progressBar: undefined,
        logProvider: tools.logProvider,
        projectPath: projectPath ?? "",
        platform: Platform.VSCode,
        signal,
      },
      approvals: () => approvalCount,
    };
  }

  function deployArgs(prepared: PreparedPublicationOutput) {
    return {
      createAppCatalogIfNotExist: false,
      packageSolutionPath: "sharepoint/config/package-solution.json",
      preparedPublicationPath: prepared.preparedPublicationPath,
      publicationId: prepared.publicationId,
      preparedPackagePath: prepared.sharePointPackagePath,
      sourcePackageSha256: prepared.sourcePackageSha256,
      sourceAppPackageSha256: prepared.sourceAppPackageSha256,
      pluginSha256: prepared.pluginSha256,
      sharePointPackageSha256: prepared.sharePointPackageSha256,
      appPackageSha256: prepared.appPackageSha256,
    };
  }

  function publishArgs(prepared: PreparedPublicationOutput) {
    return {
      appPackagePath: prepared.appPackagePath,
      preparedPublicationPath: prepared.preparedPublicationPath,
      publicationId: prepared.publicationId,
      sourcePackageSha256: prepared.sourcePackageSha256,
      sourceAppPackageSha256: prepared.sourceAppPackageSha256,
      pluginSha256: prepared.pluginSha256,
      sharePointPackageSha256: prepared.sharePointPackageSha256,
      appPackageSha256: prepared.appPackageSha256,
    };
  }

  async function journalOf(prepared: PreparedPublicationOutput) {
    const loaded = await loadPreparedPublication(
      projectPath ?? "",
      prepared.publicationId,
      prepared.preparedPublicationPath
    );
    if (loaded.isErr()) {
      throw loaded.error;
    }
    return loaded.value.journal;
  }

  function errorCode(result: Result<unknown, FxError>): string | undefined {
    return result.isErr() ? result.error.message : undefined;
  }

  function posts(): InternalAxiosRequestConfig[] {
    return requests.filter((request) => request.method === "post");
  }

  it("SPFX-PREP-06: uploads, deploys and verifies a prepared package through the app catalog", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    route = catalogRoute(catalog);
    const { context, approvals } = createContext();

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(result.isErr() ? result.error : undefined).toBeUndefined();
    expect(approvals()).toBe(2);
    expect(posts().map((request) => request.url)).toEqual([
      "/_api/web/tenantappcatalog/Add(overwrite=false, url='solution.sppkg')",
      `/_api/web/tenantappcatalog/AvailableApps/GetById('${productId}')/Deploy`,
    ]);
    const journal = await journalOf(prepared);
    expect(journal.state).toBe("sharePointVerified");
    expect(catalog.packageBytes).toBeDefined();
    expect(
      requests
        .filter((request) => request.url !== "/_api/SP_TenantSettings_Current")
        .every((request) => request.baseURL === catalogUrl)
    ).toBe(true);
  });

  it.each([
    {
      name: "Graph token for another tenant",
      claims: { graphTenantId: otherTenantId },
      code: "PreparedTenantMismatch",
    },
    {
      name: "Graph token with a non-Graph audience",
      claims: { graphAudience: "https://other.example.test" },
      code: "GraphTokenIdentityMismatch",
    },
    {
      name: "SharePoint token from a different tenant than Graph",
      claims: { sharePointTenantId: otherTenantId },
      code: "SharePointTokenIdentityMismatch",
    },
    {
      name: "SharePoint token with another resource audience",
      claims: { sharePointAudience: "https://other.example.test" },
      code: "SharePointTokenIdentityMismatch",
    },
  ])(
    "SPFX-PREP-05: refuses prepared deployment with a $name before any approval or write",
    async ({ claims, code }) => {
      const { prepared } = await prepareProject();
      route = catalogRoute({ deployed: false });
      const { context, approvals } = createContext(claims);

      const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

      expect(errorCode(result)).toBe(code);
      expect(approvals()).toBe(0);
      expect(posts()).toHaveLength(0);
      expect(requests.map((request) => request.url)).toEqual(["/_api/SP_TenantSettings_Current"]);
      expect((await journalOf(prepared)).state).toBe("prepared");
    }
  );

  describe("spfx/preparePublication driver boundary", () => {
    const outputNames = new Map(
      [
        "publicationId",
        "preparedPublicationPath",
        "sourcePackageSha256",
        "sourceAppPackageSha256",
        "pluginSha256",
        "sharePointPackagePath",
        "sharePointPackageSha256",
        "appPackagePath",
        "appPackageSha256",
      ].map((key) => [key, `PREPARED_${key.toUpperCase()}`])
    );
    const args = {
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
    };

    async function writeInputs(): Promise<void> {
      projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-sharepoint-driver-"));
      await fs.outputFile(path.join(projectPath, args.packagePath), createSppkg());
      await fs.outputFile(path.join(projectPath, args.appPackagePath), createAppPackage());
    }

    async function preparedDirectoryExists(): Promise<boolean> {
      return fs.pathExists(path.join(projectPath ?? "", ".fx", "sharepoint-publications"));
    }

    it("SPFX-PREP-05: discovers the verified catalog and writes every bound output", async () => {
      await writeInputs();
      route = catalogRoute({ deployed: false });
      const { context } = createContext();

      const result = await new PreparePublicationDriver().execute(args, context, outputNames);

      expect(result.result.isOk()).toBe(true);
      const outputs = result.result.isOk() ? result.result.value : new Map<string, string>();
      expect([...outputs.keys()].sort()).toEqual([...outputNames.values()].sort());
      expect(outputs.get("PREPARED_APPPACKAGESHA256")).toMatch(/^[a-f0-9]{64}$/);
      expect(posts()).toHaveLength(0);
    });

    it.each([
      ["missing inputs", {}, "InvalidActionInput"],
      ["unexpected inputs", { ...args, extra: "x" }, "InvalidActionInput"],
    ])("SPFX-PREP-05: rejects %s before any request", async (_name, input, code) => {
      await writeInputs();
      const { context } = createContext();

      const result = await new PreparePublicationDriver().execute(input, context, outputNames);

      expect(errorCode(result.result)).toBe(code);
      expect(requests).toHaveLength(0);
    });

    it.each([
      ["missing", undefined],
      ["incomplete", new Map([["publicationId", "PREPARED_ID"]])],
      ["invalid variable", new Map([...outputNames].map(([key]) => [key, "1-invalid"]))],
      ["duplicate variable", new Map([...outputNames].map(([key]) => [key, "SAME"]))],
    ])("SPFX-PREP-05: rejects %s output bindings before any request", async (_name, outputs) => {
      await writeInputs();
      const { context } = createContext();

      const result = await new PreparePublicationDriver().execute(args, context, outputs);

      expect(errorCode(result.result)).toBe("InvalidActionOutput");
      expect(requests).toHaveLength(0);
    });

    it.each([
      ["a non-HTTPS tenant URL", "http://contoso.example.test", "InvalidPreparedTenantContext"],
      ["an unparsable tenant URL", "not a url", "InvalidPreparedTenantContext"],
    ])("SPFX-PREP-05: refuses %s without writing a journal", async (_name, tenant, code) => {
      await writeInputs();
      const { context } = createContext();
      vi.mocked(SPFxDeployDriver.prototype.getTenant).mockResolvedValue(tenant);

      const result = await new PreparePublicationDriver().execute(args, context, outputNames);

      expect(errorCode(result.result)).toBe(code);
      expect(requests).toHaveLength(0);
      expect(await preparedDirectoryExists()).toBe(false);
    });

    it.each([
      [
        "split token tenants",
        { sharePointTenantId: otherTenantId },
        "SharePointTokenIdentityMismatch",
      ],
      [
        "a foreign SharePoint audience",
        { sharePointAudience: "https://other.example.test" },
        "SharePointTokenIdentityMismatch",
      ],
      ["a non-Graph audience", { graphAudience: "api://other" }, "GraphTokenIdentityMismatch"],
      ["a malformed token tenant", { graphTenantId: "not-a-tenant" }, "GraphTokenIdentityMismatch"],
    ])("SPFX-PREP-05: refuses %s before catalog discovery", async (_name, claims, code) => {
      await writeInputs();
      const { context } = createContext(claims);

      const result = await new PreparePublicationDriver().execute(args, context, outputNames);

      expect(errorCode(result.result)).toBe(code);
      expect(requests).toHaveLength(0);
      expect(await preparedDirectoryExists()).toBe(false);
    });

    it("SPFX-PREP-05: refuses preparation when the SharePoint token is unavailable", async () => {
      await writeInputs();
      const { context } = createContext();
      vi.mocked(Tools.getSPFxToken).mockResolvedValue(undefined);

      const result = await new PreparePublicationDriver().execute(args, context, outputNames);

      expect(errorCode(result.result)).toBe("SharePointAuthenticationUnavailable");
      expect(requests).toHaveLength(0);
    });

    it.each([
      ["no catalog", undefined],
      ["a foreign-host catalog", "https://other-tenant.example.test/sites/appcatalog"],
    ])("SPFX-PREP-05: refuses %s without writing a journal", async (_name, corporateCatalog) => {
      await writeInputs();
      route = async (config) => reply(config, 200, { CorporateCatalogUrl: corporateCatalog });
      const { context } = createContext();

      const result = await new PreparePublicationDriver().execute(args, context, outputNames);

      expect(errorCode(result.result)).toBe("TenantAppCatalogUnavailable");
      expect(posts()).toHaveLength(0);
      expect(await preparedDirectoryExists()).toBe(false);
    });

    it("SPFX-PREP-08: an already cancelled preparation makes no remote call", async () => {
      await writeInputs();
      const controller = new AbortController();
      controller.abort();
      const { context } = createContext({}, controller.signal);

      const result = await new PreparePublicationDriver().execute(args, context, outputNames);

      expect(errorCode(result.result)).toBe("PublicationCancelled");
      expect(requests).toHaveLength(0);
      expect(await preparedDirectoryExists()).toBe(false);
    });

    it.each(["Graph", "SharePoint"])(
      "SPFX-PREP-05: refuses preparation when %s token claims cannot be obtained",
      async (resource) => {
        await writeInputs();
        const { context } = createContext();
        const unavailable: Result<Record<string, unknown>, FxError> = err(
          new UserError({
            source: "fixture",
            name: "TokenUnavailable",
            message: "Claims unavailable.",
          })
        );
        const claims = vi.mocked(context.m365TokenProvider.getJsonObject);
        if (resource === "SharePoint") {
          claims.mockResolvedValueOnce(ok({ tid: tenantId, aud: "https://graph.microsoft.com" }));
        }
        claims.mockResolvedValueOnce(unavailable);
        const result = await new PreparePublicationDriver().execute(args, context, outputNames);
        expect(errorCode(result.result)).toBe(`${resource}TokenIdentityUnavailable`);
        expect(requests).toHaveLength(0);
        expect(await preparedDirectoryExists()).toBe(false);
      }
    );

    it("SPFX-PREP-05: accepts claims envelopes and resource application-ID audiences", async () => {
      await writeInputs();
      const { context } = createContext();
      vi.mocked(context.m365TokenProvider.getJsonObject)
        .mockResolvedValueOnce(
          ok({ value: { tid: tenantId, aud: "00000003-0000-0000-c000-000000000000" } })
        )
        .mockResolvedValueOnce(
          ok({ value: { tid: tenantId, aud: "00000003-0000-0ff1-ce00-000000000000" } })
        );
      route = catalogRoute({ deployed: false });
      const result = await new PreparePublicationDriver().execute(args, context, outputNames);
      expect(result.result.isErr() ? result.result.error : undefined).toBeUndefined();
      expect(posts()).toHaveLength(0);
    });

    it.each(["missing input", "invalid project"])(
      "SPFX-PREP-05: preserves the %s error at the preparation driver boundary",
      async (condition) => {
        await writeInputs();
        const { context } = createContext();
        if (condition === "invalid project")
          context.projectPath = path.join(projectPath ?? "", "missing");
        else await fs.remove(path.join(projectPath ?? "", args.packagePath));
        const result = await new PreparePublicationDriver().execute(args, context, outputNames);
        expect(errorCode(result.result)).toBe(
          condition === "invalid project" ? "ProjectPathUnavailable" : "InputReadFailed"
        );
        expect(requests).toHaveLength(0);
      }
    );

    it("SPFX-PREP-08: cancellation while obtaining the SharePoint token prevents catalog discovery", async () => {
      await writeInputs();
      const controller = new AbortController();
      const { context } = createContext({}, controller.signal);
      vi.mocked(Tools.getSPFxToken).mockImplementation(async () => {
        controller.abort();
        return "token";
      });
      const result = await new PreparePublicationDriver().execute(args, context, outputNames);
      expect(errorCode(result.result)).toBe("PublicationCancelled");
      expect(requests).toHaveLength(0);
    });
  });

  it("SPFX-PREP-05: refuses prepared deployment when the signed-in SharePoint origin differs from the journal", async () => {
    const { prepared } = await prepareProject();
    const otherOrigin = "https://other-tenant.example.test";
    route = async (config) =>
      config.url === "/_api/SP_TenantSettings_Current"
        ? reply(config, 200, { CorporateCatalogUrl: `${otherOrigin}/sites/appcatalog` })
        : reply(config, 500);
    const { context, approvals } = createContext({ sharePointAudience: otherOrigin });
    vi.mocked(SPFxDeployDriver.prototype.getTenant).mockResolvedValue(otherOrigin);

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(errorCode(result)).toBe("PreparedTenantMismatch");
    expect(approvals()).toBe(0);
    expect(posts()).toHaveLength(0);
    expect((await journalOf(prepared)).state).toBe("prepared");
  });

  it("SPFX-PREP-05: refuses application submission when the signed-in tenant changed after SharePoint verification", async () => {
    const { prepared } = await prepareProject();
    route = catalogRoute({ deployed: false });
    const deployed = await new SPFxDeployDriver().run(
      deployArgs(prepared),
      createContext().context
    );
    expect(deployed.isErr() ? deployed.error : undefined).toBeUndefined();
    requests.length = 0;
    vi.restoreAllMocks();
    const { context, approvals } = createContext({ graphTenantId: otherTenantId });

    const published = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);

    expect(published.result.isErr() ? published.result.error.name : undefined).toBe(
      "PreparedTenantMismatch"
    );
    expect(approvals()).toBe(0);
    expect(requests.map((request) => request.url)).toEqual(["/_api/SP_TenantSettings_Current"]);
    expect((await journalOf(prepared)).state).toBe("sharePointVerified");
  });

  it("SPFX-PREP-08: an active identity lock blocks another preparation, deployment and submission without remote calls", async () => {
    const { root, prepared } = await prepareProject();
    route = catalogRoute({ deployed: false });
    const journal = await journalOf(prepared);
    let releaseLock: () => void = () => undefined;
    let lockHeld: () => void = () => undefined;
    const held = new Promise<void>((resolve) => (lockHeld = resolve));
    const holder = withPreparedPublicationIdentityLock(root, journal, async () => {
      lockHeld();
      await new Promise<void>((resolve) => (releaseLock = resolve));
    });
    await held;

    const independent = await preparePublication({
      projectPath: root,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext: publicationTenantContext,
    });
    const { context, approvals } = createContext();
    const deployment = await new SPFxDeployDriver().run(deployArgs(prepared), context);
    const submission = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);
    releaseLock();
    await holder;

    expect(independent.isErr() ? independent.error.code : undefined).toBe(
      "PublicationIdentityLockUnavailable"
    );
    expect(errorCode(deployment)).toBe("PublicationIdentityLockUnavailable");
    expect(submission.result.isErr() ? submission.result.error.name : undefined).toBe(
      "PublicationIdentityLockUnavailable"
    );
    expect(approvals()).toBe(0);
    expect(requests).toHaveLength(0);

    const afterRelease = await new SPFxDeployDriver().run(deployArgs(prepared), context);
    expect(afterRelease.isErr() ? afterRelease.error : undefined).toBeUndefined();
    expect((await journalOf(prepared)).state).toBe("sharePointVerified");
  });

  it("SPFX-PREP-08: simultaneous preparations of changed sources for one identity create exactly one lineage", async () => {
    projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "prepared-sharepoint-contract-"));
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), createSppkg());
    await fs.writeFile(
      path.join(projectPath, "sharepoint", "solution-next.sppkg"),
      createSppkg("1.2.3.5")
    );
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), createAppPackage());
    const prepare = (packagePath: string) =>
      preparePublication({
        projectPath: projectPath ?? "",
        packagePath,
        appPackagePath: "appPackage/application.zip",
        tenantContext: publicationTenantContext,
      });

    const results = await Promise.all([
      prepare("sharepoint/solution.sppkg"),
      prepare("sharepoint/solution-next.sppkg"),
    ]);

    const succeeded = results.flatMap((result) => (result.isOk() ? [result.value] : []));
    const failed = results.flatMap((result) => (result.isErr() ? [result.error.code] : []));
    expect(succeeded).toHaveLength(1);
    expect(failed).toHaveLength(1);
    expect(["PublicationIdentityLockUnavailable", "UnresolvedPreparedPublication"]).toContain(
      failed[0]
    );
    const lineageRoot = path.dirname(
      path.dirname(path.join(projectPath ?? "", succeeded[0].preparedPublicationPath))
    );
    const lineages = await fs.readdir(lineageRoot, { withFileTypes: true });
    expect(
      lineages.filter((entry) => entry.isDirectory() && !entry.name.startsWith("_identity"))
    ).toHaveLength(1);
  });
  it("SPFX-PREP-08: an unknown upload blocks preparation of a rebuilt package with a new version", async () => {
    const { root, prepared } = await prepareProject();
    route = async (config) =>
      config.method === "post"
        ? Promise.reject(new AxiosError("socket hang up", AxiosError.ERR_NETWORK, config))
        : catalogRoute({ deployed: false })(config);
    const first = await new SPFxDeployDriver().run(deployArgs(prepared), createContext().context);
    expect(errorCode(first)).toBe("SharePointUploadOutcomeUnknown");
    await fs.writeFile(path.join(root, "sharepoint", "solution.sppkg"), createSppkg("1.2.3.5"));

    const rebuilt = await preparePublication({
      projectPath: root,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext: publicationTenantContext,
    });

    expect(rebuilt.isErr() ? rebuilt.error.code : undefined).toBe("UnresolvedPreparedPublication");
    expect((await journalOf(prepared)).state).toBe("unknown");
    expect(posts()).toHaveLength(1);
  });

  it("SPFX-PREP-08: cancellation during a catalog read before upload leaves the publication prepared", async () => {
    const { prepared } = await prepareProject();
    const controller = new AbortController();
    const catalog = catalogRoute({ deployed: false });
    route = async (config) => {
      if (config.url?.includes("getbytitle")) {
        controller.abort();
        throw new axios.CanceledError();
      }
      return catalog(config);
    };
    const { context, approvals } = createContext({}, controller.signal);

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(result.isErr() ? result.error : undefined).toBeInstanceOf(UserCancelError);
    expect(approvals()).toBe(0);
    expect(posts()).toHaveLength(0);
    expect((await journalOf(prepared)).state).toBe("prepared");
  });

  it("SPFX-PREP-08: cancellation between visibility polls keeps the confirmed upload and sends no deployment", async () => {
    const { prepared } = await prepareProject();
    const controller = new AbortController();
    const catalog: FakeCatalog = { deployed: false };
    const serve = catalogRoute(catalog);
    let hidePackage = true;
    route = async (config) => {
      if (hidePackage && catalog.packageBytes && config.url?.includes("getbytitle")) {
        return reply(config, 200, { value: [] });
      }
      return serve(config);
    };
    vi.mocked(sleep).mockImplementation(async () => {
      hidePackage = false;
      controller.abort();
    });
    const { context } = createContext({}, controller.signal);

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(result.isErr() ? result.error : undefined).toBeInstanceOf(UserCancelError);
    expect(sleep).toHaveBeenCalledTimes(1);
    expect(posts()).toHaveLength(1);
    const journal = await journalOf(prepared);
    expect(journal.state).toBe("sharePointUploaded");
    expect(journal.sharePointArtifact.uploadState).toBe("confirmed");
    expect(journal.sharePointArtifact.deployState).toBe("notStarted");
  });

  it.each([
    { outcome: "late success", succeed: true },
    { outcome: "late failure", succeed: false },
  ])(
    "SPFX-PREP-08: cancellation during the upload request retains the $outcome and sends no deployment",
    async ({ succeed }) => {
      const { prepared } = await prepareProject();
      const controller = new AbortController();
      const catalog = catalogRoute({ deployed: false });
      route = async (config) => {
        if (config.method === "post") {
          controller.abort();
          if (!succeed) {
            throw new AxiosError("socket hang up", AxiosError.ERR_NETWORK, config);
          }
        }
        return catalog(config);
      };
      const { context } = createContext({}, controller.signal);

      const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

      const journal = await journalOf(prepared);
      expect(posts()).toHaveLength(1);
      if (succeed) {
        expect(result.isErr() ? result.error : undefined).toBeInstanceOf(UserCancelError);
        expect(journal.state).toBe("sharePointUploaded");
        expect(journal.sharePointArtifact.uploadState).toBe("confirmed");
      } else {
        expect(errorCode(result)).toBe("SharePointUploadOutcomeUnknown");
        expect(journal.state).toBe("unknown");
        expect(journal.sharePointArtifact.uploadState).toBe("unknown");
      }
    }
  );

  it.each([
    { request: "upload", match: "tenantappcatalog/Add(", code: "SharePointUploadOutcomeUnknown" },
    { request: "deployment", match: "/Deploy", code: "SharePointDeployOutcomeUnknown" },
  ])(
    "SPFX-PREP-08: a $request request that exceeds the settlement bound is recorded as unknown",
    async ({ match, code }) => {
      const { prepared } = await prepareProject();
      const timeouts: { ms: number; controller: AbortController }[] = [];
      vi.spyOn(AbortSignal, "timeout").mockImplementation((ms) => {
        const controller = new AbortController();
        timeouts.push({ ms, controller });
        return controller.signal;
      });
      const catalog = catalogRoute({ deployed: false });
      route = async (config) => {
        if (config.method === "post" && config.url?.includes(match)) {
          const pending = hangUntilAborted(config);
          timeouts.find((entry) => entry.controller.signal === config.signal)?.controller.abort();
          return pending;
        }
        return catalog(config);
      };
      const { context } = createContext();

      const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

      expect(errorCode(result)).toBe(code);
      expect(timeouts.length).toBeGreaterThan(0);
      expect(timeouts.every((entry) => entry.ms === 30_000)).toBe(true);
      expect(posts().filter((request) => request.url?.includes(match))).toHaveLength(1);
      expect(posts().at(-1)?.url).toContain(match);
      expect((await journalOf(prepared)).state).toBe("unknown");
    }
  );

  it.each([
    {
      name: "a sovereign cloud",
      setup: () => vi.stubEnv("TEAMSFX_SOVEREIGN_CLOUD_ENVIRONMENT", "GCC H"),
      args: {},
      code: "UnsupportedCloud",
    },
    {
      name: "app catalog creation",
      setup: () => undefined,
      args: { createAppCatalogIfNotExist: true },
      code: "InvalidPreparedDeployArguments",
    },
    {
      name: "a different prepared package path",
      setup: () => undefined,
      args: { preparedPackagePath: "elsewhere/solution.sppkg" },
      code: "SharePointArtifactPathMismatch",
    },
  ])(
    "SPFX-PREP-06: refuses prepared deployment for $name before any remote call",
    async ({ setup, args, code }) => {
      const { prepared } = await prepareProject();
      route = catalogRoute({ deployed: false });
      setup();
      const { context, approvals } = createContext();

      const result = await new SPFxDeployDriver().run(
        { ...deployArgs(prepared), ...args },
        context
      );
      vi.unstubAllEnvs();

      expect(errorCode(result)).toBe(code);
      expect(approvals()).toBe(0);
      expect(requests).toHaveLength(0);
      expect((await journalOf(prepared)).state).toBe("prepared");
    }
  );

  it.each([
    {
      name: "the developer SharePoint package changed",
      file: "sharepoint/solution.sppkg",
      code: "SourceHashMismatch",
    },
    {
      name: "the developer app package changed",
      file: "appPackage/application.zip",
      code: "SourceHashMismatch",
    },
    {
      name: "the staged SharePoint copy changed",
      file: "sharePoint",
      code: "ArtifactHashMismatch",
    },
    { name: "the staged app package changed", file: "application", code: "ArtifactHashMismatch" },
  ])(
    "SPFX-PREP-06: refuses prepared deployment when $name after preparation",
    async ({ file, code }) => {
      const { root, prepared } = await prepareProject();
      const target =
        file === "sharePoint"
          ? prepared.sharePointPackagePath
          : file === "application"
            ? prepared.appPackagePath
            : file;
      await fs.appendFile(path.join(root, target), Buffer.from("tampered"));
      route = catalogRoute({ deployed: false });
      const { context, approvals } = createContext();

      const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

      expect(errorCode(result)).toBe(code);
      expect(approvals()).toBe(0);
      expect(requests).toHaveLength(0);
    }
  );

  it("SPFX-PREP-06: a verified publication replays deployment without remote calls", async () => {
    const { prepared } = await prepareProject();
    await verifySharePoint(prepared, { deployed: false });
    route = catalogRoute({ deployed: true });
    const { context, approvals } = createContext();

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(result.isErr() ? result.error : undefined).toBeUndefined();
    expect(approvals()).toBe(0);
    expect(requests).toHaveLength(0);
  });

  it("SPFX-PREP-08: an already cancelled invocation makes no remote call", async () => {
    const { prepared } = await prepareProject();
    route = catalogRoute({ deployed: false });
    const controller = new AbortController();
    controller.abort();
    const { context } = createContext({}, controller.signal);

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(result.isErr() ? result.error : undefined).toBeInstanceOf(UserCancelError);
    expect(requests).toHaveLength(0);
    expect((await journalOf(prepared)).state).toBe("prepared");
  });

  it("SPFX-PREP-06: fails clearly when the SharePoint token is unavailable", async () => {
    const { prepared } = await prepareProject();
    route = catalogRoute({ deployed: false });
    const { context } = createContext();
    vi.mocked(Tools.getSPFxToken).mockResolvedValue(undefined);

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(result.isErr()).toBe(true);
    expect(requests).toHaveLength(0);
    expect((await journalOf(prepared)).state).toBe("prepared");
  });

  it.each([
    { name: "no tenant app catalog", catalog: null },
    { name: "an app catalog on another host", catalog: "https://elsewhere.example.test/sites/a" },
  ])(
    "SPFX-PREP-06: refuses deployment with $name and never creates a catalog",
    async ({ catalog }) => {
      const { prepared } = await prepareProject();
      route = async (config) =>
        config.url === "/_api/SP_TenantSettings_Current"
          ? reply(config, 200, { CorporateCatalogUrl: catalog })
          : reply(config, 500);
      const { context, approvals } = createContext();

      const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

      expect(errorCode(result)).toBe("TenantAppCatalogUnavailable");
      expect(approvals()).toBe(0);
      expect(posts()).toHaveLength(0);
    }
  );

  it("SPFX-PREP-06: refuses deployment without an interactive approval surface", async () => {
    const { prepared } = await prepareProject();
    route = catalogRoute({ deployed: false });
    const { context } = createContext();

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), {
      ...context,
      ui: undefined,
    });

    expect(errorCode(result)).toBe("ApprovalUnavailable");
    expect(posts()).toHaveLength(0);
    expect((await journalOf(prepared)).state).toBe("prepared");
  });

  it.each([
    {
      name: "a different solution under the same file name",
      existing: () => {
        const other = new AdmZip(createSppkg("1.2.3.4", "other"));
        other.updateFile(
          "AppManifest.xml",
          Buffer.from(
            `<App ProductID="{99999999-2222-4333-8444-555555555555}" Version="1.2.3.4" IsClientSideSolution="true" SkipFeatureDeployment="false" />`
          )
        );
        return other.toBuffer();
      },
      version: undefined,
      code: "SharePointPackageNameCollision",
    },
    {
      name: "a newer catalog version",
      existing: () => createSppkg("2.0.0.0"),
      version: "2.0.0.0",
      code: "SharePointPackageDowngrade",
    },
    {
      name: "different content at the same version",
      existing: () => createSppkg("1.2.3.4", "other content"),
      version: undefined,
      code: "SharePointPackageVersionCollision",
    },
    {
      name: "a catalog version that disagrees with the stored file",
      existing: () => createSppkg("1.0.0.0"),
      version: "1.1.0.0",
      code: "SharePointCatalogVersionMismatch",
    },
  ])(
    "SPFX-PREP-06: refuses to replace $name without writing",
    async ({ existing, version, code }) => {
      const { prepared } = await prepareProject();
      route = catalogRoute({ deployed: true, packageBytes: existing(), version });
      const { context, approvals } = createContext();

      const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

      expect(errorCode(result)).toBe(code);
      expect(approvals()).toBe(0);
      expect(posts()).toHaveLength(0);
      expect((await journalOf(prepared)).state).toBe("prepared");
    }
  );

  it("SPFX-PREP-06: refuses an existing catalog solution that has no matching package file", async () => {
    const { prepared } = await prepareProject();
    const serve = catalogRoute({ deployed: true });
    route = async (config) =>
      config.url?.includes("AvailableApps/GetById") && config.method === "get"
        ? reply(config, 200, {
            ID: productId,
            AppCatalogVersion: "1.0.0.0",
            Deployed: true,
            CurrentVersionDeployed: true,
            IsPackageDefaultSkipFeatureDeployment: false,
          })
        : serve(config);
    const { context } = createContext();

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(errorCode(result)).toBe("SharePointProductIdCollision");
    expect(posts()).toHaveLength(0);
  });

  it("SPFX-PREP-06: upgrades an older catalog version with an overwrite upload", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = {
      deployed: true,
      packageBytes: createSppkg("1.0.0.0"),
      version: "1.0.0.0",
    };
    route = catalogRoute(catalog);
    const { context } = createContext();

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(result.isErr() ? result.error : undefined).toBeUndefined();
    expect(posts()[0].url).toBe(
      "/_api/web/tenantappcatalog/Add(overwrite=true, url='solution.sppkg')"
    );
    expect((await journalOf(prepared)).state).toBe("sharePointVerified");
  });

  it("SPFX-PREP-06: reuses identical deployed catalog content without any write or approval", async () => {
    const { root, prepared } = await prepareProject();
    const staged = await fs.readFile(path.join(root, prepared.sharePointPackagePath));
    route = catalogRoute({ deployed: true, packageBytes: staged });
    const { context, approvals } = createContext();

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(result.isErr() ? result.error : undefined).toBeUndefined();
    expect(approvals()).toBe(0);
    expect(posts()).toHaveLength(0);
    const journal = await journalOf(prepared);
    expect(journal.state).toBe("sharePointVerified");
    expect(journal.sharePointArtifact.deployState).toBe("confirmed");
  });

  it("SPFX-PREP-08: cancellation at deployment approval keeps the confirmed upload and sends no deployment", async () => {
    const { prepared } = await prepareProject();
    const controller = new AbortController();
    route = catalogRoute({ deployed: false });
    let approvalsSeen = 0;
    const { context } = createContext({}, controller.signal, () => {
      approvalsSeen += 1;
      if (approvalsSeen === 2) controller.abort();
    });

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(result.isErr() ? result.error : undefined).toBeInstanceOf(UserCancelError);
    expect(posts()).toHaveLength(1);
    const journal = await journalOf(prepared);
    expect(journal.state).toBe("sharePointUploaded");
    expect(journal.sharePointArtifact.deployState).toBe("notStarted");
  });

  it("SPFX-PREP-06: records a definitive deployment rejection without replay", async () => {
    const { prepared } = await prepareProject();
    const serve = catalogRoute({ deployed: false });
    route = async (config) =>
      config.method === "post" && config.url?.endsWith("/Deploy")
        ? reply(config, 403)
        : serve(config);
    const { context } = createContext();

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(errorCode(result)).toBe("SharePointDeployRejected");
    const journal = await journalOf(prepared);
    expect(journal.state).toBe("rejected");
    expect(journal.sharePointArtifact.deployState).toBe("rejected");

    const replay = await new SPFxDeployDriver().run(deployArgs(prepared), createContext().context);
    expect(errorCode(replay)).toBe("InvalidPreparedSharePointState");
    expect(posts()).toHaveLength(2);
  });

  it.each([
    {
      name: "the uploaded package never becomes readable",
      hide: (url: string) => url.includes("getbytitle"),
      code: "SharePointPackageNotVisible",
    },
    {
      name: "the uploaded solution never appears in the catalog",
      hide: (url: string) => url.includes("AvailableApps/GetById"),
      code: "SharePointAppNotVisible",
    },
  ])("SPFX-PREP-06: stops bounded polling when $name", async ({ hide, code }) => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    const serve = catalogRoute(catalog);
    route = async (config) =>
      catalog.packageBytes && config.method === "get" && hide(config.url ?? "")
        ? config.url?.includes("getbytitle")
          ? reply(config, 200, { value: [] })
          : reply(config, 404)
        : serve(config);
    const { context } = createContext();

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(errorCode(result)).toBe(code);
    expect(vi.mocked(sleep).mock.calls.length).toBeGreaterThan(0);
    expect(posts()).toHaveLength(1);
    expect((await journalOf(prepared)).sharePointArtifact.deployState).toBe("notStarted");
  });

  it("SPFX-PREP-06: does not mark SharePoint verified when the catalog never reports the deployment", async () => {
    const { prepared } = await prepareProject();
    route = catalogRoute({ deployed: false, ignoreDeploy: true });
    const { context } = createContext();

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(errorCode(result)).toBe("SharePointDeploymentNotVerified");
    const journal = await journalOf(prepared);
    expect(journal.state).toBe("sharePointUploaded");
    expect(journal.sharePointArtifact.deployState).toBe("confirmed");
  });

  it("SPFX-PREP-06: refuses a catalog entry whose version differs from the uploaded package", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    const serve = catalogRoute(catalog);
    route = async (config) => {
      const response = await serve(config);
      if (config.method === "post" && config.url?.includes("tenantappcatalog/Add(")) {
        catalog.version = "9.9.9.9";
      }
      return response;
    };
    const { context } = createContext();

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(errorCode(result)).toBe("SharePointCatalogIdentityMismatch");
    expect(posts().some((request) => request.url?.endsWith("/Deploy"))).toBe(false);
  });
  interface FakeGraph {
    existing?: { id: string; definitions: { version: string; publishingState: string }[] };
    submit: (config: InternalAxiosRequestConfig) => Promise<AxiosResponse>;
  }

  function graphAndCatalogRoute(catalog: FakeCatalog, graph: FakeGraph): Route {
    const serveCatalog = catalogRoute(catalog);
    return async (config) => {
      const url = config.url ?? "";
      if (url.startsWith("/appCatalogs/teamsApps") && config.method === "get") {
        return reply(config, 200, {
          value: graph.existing
            ? [
                {
                  id: graph.existing.id,
                  externalId: appId,
                  displayName: "Fixture",
                  appDefinitions: graph.existing.definitions,
                },
              ]
            : [],
        });
      }
      if (url.startsWith("/appCatalogs/teamsApps") && config.method === "post") {
        return graph.submit(config);
      }
      return serveCatalog(config);
    };
  }

  async function verifySharePoint(prepared: PreparedPublicationOutput, catalog: FakeCatalog) {
    route = catalogRoute(catalog);
    const deployed = await new SPFxDeployDriver().run(
      deployArgs(prepared),
      createContext().context
    );
    expect(deployed.isErr() ? deployed.error : undefined).toBeUndefined();
    requests.length = 0;
    vi.restoreAllMocks();
  }

  function graphPosts(): InternalAxiosRequestConfig[] {
    return posts().filter((request) => request.url?.startsWith("/appCatalogs/teamsApps"));
  }

  it("SPFX-PREP-07: submits the exact prepared app package once after SharePoint verification and replays the confirmed result", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    route = graphAndCatalogRoute(catalog, {
      submit: async (config) => reply(config, 201, { id: "teams-app-id", externalId: appId }),
    });
    const { context, approvals } = createContext();

    const published = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);
    const replayed = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);

    expect(published.result.isErr() ? published.result.error : undefined).toBeUndefined();
    expect(published.result.isOk() ? [...published.result.value.values()] : []).toEqual([
      "teams-app-id",
    ]);
    expect(replayed.result.isOk() ? [...replayed.result.value.values()] : []).toEqual([
      "teams-app-id",
    ]);
    expect(approvals()).toBe(1);
    expect(graphPosts()).toHaveLength(1);
    expect(graphPosts()[0].url).toBe("/appCatalogs/teamsApps?requiresReview=true");
    const submittedBytes = graphPosts()[0].data;
    expect(Buffer.isBuffer(submittedBytes)).toBe(true);
    expect(
      Buffer.compare(
        Buffer.isBuffer(submittedBytes) ? submittedBytes : Buffer.alloc(0),
        await fs.readFile(path.join(projectPath ?? "", prepared.appPackagePath))
      )
    ).toBe(0);
    const journal = await journalOf(prepared);
    expect(journal.state).toBe("submitted");
    expect(journal.applicationArtifact.submittedAppId).toBe("teams-app-id");
  });

  it.each([
    {
      outcome: "a definitive rejection",
      fail: (config: InternalAxiosRequestConfig) => reply(config, 400, { error: {} }),
      code: "TeamsAppSubmissionRejected",
      state: "rejected",
      httpStatus: 400,
    },
    {
      outcome: "an uncertain network failure",
      fail: (config: InternalAxiosRequestConfig) =>
        Promise.reject(new AxiosError("socket hang up", AxiosError.ERR_NETWORK, config)),
      code: "TeamsAppSubmissionOutcomeUnknown",
      state: "unknown",
      httpStatus: undefined,
    },
  ])(
    "SPFX-PREP-08: records $outcome of the app submission without retrying and keeps the verified SharePoint stage",
    async ({ fail, code, state, httpStatus }) => {
      const { prepared } = await prepareProject();
      const catalog: FakeCatalog = { deployed: false };
      await verifySharePoint(prepared, catalog);
      route = graphAndCatalogRoute(catalog, { submit: async (config) => fail(config) });
      const { context } = createContext();

      const published = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);
      const again = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);

      expect(published.result.isErr() ? published.result.error.name : undefined).toBe(code);
      const error = published.result.isErr() ? published.result.error : undefined;
      const data: unknown = error && "data" in error ? error.data : undefined;
      expect(data).toMatchObject({
        stage: "applicationSubmission",
        lastDurableState: state,
        lastSharePointVerifiedState: "sharePointVerified",
        httpStatus,
        appPackageSha256: prepared.appPackageSha256,
      });
      expect(JSON.stringify(data)).not.toContain("socket hang up");
      expect(again.result.isErr() ? again.result.error.name : undefined).toBe(
        "PreparedPublicationNotVerified"
      );
      expect(graphPosts()).toHaveLength(1);
      const journal = await journalOf(prepared);
      expect(journal.state).toBe(state);
      expect(journal.sharePointArtifact.deployState).toBe("confirmed");
    }
  );

  it.each([
    {
      name: "an existing app already at or above the prepared version",
      graph: { definitions: [{ version: "2.3.4", publishingState: "published" }] },
      deployed: true,
      code: "TeamsAppVersionCollision",
    },
    {
      name: "an existing app with only a pending definition",
      graph: { definitions: [{ version: "1.0.0", publishingState: "submitted" }] },
      deployed: true,
      code: "TeamsAppVersionCollision",
    },
    {
      name: "SharePoint no longer reporting the deployment",
      graph: undefined,
      deployed: false,
      code: "SharePointDeploymentNoLongerVerified",
    },
  ])(
    "SPFX-PREP-07: refuses app submission with $name before approval or write",
    async ({ graph, deployed, code }) => {
      const { prepared } = await prepareProject();
      const catalog: FakeCatalog = { deployed: false };
      await verifySharePoint(prepared, catalog);
      catalog.deployed = deployed;
      route = graphAndCatalogRoute(catalog, {
        existing: graph ? { id: "teams-app-id", definitions: graph.definitions } : undefined,
        submit: async (config) => reply(config, 500),
      });
      const { context, approvals } = createContext();

      const published = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);

      expect(published.result.isErr() ? published.result.error.name : undefined).toBe(code);
      expect(approvals()).toBe(0);
      expect(posts()).toHaveLength(0);
      expect((await journalOf(prepared)).state).toBe("sharePointVerified");
    }
  );

  it("SPFX-PREP-07: updates an existing published app through a new definition", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    route = graphAndCatalogRoute(catalog, {
      existing: {
        id: "existing-app-id",
        definitions: [{ version: "1.0.0", publishingState: "published" }],
      },
      submit: async (config) => reply(config, 201, {}),
    });

    const published = await new PublishAppPackageDriver().execute(
      publishArgs(prepared),
      createContext().context
    );

    expect(published.result.isOk() ? [...published.result.value.values()] : []).toEqual([
      "existing-app-id",
    ]);
    expect(graphPosts().map((request) => request.url)).toEqual([
      "/appCatalogs/teamsApps/existing-app-id/appDefinitions?requiresReview=true",
    ]);
  });

  it("SPFX-PREP-08: cancellation after submission approval restores the verified stage without a Graph write", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    const controller = new AbortController();
    route = graphAndCatalogRoute(catalog, {
      submit: async (config) => reply(config, 201, { id: "teams-app-id", externalId: appId }),
    });
    const { context } = createContext({}, controller.signal, () => controller.abort());

    const published = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);

    expect(published.result.isErr() ? published.result.error : undefined).toBeInstanceOf(
      UserCancelError
    );
    expect(graphPosts()).toHaveLength(0);
    const journal = await journalOf(prepared);
    expect(journal.state).toBe("sharePointVerified");
    expect(journal.applicationArtifact.submissionState).toBe("notStarted");
  });

  it("SPFX-PREP-07: refuses submission when YAML inputs no longer match the prepared journal", async () => {
    const { prepared } = await prepareProject();
    route = catalogRoute({ deployed: false });
    const { context } = createContext();

    const mismatched = await new PublishAppPackageDriver().execute(
      { ...publishArgs(prepared), appPackageSha256: "0".repeat(64) },
      context
    );
    const incomplete = await new PublishAppPackageDriver().execute(
      { ...publishArgs(prepared), pluginSha256: "not-a-hash" },
      context
    );
    const unverified = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);

    expect(mismatched.result.isErr() ? mismatched.result.error.name : undefined).toBe(
      "PreparedPublicationNotVerified"
    );
    expect(incomplete.result.isErr() ? incomplete.result.error.name : undefined).toBe(
      "InvalidPreparedPublishArguments"
    );
    expect(unverified.result.isErr() ? unverified.result.error.name : undefined).toBe(
      "PreparedPublicationNotVerified"
    );
    expect(requests).toHaveLength(0);
  });
  function noSubmission(): FakeGraph {
    return {
      submit: async () => {
        throw new Error("No application submission is expected.");
      },
    };
  }

  async function expectNoSubmission(
    prepared: PreparedPublicationOutput,
    result: Result<unknown, FxError>,
    code: string
  ): Promise<void> {
    expect(result.isErr() ? result.error.name : undefined).toBe(code);
    expect(graphPosts()).toHaveLength(0);
    const journal = await journalOf(prepared);
    expect(journal.state).toBe("sharePointVerified");
    expect(journal.applicationArtifact.submissionState).toBe("notStarted");
  }

  it.each([
    { name: "staged app package", artifact: "application" },
    { name: "staged SharePoint copy", artifact: "sharePoint" },
  ])(
    "SPFX-PREP-07: refuses submission when the $name changed after SharePoint verification",
    async ({ artifact }) => {
      const { root, prepared } = await prepareProject();
      const catalog: FakeCatalog = { deployed: false };
      await verifySharePoint(prepared, catalog);
      await fs.appendFile(
        path.join(
          root,
          artifact === "application" ? prepared.appPackagePath : prepared.sharePointPackagePath
        ),
        Buffer.from("tampered")
      );
      route = graphAndCatalogRoute(catalog, noSubmission());
      const { context } = createContext();

      const published = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);

      await expectNoSubmission(prepared, published.result, "ArtifactHashMismatch");
      expect(requests).toHaveLength(0);
    }
  );

  it("SPFX-PREP-07: refuses submission when the tenant app catalog moved after verification", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    const serve = graphAndCatalogRoute(catalog, noSubmission());
    route = async (config) =>
      config.url === "/_api/SP_TenantSettings_Current"
        ? reply(config, 200, { CorporateCatalogUrl: `${tenantUrl}/sites/othercatalog` })
        : serve(config);
    const { context } = createContext();

    const published = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);

    await expectNoSubmission(prepared, published.result, "SharePointCatalogChanged");
  });

  it.each([
    { name: "has been removed", bytes: () => undefined, code: "SharePointPackageNotVisible" },
    {
      name: "contains different content",
      bytes: () => createSppkg("1.2.3.4", "replaced"),
      code: "SharePointArtifactHashMismatch",
    },
  ])("SPFX-PREP-07: refuses submission when the catalog package $name", async ({ bytes, code }) => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    catalog.packageBytes = bytes();
    route = graphAndCatalogRoute(catalog, noSubmission());
    const { context } = createContext();

    const published = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);

    await expectNoSubmission(prepared, published.result, code);
  });

  it("SPFX-PREP-07: refuses submission without an interactive approval surface", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    route = graphAndCatalogRoute(catalog, noSubmission());
    const { context } = createContext();

    const published = await new PublishAppPackageDriver().execute(publishArgs(prepared), {
      ...context,
      ui: undefined,
    });

    await expectNoSubmission(prepared, published.result, "ApprovalUnavailable");
  });

  it("SPFX-PREP-07: a declined submission approval sends no Graph write", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    route = graphAndCatalogRoute(catalog, noSubmission());
    const { context } = createContext();
    vi.mocked(context.ui!.showMessage).mockResolvedValue(ok("Cancel"));

    const published = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);

    expect(published.result.isErr() ? published.result.error : undefined).toBeInstanceOf(
      UserCancelError
    );
    expect(graphPosts()).toHaveLength(0);
    expect((await journalOf(prepared)).state).toBe("sharePointVerified");
  });

  it("SPFX-PREP-08: an already cancelled submission makes no remote call", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    route = graphAndCatalogRoute(catalog, noSubmission());
    const controller = new AbortController();
    controller.abort();
    const { context } = createContext({}, controller.signal);

    const published = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);

    expect(published.result.isErr() ? published.result.error : undefined).toBeInstanceOf(
      UserCancelError
    );
    expect(requests).toHaveLength(0);
  });

  it.each([
    { name: "SharePoint", code: "SharePointAuthenticationUnavailable" },
    { name: "Graph", code: "GraphTokenUnavailable" },
  ])(
    "SPFX-PREP-07: refuses submission when the $name token is unavailable",
    async ({ name, code }) => {
      const { prepared } = await prepareProject();
      const catalog: FakeCatalog = { deployed: false };
      await verifySharePoint(prepared, catalog);
      route = graphAndCatalogRoute(catalog, noSubmission());
      const { context } = createContext();
      if (name === "SharePoint") {
        vi.mocked(Tools.getSPFxToken).mockResolvedValue(undefined);
      } else {
        vi.spyOn(context.m365TokenProvider, "getAccessToken").mockResolvedValue(
          err(new UserError("test", code, "no token"))
        );
      }

      const published = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);

      await expectNoSubmission(prepared, published.result, code);
      if (name === "Graph") {
        const error = published.result.isErr() ? published.result.error : undefined;
        expect(error).toBeInstanceOf(UserError);
        expect(error && "data" in error ? error.data : undefined).toMatchObject({
          stage: "sharePointVerification",
          lastSharePointVerifiedState: "sharePointVerified",
          publicationId: prepared.publicationId,
        });
      }
    }
  );

  it("SPFX-PREP-07: refuses prepared submission in a sovereign cloud before any remote call", async () => {
    const { prepared } = await prepareProject();
    route = catalogRoute({ deployed: false });
    const { context } = createContext();
    vi.stubEnv("TEAMSFX_SOVEREIGN_CLOUD_ENVIRONMENT", "DoD");

    const published = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);
    vi.unstubAllEnvs();

    expect(published.result.isErr() ? published.result.error.name : undefined).toBe(
      "UnsupportedCloud"
    );
    expect(requests).toHaveLength(0);
  });
  describe("prepared artifact integrity", () => {
    const hash = (bytes: Buffer) => createHash("sha256").update(bytes).digest("hex");

    async function journalFile(prepared: PreparedPublicationOutput): Promise<string> {
      return path.join(projectPath ?? "", ...prepared.preparedPublicationPath.split("/"));
    }

    it.each([
      [
        "an absolute path",
        (p: PreparedPublicationOutput) => path.resolve(p.preparedPublicationPath),
      ],
      ["a traversal path", () => "../journal.json"],
      ["another publication's path", () => ".fx/sharepoint-publications/x/journal.json"],
    ])("SPFX-PREP-08: refuses %s as the journal location", async (_name, journalPath) => {
      const { prepared } = await prepareProject();

      const loaded = await loadPreparedPublication(
        projectPath ?? "",
        prepared.publicationId,
        journalPath(prepared)
      );

      expect(loaded.isErr() && loaded.error.code).toBe("InvalidJournalPath");
    });

    it.each([
      ["malformed JSON", "{", "JournalReadFailed"],
      ["an unsupported shape", "{}", "InvalidJournal"],
    ])("SPFX-PREP-08: refuses a journal with %s", async (_name, content, code) => {
      const { prepared } = await prepareProject();
      await fs.writeFile(await journalFile(prepared), content);

      const loaded = await loadPreparedPublication(
        projectPath ?? "",
        prepared.publicationId,
        prepared.preparedPublicationPath
      );

      expect(loaded.isErr() && loaded.error.code).toBe(code);
    });

    it("SPFX-PREP-08: refuses a missing journal and an unresolvable project", async () => {
      const { prepared } = await prepareProject();
      await fs.remove(await journalFile(prepared));

      const missing = await loadPreparedPublication(
        projectPath ?? "",
        prepared.publicationId,
        prepared.preparedPublicationPath
      );
      const noProject = await loadPreparedPublication(
        path.join(projectPath ?? "", "missing"),
        prepared.publicationId,
        prepared.preparedPublicationPath
      );

      expect(missing.isErr() && missing.error.code).toBe("JournalReadFailed");
      expect(noProject.isErr() && noProject.error.code).toBe("ProjectPathUnavailable");
    });

    it("SPFX-PREP-06: refuses artifact paths outside the project and missing artifacts", async () => {
      const { root, prepared } = await prepareProject();

      const escaped = await verifyPreparedArtifact(root, {
        path: "../outside.zip",
        sha256: prepared.appPackageSha256,
      });
      await fs.remove(path.join(root, ...prepared.appPackagePath.split("/")));
      const missing = await verifyPreparedArtifact(root, {
        path: prepared.appPackagePath,
        sha256: prepared.appPackageSha256,
      });

      expect(escaped.isErr() && escaped.error.code).toBe("InvalidArtifactPath");
      expect(missing.isErr() && missing.error.code).toBe("ArtifactReadFailed");
    });

    it("SPFX-PREP-06: refuses source paths outside the project and deleted sources", async () => {
      const { root, prepared } = await prepareProject();
      const journal = await journalOf(prepared);

      const escaped = await verifyPreparedSources(root, {
        ...journal,
        sourceSharePoint: { ...journal.sourceSharePoint, path: "../solution.sppkg" },
      });
      await fs.remove(path.join(root, "sharepoint", "solution.sppkg"));
      const deleted = await verifyPreparedSources(root, journal);

      expect(escaped.isErr() && escaped.error.code).toBe("InvalidSourcePath");
      expect(deleted.isErr() && deleted.error.code).toBe("SourceReadFailed");
    });

    it("SPFX-PREP-06: refuses a SharePoint copy whose identity differs from its prepared source", async () => {
      const { root, prepared } = await prepareProject();
      const journal = await journalOf(prepared);
      const bytes = await fs.readFile(
        path.join(root, ...prepared.sharePointPackagePath.split("/"))
      );

      const result = await verifyPreparedSharePointPackage(
        { ...journal, sourceSharePoint: { ...journal.sourceSharePoint, version: "9.9.9.9" } },
        bytes
      );

      expect(result.isErr() && result.error.code).toBe("SharePointArtifactIdentityMismatch");
    });

    it.each([
      [
        "missing its manifest",
        (zip: AdmZip) => zip.deleteFile("manifest.json"),
        "InvalidApplicationPackage",
      ],
      [
        "with a changed app identity",
        (zip: AdmZip) => {
          const manifest = JSON.parse(zip.readAsText("manifest.json"));
          zip.updateFile(
            "manifest.json",
            Buffer.from(JSON.stringify({ ...manifest, version: "9.9.9" }))
          );
        },
        "ApplicationArtifactIdentityMismatch",
      ],
      [
        "missing its declarative agent",
        (zip: AdmZip) => {
          const manifest = JSON.parse(zip.readAsText("manifest.json"));
          zip.deleteFile(manifest.copilotAgents.declarativeAgents[0].file);
        },
        "MissingAgentManifest",
      ],
      [
        "without the prepared plugin",
        (zip: AdmZip) => {
          const plugin = zip.getEntries().find((entry) => entry.entryName.startsWith("plugins/"));
          if (plugin) zip.deleteFile(plugin.entryName);
        },
        "PluginArtifactLineageMismatch",
      ],
      [
        "with a duplicated plugin action",
        (zip: AdmZip) => {
          const manifest = JSON.parse(zip.readAsText("manifest.json"));
          const agentFile: string = manifest.copilotAgents.declarativeAgents[0].file;
          const agent = JSON.parse(zip.readAsText(agentFile));
          agent.actions = [...agent.actions, ...agent.actions];
          zip.updateFile(agentFile, Buffer.from(JSON.stringify(agent)));
        },
        "PluginArtifactLineageMismatch",
      ],
    ])("SPFX-PREP-07: refuses a rehashed application package %s", async (_name, tamper, code) => {
      const { root, prepared } = await prepareProject();
      const journal = await journalOf(prepared);
      const zip = new AdmZip(
        await fs.readFile(path.join(root, ...prepared.appPackagePath.split("/")))
      );
      tamper(zip);
      const bytes = zip.toBuffer();

      const result = verifyPreparedApplicationPackage(
        {
          ...journal,
          applicationArtifact: { ...journal.applicationArtifact, sha256: hash(bytes) },
        },
        bytes
      );

      expect(result.isErr() && result.error.code).toBe(code);
    });
  });

  describe("durable mutation boundaries", () => {
    async function externalUpdate(
      prepared: PreparedPublicationOutput,
      change: (journal: PreparedPublicationJournal) => void
    ): Promise<void> {
      const journal = await journalOf(prepared);
      change(journal);
      await fs.writeJson(path.join(projectPath ?? "", prepared.preparedPublicationPath), journal);
    }

    describe.each(["deployment", "submission"])("%s preflight", (operation) => {
      it.each([
        ["missing journal", "JournalReadFailed"],
        ["changed source", "SourceHashMismatch"],
        ["SharePoint identity", "SharePointArtifactIdentityMismatch"],
        ["application identity", "ApplicationArtifactIdentityMismatch"],
      ])("SPFX-PREP-06: refuses %s before approval or network writes", async (failure, code) => {
        const { root, prepared } = await prepareProject();
        const catalog: FakeCatalog = { deployed: false };
        if (operation === "submission") await verifySharePoint(prepared, catalog);
        if (failure === "missing journal") {
          await fs.remove(path.join(root, prepared.preparedPublicationPath));
        } else if (failure === "changed source") {
          await fs.appendFile(path.join(root, "sharepoint", "solution.sppkg"), "changed");
        } else {
          await externalUpdate(prepared, (journal) => {
            if (failure === "SharePoint identity") journal.sourceSharePoint.version = "9.9.9.9";
            else journal.applicationArtifact.version = "9.9.9";
          });
        }
        route = catalogRoute(catalog);
        const { context, approvals } = createContext();
        const result =
          operation === "submission"
            ? (await new PublishAppPackageDriver().execute(publishArgs(prepared), context)).result
            : await new SPFxDeployDriver().run(deployArgs(prepared), context);
        expect(
          result.isErr()
            ? operation === "submission"
              ? result.error.name
              : result.error.message
            : undefined
        ).toBe(code);
        expect(approvals()).toBe(0);
        expect(requests).toHaveLength(0);
      });
    });

    it("SPFX-PREP-06: deployment refuses validly formatted but unbound YAML hashes", async () => {
      const { prepared } = await prepareProject();
      const { context, approvals } = createContext();
      const result = await new SPFxDeployDriver().run(
        {
          ...deployArgs(prepared),
          sourcePackageSha256: "0".repeat(64),
        },
        context
      );
      expect(errorCode(result)).toBe("PreparedPublicationHashBindingMismatch");
      expect(approvals()).toBe(0);
      expect(requests).toHaveLength(0);
    });

    it("SPFX-PREP-08: a prepared journal cannot replay a previously confirmed upload", async () => {
      const { prepared } = await prepareProject();
      await externalUpdate(prepared, (journal) => {
        journal.sharePointArtifact.uploadState = "confirmed";
      });
      route = catalogRoute({ deployed: false });
      const { context, approvals } = createContext();
      const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);
      expect(errorCode(result)).toBe("InvalidPreparedUploadState");
      expect(approvals()).toBe(0);
      expect(posts()).toHaveLength(0);
    });

    it.each(["upload", "deployment"])(
      "SPFX-PREP-11: preserves allowlisted SharePoint %s rejection diagnostics",
      async (stage) => {
        const { prepared } = await prepareProject();
        const serve = catalogRoute({ deployed: false });
        route = async (config) => {
          const matches =
            config.method === "post" &&
            (stage === "upload"
              ? config.url?.includes("tenantappcatalog/Add(")
              : config.url?.endsWith("/Deploy"));
          return matches
            ? reply(
                config,
                403,
                { error: { code: "AccessDenied", message: "do not expose this response" } },
                { "request-id": `sp-${stage}-request` }
              )
            : serve(config);
        };
        const { context } = createContext();
        const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);
        expect(errorCode(result)).toBe(
          stage === "upload" ? "SharePointUploadRejected" : "SharePointDeployRejected"
        );
        if (result.isErr()) {
          expect(result.error.data).toMatchObject({
            stage: stage === "upload" ? "sharePointUpload" : "sharePointDeploy",
            lastDurableState: "rejected",
            httpStatus: 403,
            serviceCode: "AccessDenied",
            requestId: `sp-${stage}-request`,
          });
          expect(result.error.data.lastSharePointVerifiedState).toBeUndefined();
          expect(JSON.stringify(result.error.data)).not.toContain("do not expose this response");
          expect(result.error.data).not.toHaveProperty("response");
        }
      }
    );

    const verificationFields: ("catalogUrl" | "fileName")[] = ["catalogUrl", "fileName"];
    it.each(verificationFields)(
      "SPFX-PREP-07: submission refuses missing SharePoint %s evidence without any request",
      async (field) => {
        const { prepared } = await prepareProject();
        await verifySharePoint(prepared, { deployed: false });
        await externalUpdate(prepared, (journal) => {
          delete journal.sharePointArtifact[field];
        });
        const { context, approvals } = createContext();
        const result = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);
        expect(result.result.isErr() ? result.result.error.name : undefined).toBe(
          "MissingSharePointVerificationData"
        );
        expect(approvals()).toBe(0);
        expect(requests).toHaveLength(0);
        expect((await journalOf(prepared)).state).toBe("sharePointVerified");
      }
    );

    it.each(["upload", "deployment", "submission"])(
      "SPFX-PREP-08: cancellation immediately after the %s intent is committed restores only the known pre-write state",
      async (stage) => {
        const { prepared } = await prepareProject();
        const catalog: FakeCatalog = { deployed: false };
        if (stage === "submission") await verifySharePoint(prepared, catalog);
        const controller = new AbortController();
        const rename = fs.rename.bind(fs);
        const targetState =
          stage === "upload"
            ? "sharePointUploadInFlight"
            : stage === "deployment"
              ? "sharePointDeployInFlight"
              : "appSubmissionInFlight";
        vi.spyOn(fs, "rename").mockImplementation(async (from, to) => {
          await rename(from, to);
          if (path.basename(String(to)) === "journal.json") {
            const journal = await journalOf(prepared);
            if (journal.state === targetState) controller.abort();
          }
        });
        route = graphAndCatalogRoute(catalog, {
          submit: async () => {
            throw new Error("Submission must not start.");
          },
        });
        const { context } = createContext({}, controller.signal);
        const result =
          stage === "submission"
            ? (await new PublishAppPackageDriver().execute(publishArgs(prepared), context)).result
            : await new SPFxDeployDriver().run(deployArgs(prepared), context);
        expect(result.isErr() ? result.error : undefined).toBeInstanceOf(UserCancelError);
        const journal = await journalOf(prepared);
        expect(journal.state).toBe(
          stage === "upload"
            ? "prepared"
            : stage === "deployment"
              ? "sharePointUploaded"
              : "sharePointVerified"
        );
        expect(journal.sharePointArtifact.uploadState).toBe(
          stage === "upload" ? "notStarted" : "confirmed"
        );
        expect(journal.applicationArtifact.submissionState).toBe("notStarted");
        expect(posts()).toHaveLength(stage === "deployment" ? 1 : 0);
      }
    );

    it.each([
      { stage: "upload approval", boundary: "approval", mutation: "upload", failure: false },
      { stage: "upload response", boundary: "response", mutation: "upload", failure: false },
      { stage: "upload failure", boundary: "response", mutation: "upload", failure: true },
      {
        stage: "deployment approval",
        boundary: "approval",
        mutation: "deployment",
        failure: false,
      },
      {
        stage: "deployment response",
        boundary: "response",
        mutation: "deployment",
        failure: false,
      },
      { stage: "deployment failure", boundary: "response", mutation: "deployment", failure: true },
      {
        stage: "final verification",
        boundary: "verification",
        mutation: "deployment",
        failure: false,
      },
    ])(
      "SPFX-PREP-08: an external state change at $stage is not overwritten or followed by another write",
      async ({ boundary, mutation, failure }) => {
        const { prepared } = await prepareProject();
        const catalog: FakeCatalog = { deployed: false };
        const serve = catalogRoute(catalog);
        let changed = false;
        const change = async () => {
          changed = true;
          await externalUpdate(prepared, (journal) => {
            journal.state = "unknown";
          });
        };
        route = async (config) => {
          const matches =
            config.method === "post" &&
            (mutation === "upload"
              ? config.url?.includes("tenantappcatalog/Add(")
              : config.url?.endsWith("/Deploy"));
          if (boundary === "response" && matches) {
            await change();
            if (failure) return reply(config, 400, { error: { code: "Rejected" } });
          }
          const result = await serve(config);
          if (
            boundary === "verification" &&
            !changed &&
            catalog.deployed &&
            config.url?.includes("AvailableApps")
          ) {
            await change();
          }
          return result;
        };
        const { context } = createContext();
        if (!context.ui) throw new Error("Missing fixture UI.");
        let approvals = 0;
        vi.mocked(context.ui.showMessage).mockImplementation(async (_level, _message, modal) => {
          if (modal) {
            approvals += 1;
            if (boundary === "approval" && approvals === (mutation === "upload" ? 1 : 2))
              await change();
          }
          return ok("Confirm");
        });
        const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);
        expect(changed).toBe(true);
        expect(errorCode(result)).toBe("JournalUpdateFailed");
        expect((await journalOf(prepared)).state).toBe("unknown");
        const expectedPosts =
          mutation === "upload"
            ? boundary === "approval"
              ? 0
              : 1
            : boundary === "approval"
              ? 1
              : 2;
        expect(posts()).toHaveLength(expectedPosts);
        expect(graphPosts()).toHaveLength(0);
      }
    );

    it.each(["approval", "success", "rejection"])(
      "SPFX-PREP-08: an external submission-state change at %s cannot be overwritten",
      async (boundary) => {
        const { prepared } = await prepareProject();
        const catalog: FakeCatalog = { deployed: false };
        await verifySharePoint(prepared, catalog);
        const change = async () =>
          externalUpdate(prepared, (journal) => {
            journal.applicationArtifact.submissionState = "unknown";
          });
        route = graphAndCatalogRoute(catalog, {
          submit: async (config) => {
            await change();
            return reply(config, boundary === "rejection" ? 400 : 201, {
              id: "catalog-id",
              externalId: appId,
            });
          },
        });
        const { context } = createContext();
        if (!context.ui) throw new Error("Missing fixture UI.");
        if (boundary === "approval") {
          vi.mocked(context.ui.showMessage).mockImplementation(async () => {
            await change();
            return ok("Confirm");
          });
        }
        const result = await new PublishAppPackageDriver().execute(publishArgs(prepared), context);
        expect(result.result.isErr() ? result.result.error.name : undefined).toBe(
          "JournalUpdateFailed"
        );
        expect((await journalOf(prepared)).applicationArtifact.submissionState).toBe("unknown");
        expect(graphPosts()).toHaveLength(boundary === "approval" ? 0 : 1);
      }
    );

    it.each([
      { state: "sharePointUploadInFlight", expected: "prepared", writes: 0 },
      { state: "sharePointUploaded", expected: "sharePointUploadInFlight", writes: 1 },
      { state: "sharePointDeployInFlight", expected: "sharePointUploaded", writes: 1 },
      { state: "sharePointVerified", expected: "sharePointUploaded", writes: 2 },
    ])(
      "SPFX-PREP-08: failure to persist $state retains $expected and stops subsequent writes",
      async ({ state, expected, writes }) => {
        const { prepared } = await prepareProject();
        const rename = fs.rename.bind(fs);
        vi.spyOn(fs, "rename").mockImplementation(async (from, to) => {
          if (path.basename(String(to)) === "journal.json") {
            const pending = await fs.readJson(String(from));
            if (pending.state === state) throw new Error("Journal storage unavailable.");
          }
          await rename(from, to);
        });
        route = catalogRoute({ deployed: false });
        const result = await new SPFxDeployDriver().run(
          deployArgs(prepared),
          createContext().context
        );
        expect(errorCode(result)).toBe("JournalUpdateFailed");
        expect((await journalOf(prepared)).state).toBe(expected);
        expect(posts()).toHaveLength(writes);
      }
    );
  });

  it("keeps the fixture tenant identifier aligned with the publication context", () => {
    expect(publicationTenantContext.tenantId).toBe(tenantId);
  });
});
