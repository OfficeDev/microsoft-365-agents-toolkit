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

import * as commonUtils from "../../../../../src/common/utils";
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
import { CopilotAgentPublishDriver } from "../../../../../src/component/driver/copilotAgent/publish";
import { UserCancelError } from "../../../../../src/error/common";
import { MockTools } from "../../../../core/utils";
import {
  createAppPackage,
  createSppkg,
  mockTokenClaims,
  multipartPart,
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
    request: { method: config.method?.toUpperCase() },
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
    vi.unstubAllEnvs();
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
    vi.spyOn(commonUtils, "waitSeconds").mockResolvedValue();
    vi.stubEnv("TEAMSFX_RAI_VALIDATION_ENABLED", "false");
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

  function publishArgs(prepared: PreparedPublicationOutput, scope = "tenant") {
    return {
      appPackagePath: prepared.appPackagePath,
      scope,
      preparedPublicationPath: prepared.preparedPublicationPath,
      publicationId: prepared.publicationId,
      sourcePackageSha256: prepared.sourcePackageSha256,
      sourceAppPackageSha256: prepared.sourceAppPackageSha256,
      pluginSha256: prepared.pluginSha256,
      sharePointPackageSha256: prepared.sharePointPackageSha256,
      appPackageSha256: prepared.appPackageSha256,
    };
  }

  const publishOutputNames = new Map([
    ["titleId", "M365_TITLE_ID"],
    ["appId", "M365_APP_ID"],
    ["shareLink", "M365_SHARE_LINK"],
  ]);

  function publish(
    prepared: PreparedPublicationOutput,
    context: DriverContext,
    scope?: string,
    outputs: Map<string, string> = publishOutputNames
  ) {
    return new CopilotAgentPublishDriver().execute(publishArgs(prepared, scope), context, outputs);
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
    const { context, approvals } = createContext({ sharePointTenantId: otherTenantId });

    const published = await publish(prepared, context);

    expect(published.result.isErr() ? published.result.error.name : undefined).toBe(
      "SharePointTokenIdentityMismatch"
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
    const submission = await publish(prepared, context);
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
      replacement: () => createAppPackage("2.3.5"),
    },
    {
      name: "the staged SharePoint copy changed",
      file: "sharePoint",
      code: "ArtifactHashMismatch",
    },
    { name: "the staged app package changed", file: "application", code: "ArtifactHashMismatch" },
  ])(
    "SPFX-PREP-06: refuses prepared deployment when $name after preparation",
    async ({ file, code, replacement }) => {
      const { root, prepared } = await prepareProject();
      const target =
        file === "sharePoint"
          ? prepared.sharePointPackagePath
          : file === "application"
            ? prepared.appPackagePath
            : file;
      if (replacement) {
        await fs.writeFile(path.join(root, target), replacement());
      } else {
        await fs.appendFile(path.join(root, target), Buffer.from("tampered"));
      }
      route = catalogRoute({ deployed: false });
      const { context, approvals } = createContext();

      const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

      expect(errorCode(result)).toBe(code);
      expect(approvals()).toBe(0);
      expect(requests).toHaveLength(0);
    }
  );

  it("SPFX-PREP-21: a verified publication replays deployment after the developer app package is rebuilt with identical content", async () => {
    const { root, prepared } = await prepareProject();
    await verifySharePoint(prepared, { deployed: false });
    const applicationPath = path.join(root, "appPackage", "application.zip");
    const rebuilt = createAppPackage("2.3.4", {
      entryTime: new Date(2021, 5, 15, 12, 30, 40),
      reverseOrder: true,
    });
    expect(rebuilt.equals(await fs.readFile(applicationPath))).toBe(false);
    await fs.writeFile(applicationPath, rebuilt);
    route = catalogRoute({ deployed: true });
    const { context, approvals } = createContext();

    const result = await new SPFxDeployDriver().run(deployArgs(prepared), context);

    expect(result.isErr() ? result.error : undefined).toBeUndefined();
    expect(approvals()).toBe(0);
    expect(requests).toHaveLength(0);
  });

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

  const mosTitleId = "T_fixture-title";
  const mosAppId = "fixture-mos-app";
  const mosShareLink = "https://share.example.test/apps/fixture";

  interface FakeMos {
    environment?: Route;
    upload?: Route;
    status?: Route;
    sharingInfo?: Route;
  }

  /** Serves the MOS3 Builder API used by prepared submission next to the SharePoint catalog. */
  function mosAndCatalogRoute(catalog: FakeCatalog, mos: FakeMos = {}): Route {
    const serveCatalog = catalogRoute(catalog);
    return async (config) => {
      const url = config.url ?? "";
      if (config.method === "get" && url === "/config/v1/environment") {
        return mos.environment
          ? mos.environment(config)
          : reply(config, 200, { titlesServiceUrl: "https://titles.example.test" });
      }
      if (config.method === "post" && url === "/builder/v1/users/packages") {
        return mos.upload ? mos.upload(config) : reply(config, 202, { statusId: "status-1" });
      }
      if (config.method === "get" && url.startsWith("/builder/v1/users/packages/status/")) {
        return mos.status
          ? mos.status(config)
          : reply(config, 200, { titleId: mosTitleId, appId: mosAppId });
      }
      if (config.method === "get" && url.startsWith("/marketplace/v1/users/titles/")) {
        return mos.sharingInfo
          ? mos.sharingInfo(config)
          : reply(config, 200, { unifiedStoreLink: mosShareLink });
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

  function mosPosts(): InternalAxiosRequestConfig[] {
    return posts().filter((request) => request.url === "/builder/v1/users/packages");
  }

  function mosRequests(): InternalAxiosRequestConfig[] {
    return requests.filter((request) =>
      ["/config/", "/builder/", "/marketplace/"].some((prefix) => request.url?.startsWith(prefix))
    );
  }

  function outputsOf(result: Result<Map<string, string>, FxError>): Record<string, string> {
    return result.isOk() ? Object.fromEntries(result.value) : {};
  }

  it("SPFX-PREP-07: submits the exact prepared app package once after SharePoint verification and replays the confirmed result", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    route = mosAndCatalogRoute(catalog);
    const { context, approvals } = createContext();

    const published = await publish(prepared, context);
    const replayed = await publish(prepared, context);

    expect(published.result.isErr() ? published.result.error : undefined).toBeUndefined();
    expect(outputsOf(published.result)).toEqual({
      M365_TITLE_ID: mosTitleId,
      M365_APP_ID: mosAppId,
    });
    expect(outputsOf(replayed.result)).toEqual(outputsOf(published.result));
    expect(approvals()).toBe(1);
    expect(mosPosts()).toHaveLength(1);
    const [upload] = mosPosts();
    expect(upload.params).toEqual({ scope: "Tenant" });
    expect(upload.headers.get("Authorization")).toBe("Bearer fakeToken");
    expect(
      Buffer.compare(
        multipartPart(upload.data, "package") ?? Buffer.alloc(0),
        await fs.readFile(path.join(projectPath ?? "", prepared.appPackagePath))
      )
    ).toBe(0);
    expect(JSON.parse(multipartPart(upload.data, "info")?.toString("utf8") ?? "null")).toEqual({
      builderName: "TeamsToolKit",
    });
    const journal = await journalOf(prepared);
    expect(journal.state).toBe("submitted");
    expect(journal.applicationArtifact.submissionState).toBe("confirmed");
    expect(journal.applicationArtifact.submission).toEqual({
      scope: "Tenant",
      titleId: mosTitleId,
      appId: mosAppId,
    });
  });

  it("SPFX-PREP-22: a confirmed submission replays without a second MOS3 request after the developer app package is rebuilt with identical content", async () => {
    const { root, prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    route = mosAndCatalogRoute(catalog);
    const { context, approvals } = createContext();
    const published = await publish(prepared, context);
    const applicationPath = path.join(root, "appPackage", "application.zip");
    const rebuilt = createAppPackage("2.3.4", {
      entryTime: new Date(2021, 5, 15, 12, 30, 40),
      reverseOrder: true,
    });
    expect(rebuilt.equals(await fs.readFile(applicationPath))).toBe(false);
    await fs.writeFile(applicationPath, rebuilt);

    const replayed = await publish(prepared, context);

    expect(published.result.isErr() ? published.result.error : undefined).toBeUndefined();
    expect(replayed.result.isErr() ? replayed.result.error : undefined).toBeUndefined();
    expect(outputsOf(replayed.result)).toEqual(outputsOf(published.result));
    expect(approvals()).toBe(1);
    expect(mosPosts()).toHaveLength(1);
    expect((await journalOf(prepared)).state).toBe("submitted");
  });

  it.each([
    {
      outcome: "a definitive rejection",
      fail: (config: InternalAxiosRequestConfig) =>
        reply(config, 400, { error: { code: "InvalidPackage" } }),
      code: "MosSubmissionRejected",
      state: "rejected",
      httpStatus: 400,
    },
    {
      outcome: "an uncertain network failure",
      fail: (config: InternalAxiosRequestConfig) =>
        Promise.reject(new AxiosError("socket hang up", AxiosError.ERR_NETWORK, config)),
      code: "MosSubmissionOutcomeUnknown",
      state: "unknown",
      httpStatus: undefined,
    },
  ])(
    "SPFX-PREP-08: records $outcome of the app submission without retrying and keeps the verified SharePoint stage",
    async ({ fail, code, state, httpStatus }) => {
      const { prepared } = await prepareProject();
      const catalog: FakeCatalog = { deployed: false };
      await verifySharePoint(prepared, catalog);
      route = mosAndCatalogRoute(catalog, { upload: async (config) => fail(config) });
      const { context } = createContext();

      const published = await publish(prepared, context);
      const again = await publish(prepared, context);

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
      expect(mosPosts()).toHaveLength(1);
      const journal = await journalOf(prepared);
      expect(journal.state).toBe(state);
      expect(journal.sharePointArtifact.deployState).toBe("confirmed");
    }
  );

  it("SPFX-PREP-07: refuses app submission when SharePoint no longer reports the deployment, before approval or write", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    catalog.deployed = false;
    route = mosAndCatalogRoute(catalog);
    const { context, approvals } = createContext();

    const published = await publish(prepared, context);

    expect(published.result.isErr() ? published.result.error.name : undefined).toBe(
      "SharePointDeploymentNoLongerVerified"
    );
    expect(approvals()).toBe(0);
    expect(mosRequests()).toHaveLength(0);
    expect((await journalOf(prepared)).state).toBe("sharePointVerified");
  });

  it("SPFX-PREP-19: cancellation after submission approval restores the verified stage without a MOS3 request", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    const controller = new AbortController();
    route = mosAndCatalogRoute(catalog);
    const { context } = createContext({}, controller.signal, () => controller.abort());

    const published = await publish(prepared, context);

    expect(published.result.isErr() ? published.result.error : undefined).toBeInstanceOf(
      UserCancelError
    );
    expect(mosRequests()).toHaveLength(0);
    const journal = await journalOf(prepared);
    expect(journal.state).toBe("sharePointVerified");
    expect(journal.applicationArtifact.submissionState).toBe("notStarted");
  });

  it("SPFX-PREP-07: refuses submission when YAML inputs no longer match the prepared journal", async () => {
    const { prepared } = await prepareProject();
    route = catalogRoute({ deployed: false });
    const { context } = createContext();
    const execute = (args: ReturnType<typeof publishArgs>) =>
      new CopilotAgentPublishDriver().execute(args, context, publishOutputNames);

    const mismatched = await execute({
      ...publishArgs(prepared),
      appPackageSha256: "0".repeat(64),
    });
    const incomplete = await execute({ ...publishArgs(prepared), pluginSha256: "not-a-hash" });
    const unverified = await execute(publishArgs(prepared));

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
  function noSubmission(): FakeMos {
    return {
      upload: async () => {
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
    expect(mosPosts()).toHaveLength(0);
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
      route = mosAndCatalogRoute(catalog, noSubmission());
      const { context } = createContext();

      const published = await publish(prepared, context);

      await expectNoSubmission(prepared, published.result, "ArtifactHashMismatch");
      expect(requests).toHaveLength(0);
    }
  );

  it("SPFX-PREP-07: refuses submission when the tenant app catalog moved after verification", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    const serve = mosAndCatalogRoute(catalog, noSubmission());
    route = async (config) =>
      config.url === "/_api/SP_TenantSettings_Current"
        ? reply(config, 200, { CorporateCatalogUrl: `${tenantUrl}/sites/othercatalog` })
        : serve(config);
    const { context } = createContext();

    const published = await publish(prepared, context);

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
    route = mosAndCatalogRoute(catalog, noSubmission());
    const { context } = createContext();

    const published = await publish(prepared, context);

    await expectNoSubmission(prepared, published.result, code);
  });

  it("SPFX-PREP-07: refuses submission without an interactive approval surface", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    route = mosAndCatalogRoute(catalog, noSubmission());
    const { context } = createContext();

    const published = await publish(prepared, { ...context, ui: undefined });

    await expectNoSubmission(prepared, published.result, "ApprovalUnavailable");
  });

  it("SPFX-PREP-07: a declined submission approval sends no MOS3 request", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    route = mosAndCatalogRoute(catalog, noSubmission());
    const { context } = createContext();
    vi.mocked(context.ui!.showMessage).mockResolvedValue(ok("Cancel"));

    const published = await publish(prepared, context);

    expect(published.result.isErr() ? published.result.error : undefined).toBeInstanceOf(
      UserCancelError
    );
    expect(mosRequests()).toHaveLength(0);
    expect((await journalOf(prepared)).state).toBe("sharePointVerified");
  });

  it("SPFX-PREP-19: an already cancelled submission makes no remote call", async () => {
    const { prepared } = await prepareProject();
    const catalog: FakeCatalog = { deployed: false };
    await verifySharePoint(prepared, catalog);
    route = mosAndCatalogRoute(catalog, noSubmission());
    const controller = new AbortController();
    controller.abort();
    const { context } = createContext({}, controller.signal);

    const published = await publish(prepared, context);

    expect(published.result.isErr() ? published.result.error : undefined).toBeInstanceOf(
      UserCancelError
    );
    expect(requests).toHaveLength(0);
  });

  it.each([
    { name: "SharePoint", code: "SharePointAuthenticationUnavailable", approvals: 0 },
    { name: "MOS3", code: "MosTokenUnavailable", approvals: 1 },
  ])(
    "SPFX-PREP-07: refuses submission when the $name token is unavailable",
    async ({ name, code, approvals: expectedApprovals }) => {
      const { prepared } = await prepareProject();
      const catalog: FakeCatalog = { deployed: false };
      await verifySharePoint(prepared, catalog);
      route = mosAndCatalogRoute(catalog, noSubmission());
      const { context, approvals } = createContext();
      if (name === "SharePoint") {
        vi.mocked(Tools.getSPFxToken).mockResolvedValue(undefined);
      } else {
        vi.spyOn(context.m365TokenProvider, "getAccessToken").mockResolvedValue(
          err(new UserError("test", code, "no token"))
        );
      }

      const published = await publish(prepared, context);

      await expectNoSubmission(prepared, published.result, code);
      expect(approvals()).toBe(expectedApprovals);
      if (name === "MOS3") {
        const error = published.result.isErr() ? published.result.error : undefined;
        expect(error).toBeInstanceOf(UserError);
        expect(error && "data" in error ? error.data : undefined).toMatchObject({
          stage: "applicationPreflight",
          lastSharePointVerifiedState: "sharePointVerified",
          publicationId: prepared.publicationId,
        });
      }
    }
  );

  it("SPFX-PREP-18: refuses prepared submission in a sovereign cloud before any remote call", async () => {
    const { prepared } = await prepareProject();
    route = catalogRoute({ deployed: false });
    const { context } = createContext();
    vi.stubEnv("TEAMSFX_SOVEREIGN_CLOUD_ENVIRONMENT", "DoD");

    const published = await publish(prepared, context);
    vi.unstubAllEnvs();

    expect(published.result.isErr() ? published.result.error.name : undefined).toBe(
      "UnsupportedCloud"
    );
    expect(requests).toHaveLength(0);
  });

  describe("prepared MOS3 submission contracts", () => {
    const statusRequests = () =>
      requests.filter((request) => request.url?.startsWith("/builder/v1/users/packages/status/"));

    function dataOf(result: Result<unknown, FxError>): unknown {
      const error = result.isErr() ? result.error : undefined;
      return error && "data" in error ? error.data : undefined;
    }

    interface SubmitOptions {
      claims?: TokenClaimsOverride;
      scope?: string;
      outputs?: Map<string, string>;
      signal?: AbortSignal;
      onApproval?: () => void;
      rai?: boolean;
    }

    /** Verifies the SharePoint stage, then runs one prepared submission against the given MOS3 fake. */
    async function submit(mos: FakeMos, options: SubmitOptions = {}) {
      const { prepared } = await prepareProject();
      const catalog: FakeCatalog = { deployed: false };
      await verifySharePoint(prepared, catalog);
      route = mosAndCatalogRoute(catalog, mos);
      const { context, approvals } = createContext(
        options.claims,
        options.signal,
        options.onApproval
      );
      if (options.rai) {
        vi.stubEnv("TEAMSFX_RAI_VALIDATION_ENABLED", "true");
      }
      const published = await publish(prepared, context, options.scope, options.outputs);
      return { prepared, catalog, context, approvals, published };
    }

    it.each([
      {
        name: "SharePoint token audience is not the prepared tenant",
        claims: { sharePointAudience: "https://other.example.test" },
        code: "SharePointTokenIdentityMismatch",
      },
      {
        name: "SharePoint token claims are unreadable",
        claims: { sharePointClaimsUnavailable: true },
        code: "SharePointTokenIdentityUnavailable",
      },
      {
        name: "MOS3 token belongs to another tenant",
        claims: { mosTenantId: otherTenantId },
        code: "MosTokenIdentityMismatch",
      },
      {
        name: "MOS3 token claims are unreadable",
        claims: { mosClaimsUnavailable: true },
        code: "MosTokenIdentityUnavailable",
      },
    ])(
      "SPFX-PREP-15: refuses submission before approval when the $name",
      async ({ claims, code }) => {
        const { prepared, approvals, published } = await submit(noSubmission(), { claims });

        await expectNoSubmission(prepared, published.result, code);
        expect(approvals()).toBe(0);
        expect(requests.map((request) => request.url)).toEqual(["/_api/SP_TenantSettings_Current"]);
        expect(dataOf(published.result)).toMatchObject({ stage: "sharePointVerification" });
      }
    );

    it.each([
      { name: "a non-error rejection", reason: "token service said: secret-detail" },
      { name: "an unclassified error", reason: new Error("token service said: secret-detail") },
    ])(
      "SPFX-PREP-11: $name after approval is reported as PreparedPublicationFailed with allowlisted diagnostics only",
      async ({ reason }) => {
        const { prepared } = await prepareProject();
        const catalog: FakeCatalog = { deployed: false };
        await verifySharePoint(prepared, catalog);
        route = mosAndCatalogRoute(catalog, noSubmission());
        const { context, approvals } = createContext();
        vi.spyOn(context.m365TokenProvider, "getAccessToken").mockRejectedValue(reason);

        const published = await publish(prepared, context);

        await expectNoSubmission(prepared, published.result, "PreparedPublicationFailed");
        const error = published.result.isErr() ? published.result.error : undefined;
        expect(error).toBeInstanceOf(UserError);
        expect(dataOf(published.result)).toMatchObject({
          stage: "applicationPreflight",
          lastDurableState: "sharePointVerified",
          appPackageSha256: prepared.appPackageSha256,
        });
        expect(
          JSON.stringify({ message: error?.message, data: dataOf(published.result) })
        ).not.toContain("secret-detail");
        expect(approvals()).toBe(1);
        expect(mosRequests()).toHaveLength(0);
      }
    );

    it.each([
      {
        name: "the upload fails with a server error",
        mos: { upload: async (config: InternalAxiosRequestConfig) => reply(config, 500) },
        polls: 0,
        httpStatus: 500,
      },
      {
        name: "the upload times out upstream",
        mos: { upload: async (config: InternalAxiosRequestConfig) => reply(config, 408) },
        polls: 0,
        httpStatus: 408,
      },
      {
        name: "the upload is throttled",
        mos: { upload: async (config: InternalAxiosRequestConfig) => reply(config, 429) },
        polls: 0,
        httpStatus: 429,
      },
      {
        name: "the upload is accepted without a status identifier",
        mos: { upload: async (config: InternalAxiosRequestConfig) => reply(config, 202, {}) },
        polls: 0,
      },
      {
        name: "the upload returns an unsafe status identifier",
        mos: {
          upload: async (config: InternalAxiosRequestConfig) =>
            reply(config, 202, { statusId: "../escape" }),
        },
        polls: 0,
      },
      {
        name: "the status poll fails with a server error",
        mos: { status: async (config: InternalAxiosRequestConfig) => reply(config, 500) },
        polls: 1,
        httpStatus: 500,
      },
      {
        name: "the status reports malformed identifiers",
        mos: {
          status: async (config: InternalAxiosRequestConfig) =>
            reply(config, 200, { titleId: "not valid!", appId: mosAppId }),
        },
        polls: 1,
      },
      {
        name: "asynchronous validation reports an unrecognized status",
        mos: {
          status: async (config: InternalAxiosRequestConfig) =>
            reply(config, 200, { validationInfo: { status: "unrecognized" } }),
        },
        rai: true,
        polls: 1,
      },
    ])(
      "SPFX-PREP-16: submits once and records an unknown outcome when $name",
      async ({ mos, rai, httpStatus, polls }) => {
        const { prepared, context, published } = await submit(mos, { rai });

        expect(published.result.isErr() ? published.result.error.name : undefined).toBe(
          "MosSubmissionOutcomeUnknown"
        );
        expect(mosPosts()).toHaveLength(1);
        expect(statusRequests()).toHaveLength(polls);
        const data = dataOf(published.result);
        expect(data).toMatchObject({ stage: "applicationSubmission", lastDurableState: "unknown" });
        if (httpStatus !== undefined) {
          expect(data).toMatchObject({ httpStatus });
        }
        expect(JSON.stringify(data)).not.toContain("escape");
        const journal = await journalOf(prepared);
        expect(journal.state).toBe("unknown");
        expect(journal.applicationArtifact.submissionState).toBe("unknown");
        const again = await publish(prepared, context);
        expect(again.result.isErr() ? again.result.error.name : undefined).toBe(
          "PreparedPublicationNotVerified"
        );
        expect(mosPosts()).toHaveLength(1);
      }
    );

    it("SPFX-PREP-16: records an unknown outcome once the bounded status poll is exhausted", async () => {
      const { prepared, published } = await submit({
        status: async (config) => reply(config, 202, {}),
      });

      expect(published.result.isErr() ? published.result.error.name : undefined).toBe(
        "MosSubmissionOutcomeUnknown"
      );
      expect(mosPosts()).toHaveLength(1);
      expect(statusRequests()).toHaveLength(90);
      const waits = vi.mocked(commonUtils.waitSeconds).mock.calls;
      expect(waits).toHaveLength(89);
      expect(waits.every(([seconds]) => seconds === 7)).toBe(true);
      expect((await journalOf(prepared)).state).toBe("unknown");
    });

    it("SPFX-PREP-16: waits for asynchronous validation and confirms the submission once it succeeds", async () => {
      let polls = 0;
      const { prepared, published } = await submit(
        {
          status: async (config) => {
            polls += 1;
            return polls < 3
              ? reply(config, 200, { validationInfo: { status: "pending" } })
              : reply(config, 200, {
                  validationInfo: { status: "successful" },
                  titleId: mosTitleId,
                  appId: mosAppId,
                });
          },
        },
        { rai: true }
      );

      expect(published.result.isErr() ? published.result.error : undefined).toBeUndefined();
      expect(outputsOf(published.result)).toEqual({
        M365_TITLE_ID: mosTitleId,
        M365_APP_ID: mosAppId,
      });
      expect(mosPosts()).toHaveLength(1);
      expect(mosPosts()[0].params).toEqual({ scope: "Tenant", isRAIValidationAsync: true });
      expect(statusRequests()).toHaveLength(3);
      expect(
        statusRequests().every((request) => request.params?.isRAIValidationAsync === true)
      ).toBe(true);
      expect(vi.mocked(commonUtils.waitSeconds)).toHaveBeenCalledTimes(2);
      expect((await journalOf(prepared)).state).toBe("submitted");
    });

    it("SPFX-PREP-16: records a definitive rejection without leaking failure details when asynchronous validation fails", async () => {
      const { prepared, context, published } = await submit(
        {
          status: async (config) =>
            reply(config, 200, {
              validationInfo: {
                status: "failed",
                failureReasons: [{ errorCode: "RaiBlocked", errorDetail: "confidential detail" }],
              },
            }),
        },
        { rai: true }
      );

      expect(published.result.isErr() ? published.result.error.name : undefined).toBe(
        "MosSubmissionRejected"
      );
      expect(JSON.stringify(dataOf(published.result))).not.toContain("confidential");
      expect(mosPosts()).toHaveLength(1);
      const journal = await journalOf(prepared);
      expect(journal.state).toBe("rejected");
      expect(journal.applicationArtifact.submissionState).toBe("rejected");
      const again = await publish(prepared, context);
      expect(again.result.isErr() ? again.result.error.name : undefined).toBe(
        "PreparedPublicationNotVerified"
      );
      expect(mosPosts()).toHaveLength(1);
    });

    it.each([
      { name: "is not https", body: { titlesServiceUrl: "http://titles.example.test" } },
      {
        name: "embeds credentials",
        body: { titlesServiceUrl: "https://user:secret@titles.example.test" },
      },
      {
        name: "carries a query string",
        body: { titlesServiceUrl: "https://titles.example.test/?next=1" },
      },
      {
        name: "carries a fragment",
        body: { titlesServiceUrl: "https://titles.example.test/#fragment" },
      },
      { name: "is missing", body: {} },
    ])(
      "SPFX-PREP-16: refuses submission before any upload when the discovered title service URL $name",
      async ({ body }) => {
        const { prepared, approvals, published } = await submit({
          ...noSubmission(),
          environment: async (config) => reply(config, 200, body),
        });

        await expectNoSubmission(prepared, published.result, "MosEndpointUnavailable");
        expect(approvals()).toBe(1);
        expect(mosRequests().map((request) => request.url)).toEqual(["/config/v1/environment"]);
        expect(dataOf(published.result)).toMatchObject({ stage: "applicationPreflight" });
        expect(JSON.stringify(dataOf(published.result))).not.toContain("secret");
      }
    );

    it("SPFX-PREP-16: refuses submission when MOS3 environment discovery returns an HTTP error without retrying it", async () => {
      const { prepared, approvals, published } = await submit({
        ...noSubmission(),
        environment: async (config) => reply(config, 503),
      });

      await expectNoSubmission(prepared, published.result, "MosEndpointUnavailable");
      expect(approvals()).toBe(1);
      expect(mosRequests()).toHaveLength(1);
      expect(vi.mocked(commonUtils.waitSeconds)).not.toHaveBeenCalled();
    });

    it("SPFX-PREP-16: retries a transport failure of MOS3 environment discovery within its bound before refusing submission", async () => {
      const { prepared, published } = await submit({
        ...noSubmission(),
        environment: async (config) =>
          Promise.reject(new AxiosError("socket hang up", AxiosError.ERR_NETWORK, config)),
      });

      await expectNoSubmission(prepared, published.result, "MosEndpointUnavailable");
      expect(mosRequests()).toHaveLength(3);
      expect(vi.mocked(commonUtils.waitSeconds).mock.calls).toEqual([[1], [1]]);
    });

    const shareLinkLookupUrl = `/marketplace/v1/users/titles/${encodeURIComponent(
      mosTitleId
    )}/sharingInfo`;
    const sharedSubmissionOutputs = {
      M365_TITLE_ID: mosTitleId,
      M365_APP_ID: mosAppId,
      M365_SHARE_LINK: mosShareLink,
    };

    it("SPFX-PREP-17: refuses to replay a confirmed submission under a different scope", async () => {
      const { prepared, context, approvals, published } = await submit({});
      expect(published.result.isOk()).toBe(true);
      const recorded = mosRequests().length;

      const replay = await publish(prepared, context, "personal");

      expect(replay.result.isErr() ? replay.result.error.name : undefined).toBe(
        "PreparedSubmissionScopeMismatch"
      );
      expect(mosRequests()).toHaveLength(recorded);
      expect(mosPosts()).toHaveLength(1);
      expect(approvals()).toBe(1);
      const journal = await journalOf(prepared);
      expect(journal.state).toBe("submitted");
      expect(journal.applicationArtifact.submission).toMatchObject({ scope: "Tenant" });
    });

    it("SPFX-PREP-07: an omitted scope submits to the personal scope and the replay stays bound to it", async () => {
      const { prepared } = await prepareProject();
      const catalog: FakeCatalog = { deployed: false };
      await verifySharePoint(prepared, catalog);
      route = mosAndCatalogRoute(catalog, {});
      const { context, approvals } = createContext();
      const args = publishArgs(prepared);
      Reflect.deleteProperty(args, "scope");
      const driver = new CopilotAgentPublishDriver();

      const published = await driver.execute(args, context, publishOutputNames);
      const replay = await driver.execute(args, context, publishOutputNames);
      const differentScope = await publish(prepared, context, "tenant");

      expect(outputsOf(published.result)).toEqual({
        M365_TITLE_ID: mosTitleId,
        M365_APP_ID: mosAppId,
      });
      expect(outputsOf(replay.result)).toEqual(outputsOf(published.result));
      expect(differentScope.result.isErr() ? differentScope.result.error.name : undefined).toBe(
        "PreparedSubmissionScopeMismatch"
      );
      expect(mosPosts()).toHaveLength(1);
      expect(mosPosts()[0].params).toEqual({ scope: "Personal" });
      expect(approvals()).toBe(1);
      expect((await journalOf(prepared)).applicationArtifact.submission).toEqual({
        scope: "Personal",
        titleId: mosTitleId,
        appId: mosAppId,
      });
    });

    it("SPFX-PREP-17: a shared replay returns the recorded identifiers and re-reads only the share link", async () => {
      const { prepared, context, approvals, published } = await submit({}, { scope: "shared" });
      expect(outputsOf(published.result)).toEqual(sharedSubmissionOutputs);
      expect(mosPosts()[0].params).toEqual({ scope: "Shared" });
      const recorded = mosRequests().length;

      const replay = await publish(prepared, context, "shared");

      expect(outputsOf(replay.result)).toEqual(sharedSubmissionOutputs);
      expect(mosPosts()).toHaveLength(1);
      expect(approvals()).toBe(1);
      expect(
        mosRequests()
          .slice(recorded)
          .map((request) => request.url)
      ).toEqual(["/config/v1/environment", shareLinkLookupUrl]);
    });

    it("SPFX-PREP-17: a shared replay verifies the signed-in MOS3 tenant before reading the share link", async () => {
      const first = await submit({}, { scope: "shared" });
      expect(first.published.result.isOk()).toBe(true);
      const recorded = mosRequests().length;
      vi.restoreAllMocks();
      const { context } = createContext({ mosTenantId: otherTenantId });

      const replay = await publish(first.prepared, context, "shared");

      expect(replay.result.isErr() ? replay.result.error.name : undefined).toBe(
        "MosTokenIdentityMismatch"
      );
      expect(mosRequests()).toHaveLength(recorded);
    });

    it.each([
      {
        name: "a server error",
        serve: async (config: InternalAxiosRequestConfig) => reply(config, 500),
      },
      {
        name: "a link that is not https",
        serve: async (config: InternalAxiosRequestConfig) =>
          reply(config, 200, { unifiedStoreLink: "http://share.example.test/apps/fixture" }),
      },
      {
        name: "a link with whitespace",
        serve: async (config: InternalAxiosRequestConfig) =>
          reply(config, 200, { unifiedStoreLink: "https://share.example.test/apps/fix ture" }),
      },
      {
        name: "a response without a link",
        serve: async (config: InternalAxiosRequestConfig) => reply(config, 200, {}),
      },
    ])(
      "SPFX-PREP-17: $name while reading the share link keeps the submission confirmed and a rerun retries only the lookup",
      async ({ serve }) => {
        const first = await submit({ sharingInfo: serve }, { scope: "shared" });

        expect(first.published.result.isErr() ? first.published.result.error.name : undefined).toBe(
          "MosShareLinkUnavailable"
        );
        expect(mosPosts()).toHaveLength(1);
        const journal = await journalOf(first.prepared);
        expect(journal.state).toBe("submitted");
        expect(journal.applicationArtifact.submissionState).toBe("confirmed");
        expect(journal.applicationArtifact.submission).toMatchObject({
          scope: "Shared",
          titleId: mosTitleId,
          appId: mosAppId,
        });

        route = mosAndCatalogRoute(first.catalog, {});
        const rerun = await publish(first.prepared, first.context, "shared");

        expect(outputsOf(rerun.result)).toEqual(sharedSubmissionOutputs);
        expect(mosPosts()).toHaveLength(1);
        expect(first.approvals()).toBe(1);
      }
    );

    it("SPFX-PREP-17: a shared submission without a share link output never reads the link", async () => {
      const outputs = new Map([
        ["titleId", "M365_TITLE_ID"],
        ["appId", "M365_APP_ID"],
      ]);

      const { prepared, context, published } = await submit({}, { scope: "shared", outputs });
      const replay = await publish(prepared, context, "shared", outputs);

      expect(outputsOf(published.result)).toEqual({
        M365_TITLE_ID: mosTitleId,
        M365_APP_ID: mosAppId,
      });
      expect(outputsOf(replay.result)).toEqual(outputsOf(published.result));
      expect(requests.some((request) => request.url?.startsWith("/marketplace/"))).toBe(false);
      expect(mosRequests()).toHaveLength(3);
    });

    it.each<{ name: string; override: Record<string, unknown> }>([
      { name: "a missing hash", override: { sharePointPackageSha256: undefined } },
      { name: "a non-hexadecimal hash", override: { appPackageSha256: "z".repeat(64) } },
      { name: "a missing publication id", override: { publicationId: undefined } },
      {
        name: "a missing prepared publication path",
        override: { preparedPublicationPath: undefined },
      },
      { name: "an empty application package path", override: { appPackagePath: "" } },
      { name: "an unknown scope", override: { scope: "everyone" } },
      { name: "a non-string scope", override: { scope: 7 } },
    ])(
      "SPFX-PREP-18: refuses $name before any approval or remote request",
      async ({ override }) => {
        const { prepared } = await prepareProject();
        route = catalogRoute({ deployed: false });
        const { context, approvals } = createContext();

        const { result } = await new CopilotAgentPublishDriver().execute(
          { ...publishArgs(prepared), ...override },
          context,
          publishOutputNames
        );

        expect(result.isErr() ? result.error.name : undefined).toBe(
          "InvalidPreparedPublishArguments"
        );
        expect(approvals()).toBe(0);
        expect(requests).toHaveLength(0);
      }
    );

    it("SPFX-PREP-18: partial prepared inputs never fall back to the unprepared upload", async () => {
      const { prepared } = await prepareProject();
      route = catalogRoute({ deployed: false });
      const { context, approvals } = createContext();
      const { appPackagePath, scope, preparedPublicationPath } = publishArgs(prepared);

      const { result } = await new CopilotAgentPublishDriver().execute(
        { appPackagePath, scope, preparedPublicationPath },
        context,
        publishOutputNames
      );

      expect(result.isErr() ? result.error.name : undefined).toBe(
        "InvalidPreparedPublishArguments"
      );
      expect(approvals()).toBe(0);
      expect(requests).toHaveLength(0);
    });

    it("SPFX-PREP-18: a missing required output mapping is refused before any remote request", async () => {
      const { prepared } = await prepareProject();
      route = catalogRoute({ deployed: false });
      const { context, approvals } = createContext();
      const driver = new CopilotAgentPublishDriver();

      const withoutAppId = await driver.execute(
        publishArgs(prepared),
        context,
        new Map([["titleId", "M365_TITLE_ID"]])
      );
      const withoutOutputs = await driver.run(publishArgs(prepared), context);

      expect(withoutAppId.result.isErr() ? withoutAppId.result.error.name : undefined).toBe(
        "InvalidPreparedPublishArguments"
      );
      expect(withoutOutputs.isErr() ? withoutOutputs.error.name : undefined).toBe(
        "InvalidPreparedPublishArguments"
      );
      expect(approvals()).toBe(0);
      expect(requests).toHaveLength(0);
    });

    it("SPFX-PREP-18: refuses an application package other than the staged artifact before any approval", async () => {
      const { prepared } = await prepareProject();
      route = catalogRoute({ deployed: false });
      const { context, approvals } = createContext();

      const { result } = await new CopilotAgentPublishDriver().execute(
        { ...publishArgs(prepared), appPackagePath: "sharepoint/solution.sppkg" },
        context,
        publishOutputNames
      );

      expect(result.isErr() ? result.error.name : undefined).toBe(
        "ApplicationArtifactPathMismatch"
      );
      expect(approvals()).toBe(0);
      expect(requests).toHaveLength(0);
    });

    it("SPFX-PREP-18: refuses an application package path that does not exist with a typed error", async () => {
      const { prepared } = await prepareProject();
      route = catalogRoute({ deployed: false });
      const { context, approvals } = createContext();

      const { result } = await new CopilotAgentPublishDriver().execute(
        { ...publishArgs(prepared), appPackagePath: "build/appPackage/missing.zip" },
        context,
        publishOutputNames
      );

      expect(result.isErr() ? result.error.name : undefined).toBe(
        "ApplicationArtifactPathMismatch"
      );
      expect(approvals()).toBe(0);
      expect(requests).toHaveLength(0);
    });

    it("SPFX-PREP-18: refuses GCC High before any approval or remote request", async () => {
      const { prepared } = await prepareProject();
      route = catalogRoute({ deployed: false });
      const { context, approvals } = createContext();
      vi.stubEnv("TEAMSFX_SOVEREIGN_CLOUD_ENVIRONMENT", "GCC H");

      const { result } = await publish(prepared, context);

      expect(result.isErr() ? result.error.name : undefined).toBe("UnsupportedCloud");
      expect(approvals()).toBe(0);
      expect(requests).toHaveLength(0);
    });

    it("SPFX-PREP-08: simultaneous submissions of one prepared publication upload exactly once", async () => {
      const { prepared } = await prepareProject();
      const catalog: FakeCatalog = { deployed: false };
      await verifySharePoint(prepared, catalog);
      route = mosAndCatalogRoute(catalog, {});
      const { context, approvals } = createContext();

      const results = await Promise.all([publish(prepared, context), publish(prepared, context)]);

      const failures = results.flatMap(({ result }) => (result.isErr() ? [result.error.name] : []));
      expect(failures.length).toBeLessThan(2);
      expect(failures.every((name) => name === "PublicationIdentityLockUnavailable")).toBe(true);
      expect(mosPosts()).toHaveLength(1);
      expect(approvals()).toBe(1);
      expect((await journalOf(prepared)).state).toBe("submitted");
    });

    it("SPFX-PREP-19: an upload that never settles ends at the 30 second bound as an unknown outcome", async () => {
      const timeouts: { ms: number; controller: AbortController }[] = [];

      const { prepared, published } = await submit(
        {
          upload: async (config) => {
            const pending = hangUntilAborted(config);
            timeouts.find((entry) => entry.controller.signal === config.signal)?.controller.abort();
            return pending;
          },
        },
        {
          onApproval: () => {
            vi.spyOn(AbortSignal, "timeout").mockImplementation((ms) => {
              const controller = new AbortController();
              timeouts.push({ ms, controller });
              return controller.signal;
            });
          },
        }
      );

      expect(published.result.isErr() ? published.result.error.name : undefined).toBe(
        "MosSubmissionOutcomeUnknown"
      );
      expect(timeouts.length).toBeGreaterThan(0);
      expect(timeouts.every((entry) => entry.ms === 30_000)).toBe(true);
      expect(mosPosts()).toHaveLength(1);
      const journal = await journalOf(prepared);
      expect(journal.state).toBe("unknown");
      expect(journal.applicationArtifact.submissionState).toBe("unknown");
    });

    it("SPFX-PREP-19: cancelling while the upload is outstanding persists an unknown outcome and never polls", async () => {
      const controller = new AbortController();

      const { prepared, published } = await submit(
        {
          upload: async (config) => {
            controller.abort();
            return reply(config, 202, { statusId: "status-1" });
          },
        },
        { signal: controller.signal }
      );

      expect(published.result.isErr() ? published.result.error.name : undefined).toBe(
        "MosSubmissionOutcomeUnknown"
      );
      expect(dataOf(published.result)).toMatchObject({
        stage: "applicationSubmission",
        lastDurableState: "unknown",
      });
      expect(mosPosts()).toHaveLength(1);
      expect(statusRequests()).toHaveLength(0);
      const journal = await journalOf(prepared);
      expect(journal.state).toBe("unknown");
      expect(journal.applicationArtifact.submissionState).toBe("unknown");
    });

    it("SPFX-PREP-19: cancelling during the status poll persists an unknown outcome", async () => {
      const controller = new AbortController();

      const { prepared, published } = await submit(
        {
          status: async (config) => {
            controller.abort();
            return reply(config, 202, {});
          },
        },
        { signal: controller.signal }
      );

      expect(published.result.isErr() ? published.result.error.name : undefined).toBe(
        "MosSubmissionOutcomeUnknown"
      );
      expect(mosPosts()).toHaveLength(1);
      expect(statusRequests()).toHaveLength(1);
      const journal = await journalOf(prepared);
      expect(journal.state).toBe("unknown");
      expect(journal.applicationArtifact.submissionState).toBe("unknown");
    });

    it("SPFX-PREP-19: cancelling while the share link is read keeps the confirmed submission and a rerun completes it", async () => {
      const controller = new AbortController();

      const first = await submit(
        {
          sharingInfo: async (config) => {
            controller.abort();
            return reply(config, 200, { unifiedStoreLink: mosShareLink });
          },
        },
        { scope: "shared", signal: controller.signal }
      );

      expect(
        first.published.result.isErr() ? first.published.result.error : undefined
      ).toBeInstanceOf(UserCancelError);
      expect(mosPosts()).toHaveLength(1);
      const journal = await journalOf(first.prepared);
      expect(journal.state).toBe("submitted");
      expect(journal.applicationArtifact.submissionState).toBe("confirmed");

      vi.restoreAllMocks();
      route = mosAndCatalogRoute(first.catalog, {});
      const { context } = createContext();
      const rerun = await publish(first.prepared, context, "shared");

      expect(outputsOf(rerun.result)).toEqual(sharedSubmissionOutputs);
      expect(mosPosts()).toHaveLength(1);
    });
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
            ? (await publish(prepared, context)).result
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
          expect(result.error.data?.lastSharePointVerifiedState).toBeUndefined();
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
        const result = await publish(prepared, context);
        expect(result.result.isErr() ? result.result.error.name : undefined).toBe(
          "MissingSharePointVerificationData"
        );
        expect(approvals()).toBe(0);
        expect(requests).toHaveLength(0);
        expect((await journalOf(prepared)).state).toBe("sharePointVerified");
      }
    );

    it.each([
      { name: "not a URL", value: "not a url" },
      { name: "not https", value: "http://tenant.example.test/sites/appcatalog" },
      { name: "carrying a query", value: `${catalogUrl}?source=journal` },
      { name: "carrying a fragment", value: `${catalogUrl}#journal` },
      {
        name: "carrying credentials",
        value: "https://user:password@tenant.example.test/sites/appcatalog",
      },
    ])(
      "SPFX-PREP-07: submission refuses a recorded SharePoint catalog URL that is $name before approval or write",
      async ({ value }) => {
        const { prepared } = await prepareProject();
        const catalog: FakeCatalog = { deployed: false };
        await verifySharePoint(prepared, catalog);
        await externalUpdate(prepared, (journal) => {
          journal.sharePointArtifact.catalogUrl = value;
        });
        route = mosAndCatalogRoute(catalog, noSubmission());
        const { context, approvals } = createContext();

        const result = await publish(prepared, context);

        expect(result.result.isErr() ? result.result.error.name : undefined).toBe(
          "InvalidCatalogUrl"
        );
        expect(approvals()).toBe(0);
        expect(mosRequests()).toHaveLength(0);
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
        route = mosAndCatalogRoute(catalog, noSubmission());
        const { context } = createContext({}, controller.signal);
        const result =
          stage === "submission"
            ? (await publish(prepared, context)).result
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
        expect(mosRequests()).toHaveLength(0);
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
        route = mosAndCatalogRoute(catalog, {
          upload: async (config) => {
            await change();
            return boundary === "rejection"
              ? reply(config, 400, { error: { code: "InvalidPackage" } })
              : reply(config, 202, { statusId: "status-1" });
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
        const result = await publish(prepared, context);
        expect(result.result.isErr() ? result.result.error.name : undefined).toBe(
          "JournalUpdateFailed"
        );
        expect((await journalOf(prepared)).applicationArtifact.submissionState).toBe("unknown");
        expect(mosPosts()).toHaveLength(boundary === "approval" ? 0 : 1);
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
