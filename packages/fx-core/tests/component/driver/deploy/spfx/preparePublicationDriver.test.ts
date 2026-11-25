// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios, { AxiosAdapter } from "axios";
import * as fs from "fs-extra";
import os from "os";
import path from "path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { FxError, ok, err, Platform, Result, SystemError, UserError } from "@microsoft/teamsfx-api";

import { DriverContext } from "../../../../../src/component/driver/interface/commonArgs";
import { PreparePublicationDriver } from "../../../../../src/component/driver/deploy/spfx/preparePublicationDriver";
import { SPOClient } from "../../../../../src/component/driver/deploy/spfx/utility/spoClient";
import { verifyPreparedTokenIdentity } from "../../../../../src/component/driver/deploy/spfx/utility/tokenIdentity";
import { MockedAzureAccountProvider, MockedM365Provider } from "../../../../core/utils";
import { MockedLogProvider } from "../../../../plugins/solution/util";
import { createAppPackage, createSppkg, tenantId, tenantUrl } from "./preparedPublicationFixtures";

describe("SPFX-PREP-05/09: SharePoint publication preparation driver", () => {
  const directories: string[] = [];
  const originalAdapter = axios.defaults.adapter;

  afterEach(async () => {
    axios.defaults.adapter = originalAdapter;
    vi.restoreAllMocks();
    await Promise.all(directories.splice(0).map((directory) => fs.remove(directory)));
  });

  async function createProject(): Promise<string> {
    const projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "publication-driver-"));
    directories.push(projectPath);
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), createSppkg());
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), createAppPackage());
    return projectPath;
  }

  interface ContextOptions {
    claimsResult?: Result<Record<string, unknown>, FxError>;
    graphData?: Record<string, unknown>;
    signal?: AbortSignal;
  }

  function createContext(
    projectPath: string,
    claims: Record<string, unknown>,
    options: ContextOptions = {}
  ): DriverContext {
    const provider = new MockedM365Provider();
    vi.spyOn(provider, "getJsonObject").mockResolvedValue(options.claimsResult ?? ok(claims));
    const graphAdapter: AxiosAdapter = async (config) => ({
      data: options.graphData ?? { webUrl: tenantUrl },
      status: 200,
      statusText: "OK",
      headers: {},
      config,
      request: {},
    });
    axios.defaults.adapter = graphAdapter;
    return {
      azureAccountProvider: new MockedAzureAccountProvider(),
      m365TokenProvider: provider,
      ui: undefined,
      progressBar: undefined,
      logProvider: new MockedLogProvider(),
      projectPath,
      platform: Platform.VSCode,
      signal: options.signal,
    };
  }

  const validArgs = {
    packagePath: "sharepoint/solution.sppkg",
    appPackagePath: "appPackage/application.zip",
  };
  const validClaims = { tid: tenantId, aud: "https://graph.microsoft.com" };

  function validOutputs(): Map<string, string> {
    return new Map([
      ["sharePointPackagePath", "STAGED_SPPKG"],
      ["appPackagePath", "STAGED_APP_PACKAGE"],
    ]);
  }

  interface DriverCall {
    args?: unknown;
    context?: DriverContext;
    outputs?: Map<string, string>;
  }

  async function expectDriverError(
    projectPath: string,
    code: string,
    errorType: typeof UserError | typeof SystemError,
    call: DriverCall = {}
  ): Promise<void> {
    const getCatalogSpy = vi.spyOn(SPOClient, "getAppCatalogSite");

    const result = await new PreparePublicationDriver().execute(
      "args" in call ? call.args : validArgs,
      call.context ?? createContext(projectPath, validClaims),
      "outputs" in call ? call.outputs : validOutputs()
    );

    expect(result.result.isErr()).toBe(true);
    if (result.result.isOk()) return;
    expect(result.result.error).toBeInstanceOf(errorType);
    expect(result.result.error.message).toBe(code);
    expect(result.summaries).toEqual([]);
    expect(await fs.pathExists(path.join(projectPath, ".fx"))).toBe(false);
    expect(getCatalogSpy).not.toHaveBeenCalled();
  }

  it("writes staged output variables after verified tenant discovery without app-catalog calls", async () => {
    const projectPath = await createProject();
    const getCatalogSpy = vi.spyOn(SPOClient, "getAppCatalogSite");
    const createCatalogSpy = vi.spyOn(SPOClient, "createAppCatalog");
    const outputs = new Map([
      ["sharePointPackagePath", "STAGED_SPPKG"],
      ["appPackagePath", "STAGED_APP_PACKAGE"],
    ]);

    const result = await new PreparePublicationDriver().execute(
      {
        packagePath: "sharepoint/solution.sppkg",
        appPackagePath: "appPackage/application.zip",
      },
      createContext(projectPath, {
        tid: tenantId,
        aud: "https://graph.microsoft.com",
      }),
      outputs
    );

    expect(result.result.isOk()).toBe(true);
    if (result.result.isErr()) return;
    expect(result.result.value.get("STAGED_SPPKG")).toMatch(
      /^\.fx\/sharepoint-publications\/[0-9a-f-]+\/sharepoint\.sppkg$/
    );
    expect(result.result.value.get("STAGED_APP_PACKAGE")).toMatch(
      /^\.fx\/sharepoint-publications\/[0-9a-f-]+\/application\.zip$/
    );
    expect(result.result.value.get("STAGED_SPPKG")).not.toBe(
      result.result.value.get("STAGED_APP_PACKAGE")
    );
    expect(getCatalogSpy).not.toHaveBeenCalled();
    expect(createCatalogSpy).not.toHaveBeenCalled();

    for (const variable of ["STAGED_SPPKG", "STAGED_APP_PACKAGE"]) {
      const stagedPath = result.result.value.get(variable);
      expect(stagedPath).toBeDefined();
      expect(await fs.pathExists(path.join(projectPath, stagedPath ?? ""))).toBe(true);
    }
  });

  it("rejects a mismatched token audience before staging or touching SharePoint", async () => {
    const projectPath = await createProject();
    const getCatalogSpy = vi.spyOn(SPOClient, "getAppCatalogSite");
    const createCatalogSpy = vi.spyOn(SPOClient, "createAppCatalog");

    const result = await new PreparePublicationDriver().execute(
      {
        packagePath: "sharepoint/solution.sppkg",
        appPackagePath: "appPackage/application.zip",
      },
      createContext(projectPath, {
        tid: tenantId,
        aud: "https://sharepoint.example.test",
      }),
      new Map([
        ["sharePointPackagePath", "STAGED_SPPKG"],
        ["appPackagePath", "STAGED_APP_PACKAGE"],
      ])
    );

    expect(result.result.isErr()).toBe(true);
    expect(await fs.pathExists(path.join(projectPath, ".fx"))).toBe(false);
    expect(getCatalogSpy).not.toHaveBeenCalled();
    expect(createCatalogSpy).not.toHaveBeenCalled();
  });

  it.each<[string, unknown]>([
    ["a non-object input", "sharepoint/solution.sppkg"],
    ["an empty package path", { ...validArgs, packagePath: "" }],
    ["a missing application package path", { packagePath: validArgs.packagePath }],
    ["an unexpected input", { ...validArgs, createAppCatalogIfNotExist: true }],
  ])("SPFX-PREP-04: rejects %s before discovery", async (_name, args) => {
    const projectPath = await createProject();

    await expectDriverError(projectPath, "InvalidActionInput", UserError, { args });
  });

  it.each<[string, Map<string, string> | undefined]>([
    ["missing output mappings", undefined],
    ["an incomplete output mapping", new Map([["sharePointPackagePath", "STAGED_SPPKG"]])],
    [
      "an unknown output key",
      new Map([
        ["sharePointPackagePath", "STAGED_SPPKG"],
        ["otherPath", "STAGED_APP_PACKAGE"],
      ]),
    ],
    [
      "duplicate output variables",
      new Map([
        ["sharePointPackagePath", "STAGED_PATH"],
        ["appPackagePath", "STAGED_PATH"],
      ]),
    ],
    [
      "an invalid output variable name",
      new Map([
        ["sharePointPackagePath", "STAGED-SPPKG"],
        ["appPackagePath", "STAGED_APP_PACKAGE"],
      ]),
    ],
  ])("SPFX-PREP-05: rejects %s before discovery", async (_name, outputs) => {
    const projectPath = await createProject();

    await expectDriverError(projectPath, "InvalidActionOutput", UserError, { outputs });
  });

  it("SPFX-PREP-04: reports an unreadable SharePoint package as a user error", async () => {
    const projectPath = await createProject();

    await expectDriverError(projectPath, "InputReadFailed", UserError, {
      args: { ...validArgs, packagePath: "sharepoint/missing.sppkg" },
    });
  });

  it("SPFX-PREP-04: reports an unavailable project as a system error", async () => {
    const projectPath = await createProject();
    const missingProjectPath = path.join(projectPath, "missing");

    await expectDriverError(missingProjectPath, "ProjectPathUnavailable", SystemError, {
      context: createContext(missingProjectPath, validClaims),
    });
  });

  it("SPFX-PREP-05: stops before tenant discovery when cancelled", async () => {
    const projectPath = await createProject();
    const controller = new AbortController();
    controller.abort();
    const context = createContext(projectPath, validClaims, { signal: controller.signal });

    await expectDriverError(projectPath, "PublicationCancelled", UserError, { context });
    expect(context.m365TokenProvider.getJsonObject).not.toHaveBeenCalled();
  });

  it.each<[string, ContextOptions, string]>([
    [
      "a tenant URL without HTTPS",
      { graphData: { webUrl: "http://tenant.example.test" } },
      "InvalidPreparedTenantContext",
    ],
    ["a tenant lookup without a URL", { graphData: {} }, "TenantContextDiscoveryFailed"],
    [
      "unavailable token claims",
      { claimsResult: err(new UserError("Test", "TokenUnavailable", "Token unavailable")) },
      "GraphTokenIdentityUnavailable",
    ],
    [
      "a token tenant that is not a GUID",
      { claimsResult: ok({ tid: "tenant", aud: "https://graph.microsoft.com" }) },
      "GraphTokenIdentityMismatch",
    ],
  ])("SPFX-PREP-09: rejects %s before staging", async (_name, options, code) => {
    const projectPath = await createProject();

    await expectDriverError(projectPath, code, UserError, {
      context: createContext(projectPath, validClaims, options),
    });
  });

  it("SPFX-PREP-09: accepts token claims nested under a value object", async () => {
    const projectPath = await createProject();

    const result = await new PreparePublicationDriver().execute(
      validArgs,
      createContext(projectPath, {
        value: { tid: tenantId, aud: "00000003-0000-0000-c000-000000000000" },
      }),
      validOutputs()
    );

    expect(result.result.isOk()).toBe(true);
    expect(result.summaries).toHaveLength(1);
  });

  it("SPFX-PREP-09: rejects a verified token for a non-HTTPS tenant origin", async () => {
    const provider = new MockedM365Provider();
    vi.spyOn(provider, "getJsonObject").mockResolvedValue(ok(validClaims));

    await expect(
      verifyPreparedTokenIdentity(provider, "http://tenant.example.test")
    ).rejects.toMatchObject({ code: "InvalidPreparedTenantContext" });
  });
});
