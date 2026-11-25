// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import axios, { AxiosAdapter } from "axios";
import * as fs from "fs-extra";
import os from "os";
import path from "path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { ok, Platform } from "@microsoft/teamsfx-api";

import { DriverContext } from "../../../../../src/component/driver/interface/commonArgs";
import { PreparePublicationDriver } from "../../../../../src/component/driver/deploy/spfx/preparePublicationDriver";
import { SPOClient } from "../../../../../src/component/driver/deploy/spfx/utility/spoClient";
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

  function createContext(projectPath: string, claims: Record<string, unknown>): DriverContext {
    const provider = new MockedM365Provider();
    vi.spyOn(provider, "getJsonObject").mockResolvedValue(ok(claims));
    const graphAdapter: AxiosAdapter = async (config) => ({
      data: { webUrl: tenantUrl },
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
    };
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
});
