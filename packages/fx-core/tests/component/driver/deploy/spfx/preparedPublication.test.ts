// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import * as fs from "fs-extra";
import os from "os";
import path from "path";
import { afterEach, describe, expect, it } from "vitest";

import {
  appId,
  createAppPackage,
  createSppkg,
  productId,
  publicationTenantContext,
} from "./preparedPublicationFixtures";
import { preparePublication } from "../../../../../src/component/driver/deploy/spfx/utility/publicationArtifacts";

describe("SharePoint publication artifacts", () => {
  const directories: string[] = [];

  afterEach(async () => {
    await Promise.all(directories.splice(0).map((directory) => fs.remove(directory)));
  });

  it("SPFX-PREP-05: writes staged packages without a journal or source mutation", async () => {
    const projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "publication-artifacts-"));
    directories.push(projectPath);
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    const sourceSharePoint = createSppkg();
    const sourceApplication = createAppPackage();
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), sourceSharePoint);
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), sourceApplication);

    const result = await preparePublication({
      projectPath,
      packagePath: "sharepoint/solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext: publicationTenantContext,
    });

    expect(result.isErr() ? result.error : undefined).toBeUndefined();
    if (result.isErr()) return;
    expect(result.value).toEqual({
      sharePointPackagePath: expect.stringMatching(
        /^\.fx\/sharepoint-publications\/[0-9a-f-]+\/sharepoint\.sppkg$/
      ),
      appPackagePath: expect.stringMatching(
        /^\.fx\/sharepoint-publications\/[0-9a-f-]+\/application\.zip$/
      ),
    });
    expect(result.value.sharePointPackagePath).not.toBe(result.value.appPackagePath);
    expect(await fs.readFile(path.join(projectPath, "sharepoint", "solution.sppkg"))).toEqual(
      sourceSharePoint
    );
    expect(await fs.readFile(path.join(projectPath, "appPackage", "application.zip"))).toEqual(
      sourceApplication
    );

    const stagedSharePoint = new AdmZip(
      await fs.readFile(path.join(projectPath, result.value.sharePointPackagePath))
    );
    expect(stagedSharePoint.getEntry("ClientSideAssets/agent.zip")).toBeNull();
    expect(stagedSharePoint.readAsText("_rels/ClientSideAssets.xml.rels")).not.toContain(
      "agent.zip"
    );

    const stagedApplication = new AdmZip(
      await fs.readFile(path.join(projectPath, result.value.appPackagePath))
    );
    const manifest = JSON.parse(stagedApplication.readAsText("manifest.json"));
    expect(manifest.id).toBe(appId);
    const agent = JSON.parse(stagedApplication.readAsText("declarativeAgent.json"));
    expect(agent.instructions).toBe("Keep these instructions");
    const plugin = JSON.parse(stagedApplication.readAsText(agent.actions[0].file));
    expect(plugin.runtimes[0].spec.url).toBe(
      `https://tenant.example.test/_api/mcp/spfx?solutionId=${productId}`
    );

    const stagingDirectory = path.dirname(
      path.join(projectPath, result.value.sharePointPackagePath)
    );
    expect((await fs.readdir(stagingDirectory)).sort()).toEqual([
      "application.zip",
      "sharepoint.sppkg",
    ]);
  });

  it("SPFX-PREP-04: rejects invalid input paths without creating staging artifacts", async () => {
    const projectPath = await fs.mkdtemp(path.join(os.tmpdir(), "publication-invalid-input-"));
    directories.push(projectPath);
    await fs.ensureDir(path.join(projectPath, "sharepoint"));
    await fs.ensureDir(path.join(projectPath, "appPackage"));
    await fs.writeFile(path.join(projectPath, "sharepoint", "solution.sppkg"), createSppkg());
    await fs.writeFile(path.join(projectPath, "appPackage", "application.zip"), createAppPackage());

    const result = await preparePublication({
      projectPath,
      packagePath: "../solution.sppkg",
      appPackagePath: "appPackage/application.zip",
      tenantContext: publicationTenantContext,
    });

    expect(result.isErr()).toBe(true);
    if (result.isOk()) return;
    expect(result.error.code).toBe("InvalidInputPath");
    expect(await fs.pathExists(path.join(projectPath, ".fx"))).toBe(false);
  });
});
