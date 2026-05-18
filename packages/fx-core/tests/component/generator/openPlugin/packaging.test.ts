// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Platform } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import { expect } from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as os from "os";
import * as path from "path";
import { featureFlagManager, FeatureFlags } from "../../../../src/common/featureFlags";
import { CreateAppPackageDriver } from "../../../../src/component/driver/teamsApp/createAppPackage";
import { CreateAppPackageArgs } from "../../../../src/component/driver/teamsApp/interfaces/CreateAppPackageArgs";
import { convertOpenPlugin } from "../../../../src/component/generator/openPlugin/generator";
import { MockedM365Provider } from "../../../core/utils";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";

async function tmp(prefix: string): Promise<string> {
  return await fs.mkdtemp(path.join(os.tmpdir(), prefix));
}

async function seedSamplePlugin(root: string): Promise<void> {
  await fs.ensureDir(path.join(root, ".plugin"));
  await fs.writeJSON(path.join(root, ".plugin", "plugin.json"), {
    name: "pkg-demo",
    version: "1.0.0",
    description: "demo",
    homepage: "https://example.com",
    author: { name: "Demo" },
  });
  await fs.ensureDir(path.join(root, "skills", "hello"));
  await fs.writeFile(
    path.join(root, "skills", "hello", "SKILL.md"),
    "---\nname: hello\ndescription: hi\n---\nbody"
  );
  await fs.ensureDir(path.join(root, "skills", "hello", "nested"));
  await fs.writeFile(path.join(root, "skills", "hello", "nested", "helper.md"), "# nested helper");
}

describe("openPlugin → teamsApp/zipAppPackage end-to-end", () => {
  const driver = new CreateAppPackageDriver();
  let pluginDir: string;
  let projectDir: string;
  let envRestore: RestoreFn | undefined;

  beforeEach(async () => {
    pluginDir = await tmp("op-pkg-plugin-");
    projectDir = await tmp("op-pkg-proj-");
    await fs.remove(projectDir);
    await seedSamplePlugin(pluginDir);
    envRestore = mockedEnv({
      TEAMSFX_ENV: "dev",
      TEAMS_APP_ID: "00000000-0000-0000-0000-000000000000",
    });
  });

  afterEach(async () => {
    await fs.remove(pluginDir);
    await fs.remove(projectDir);
    if (envRestore) {
      envRestore();
      envRestore = undefined;
    }
  });

  it("zips skill folders when TEAMSFX_AGENT_SKILLS is on", async () => {
    const convertRes = await convertOpenPlugin({
      path: pluginDir,
      output: projectDir,
      privacyUrl: "https://example.com/privacy",
      termsUrl: "https://example.com/terms",
    });
    if (convertRes.isErr()) throw new Error(convertRes.error.message);

    const wasEnabled = featureFlagManager.getBooleanValue(FeatureFlags.AgentSkillsManifest);
    featureFlagManager.setBooleanValue(FeatureFlags.AgentSkillsManifest, true);
    try {
      const args: CreateAppPackageArgs = {
        manifestPath: path.join(projectDir, "appPackage", "manifest.json"),
        outputZipPath: path.join(projectDir, "appPackage", "build", "appPackage.dev.zip"),
        outputFolder: path.join(projectDir, "appPackage", "build"),
      };
      const ctx: any = {
        m365TokenProvider: new MockedM365Provider(),
        projectPath: projectDir,
        platform: Platform.CLI,
        logProvider: new MockedLogProvider(),
        ui: new MockedUserInteraction(),
        addTelemetryProperties: () => {},
      };
      const buildRes = (await driver.execute(args, ctx)).result;
      if (buildRes.isErr()) throw new Error(buildRes.error.message);

      const zip = new AdmZip(args.outputZipPath);
      const entries = zip.getEntries().map((e) => e.entryName);
      expect(entries).to.include("manifest.json");
      expect(entries).to.include("color.png");
      expect(entries).to.include("outline.png");
      expect(
        entries.some((e) => e === "skills/hello/SKILL.md" || e === "skills\\hello\\SKILL.md")
      ).to.equal(true);
      expect(
        entries.some(
          (e) => e === "skills/hello/nested/helper.md" || e === "skills\\hello\\nested\\helper.md"
        )
      ).to.equal(true);
    } finally {
      featureFlagManager.setBooleanValue(FeatureFlags.AgentSkillsManifest, wasEnabled);
    }
  });

  it("does NOT zip skill folders when TEAMSFX_AGENT_SKILLS is off", async () => {
    const convertRes = await convertOpenPlugin({
      path: pluginDir,
      output: projectDir,
      privacyUrl: "https://example.com/privacy",
      termsUrl: "https://example.com/terms",
    });
    if (convertRes.isErr()) throw new Error(convertRes.error.message);

    const wasEnabled = featureFlagManager.getBooleanValue(FeatureFlags.AgentSkillsManifest);
    featureFlagManager.setBooleanValue(FeatureFlags.AgentSkillsManifest, false);
    try {
      const args: CreateAppPackageArgs = {
        manifestPath: path.join(projectDir, "appPackage", "manifest.json"),
        outputZipPath: path.join(projectDir, "appPackage", "build", "appPackage.dev.zip"),
        outputFolder: path.join(projectDir, "appPackage", "build"),
      };
      const ctx: any = {
        m365TokenProvider: new MockedM365Provider(),
        projectPath: projectDir,
        platform: Platform.CLI,
        logProvider: new MockedLogProvider(),
        ui: new MockedUserInteraction(),
        addTelemetryProperties: () => {},
      };
      const buildRes = (await driver.execute(args, ctx)).result;
      if (buildRes.isErr()) throw new Error(buildRes.error.message);

      const zip = new AdmZip(args.outputZipPath);
      const entries = zip.getEntries().map((e) => e.entryName);
      expect(entries.some((e) => e.startsWith("skills/") || e.startsWith("skills\\"))).to.equal(
        false
      );
    } finally {
      featureFlagManager.setBooleanValue(FeatureFlags.AgentSkillsManifest, wasEnabled);
    }
  });
});
