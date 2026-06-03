// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ok } from "@microsoft/teamsfx-api";
import { expect } from "chai";
import fs from "fs-extra";
import "mocha";
import * as os from "os";
import * as path from "path";
import sinon from "sinon";
import { setTools } from "../../../../src/common/globalVars";
import { Generator } from "../../../../src/component/generator/generator";
import { exportOpenPlugin } from "../../../../src/component/generator/openPlugin/exporter";
import { importOpenPlugin } from "../../../../src/component/generator/openPlugin/importer";
import { MockTools } from "../../../core/utils";
import { scaffoldOpenPluginTemplateFromSource } from "./testTemplateScaffold";

async function tmp(prefix: string): Promise<string> {
  return await fs.mkdtemp(path.join(os.tmpdir(), prefix));
}

async function seedSamplePlugin(root: string): Promise<void> {
  await fs.ensureDir(path.join(root, ".plugin"));
  await fs.writeJSON(path.join(root, ".plugin", "plugin.json"), {
    name: "demo-plugin",
    version: "1.2.3",
    description: "A demo Open Plugin for the round-trip test.",
    author: { name: "Jane Doe", url: "https://example.com" },
    homepage: "https://example.com",
  });
  await fs.writeJSON(path.join(root, ".mcp.json"), {
    mcpServers: { web: { url: "https://web.example.com/api", description: "remote" } },
  });
  await fs.ensureDir(path.join(root, "skills", "alpha-skill"));
  await fs.writeFile(
    path.join(root, "skills", "alpha-skill", "SKILL.md"),
    "---\nname: alpha-skill\ndescription: hi\n---\nbody"
  );
}

describe("openPlugin.roundtrip (import → export → import)", () => {
  setTools(new MockTools());
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    sandbox.stub(Generator, "generateTemplate").callsFake(async (ctx, dest) => {
      const appName = ctx.templateVariables?.appName ?? "";
      await scaffoldOpenPluginTemplateFromSource(dest, { appName });
      return ok(undefined);
    });
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("round-trips losslessly without needing --privacy-url/--terms-url the second time", async () => {
    const pluginDir = await tmp("op-rt-plugin-");
    const projectA = await tmp("op-rt-projA-");
    const exportedDir = await tmp("op-rt-exp-");
    const projectB = await tmp("op-rt-projB-");
    await fs.remove(projectA);
    await fs.remove(exportedDir);
    await fs.remove(projectB);
    try {
      await seedSamplePlugin(pluginDir);

      // 1. First import: requires --privacy-url + --terms-url (no extension block yet).
      const firstImport = await importOpenPlugin({
        path: pluginDir,
        output: projectA,
        privacyUrl: "https://example.com/privacy",
        termsUrl: "https://example.com/terms",
      });
      if (firstImport.isErr()) throw new Error(firstImport.error.message);
      const manifestA = await fs.readJSON(path.join(projectA, "appPackage", "manifest.json"));

      // 2. Export back to an open-plugin directory.
      const exp = await exportOpenPlugin({ path: projectA, output: exportedDir });
      if (exp.isErr()) throw new Error(exp.error.message);
      const exportedPlugin = await fs.readJSON(path.join(exportedDir, ".plugin", "plugin.json"));
      expect(exportedPlugin["x-microsoft-365-agents-toolkit"]).to.exist;
      expect(exportedPlugin["x-microsoft-365-agents-toolkit"].developer.privacyUrl).to.equal(
        "https://example.com/privacy"
      );

      // 3. Re-import the exported directory WITHOUT providing privacy/terms URLs.
      //    The mapper should pick them up from the extension block.
      const secondImport = await importOpenPlugin({
        path: exportedDir,
        output: projectB,
      });
      if (secondImport.isErr()) throw new Error(secondImport.error.message);
      const manifestB = await fs.readJSON(path.join(projectB, "appPackage", "manifest.json"));

      // Core identity fields must round-trip.
      expect(manifestB.id).to.equal(manifestA.id);
      expect(manifestB.developer.privacyUrl).to.equal(manifestA.developer.privacyUrl);
      expect(manifestB.developer.termsOfUseUrl).to.equal(manifestA.developer.termsOfUseUrl);
      expect(manifestB.developer.websiteUrl).to.equal(manifestA.developer.websiteUrl);
      expect(manifestB.name).to.deep.equal(manifestA.name);
      expect(manifestB.description).to.deep.equal(manifestA.description);
      expect(manifestB.agentSkills).to.deep.equal(manifestA.agentSkills);
      expect(manifestB.agentConnectors).to.deep.equal(manifestA.agentConnectors);
    } finally {
      await fs.remove(pluginDir);
      await fs.remove(projectA);
      await fs.remove(exportedDir);
      await fs.remove(projectB);
    }
  });
});
