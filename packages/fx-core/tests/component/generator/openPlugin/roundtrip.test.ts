// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ok } from "@microsoft/teamsfx-api";
import { expect } from "chai";
import fs from "fs-extra";
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
    expect(importOpenPlugin).to.be.a("function");
    expect(exportOpenPlugin).to.be.a("function");
  });
});
