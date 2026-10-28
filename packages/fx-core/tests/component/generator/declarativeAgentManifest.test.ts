// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AppManifestUtils, DeclarativeAgentManifestConverter } from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { globSync } from "glob";
import Mustache from "mustache";
import path from "path";
import { assert } from "vitest";

const templatesRoot = path.resolve(__dirname, "../../../../../templates");
const templates = globSync("{vsc,vs,v4}/**/{declarativeAgent,repairDeclarativeAgent}.json.tpl", {
  cwd: templatesRoot,
})
  .map((file) => file.split(path.sep).join("/"))
  .sort();

describe("Declarative agent manifest templates", () => {
  it("SCN-CREATE-SKILL-08: covers legacy, v4, Visual Studio, and repair templates", () => {
    for (const prefix of ["vsc/", "vs/", "v4/"]) {
      assert.isTrue(
        templates.some((file) => file.startsWith(prefix)),
        prefix
      );
    }
    assert.isTrue(templates.some((file) => file.endsWith("/repairDeclarativeAgent.json.tpl")));
  });

  it.each(templates)("SCN-CREATE-SKILL-08: %s renders a valid v1.9 manifest", async (file) => {
    const source = await fs.readFile(path.join(templatesRoot, file), "utf8");
    const rendered = Mustache.render(source, { appName: "TestAgent" });
    const agent = DeclarativeAgentManifestConverter.jsonToManifest(rendered);

    assert.equal(agent.version, "v1.9");
    assert.equal(
      agent.$schema,
      "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.9/schema.json"
    );
    assert.deepEqual(await AppManifestUtils.validateAgainstSchema(agent), []);
    assert.deepEqual(
      JSON.parse(DeclarativeAgentManifestConverter.manifestToJson(agent)),
      JSON.parse(rendered)
    );
  });
});
