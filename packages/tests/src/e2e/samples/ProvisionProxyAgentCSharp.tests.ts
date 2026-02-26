// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author quke@microsoft.com
 */

// ProxyAgent-CSharp is a C# solution whose ATK project (m365agents.yml and
// infra) lives inside the M365Agent/ subfolder. Full provision also requires
// a live Azure AI Foundry endpoint and OAuth App Registration, so provision
// is skipped in automated tests.

import { expect } from "chai";
import * as fs from "fs-extra";
import * as path from "path";
import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";

class ProxyAgentCSharpTestCase extends CaseFactory {
  public override async onAfterCreate(projectPath: string): Promise<void> {
    expect(fs.pathExistsSync(path.resolve(projectPath, "M365Agent", "infra")))
      .to.be.true;
  }
}

new ProxyAgentCSharpTestCase(
  TemplateProjectFolder.ProxyAgentCSharp,
  0,
  "quke@microsoft.com",
  ["bot"],
  { skipProvision: true }
).test();
