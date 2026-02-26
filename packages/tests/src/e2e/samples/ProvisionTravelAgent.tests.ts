// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author quke@microsoft.com
 */

// travel-agent is a pure C# .NET solution (TravelAgent.slnx) with no
// m365agents.yml or ATK infra folder. Only project download is validated.

import { expect } from "chai";
import * as fs from "fs-extra";
import * as path from "path";
import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";

class TravelAgentTestCase extends CaseFactory {
  public override async onAfterCreate(projectPath: string): Promise<void> {
    expect(fs.pathExistsSync(path.resolve(projectPath, "TravelAgent.slnx"))).to
      .be.true;
  }
}

new TravelAgentTestCase(
  TemplateProjectFolder.TravelAgent,
  0,
  "quke@microsoft.com",
  [],
  { skipProvision: true }
).test();
