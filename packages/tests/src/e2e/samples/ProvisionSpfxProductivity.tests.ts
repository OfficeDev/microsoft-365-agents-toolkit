// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Huihui Wu <huihuiwu@microsoft.com>
 */

import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";
import * as fs from "fs-extra";
import * as path from "path";
import { expect } from "chai";

class SpfxProductivityTestCase extends CaseFactory {
  override async onAfterCreate(projectPath: string): Promise<void> {
    expect(fs.pathExistsSync(path.resolve(projectPath, "src", "src"))).to.be
      .true;
  }
}

new SpfxProductivityTestCase(
  TemplateProjectFolder.SpfxProductivity,
  24753056,
  "huihuiwu@microsoft.com",
  ["spfx"]
).test();
