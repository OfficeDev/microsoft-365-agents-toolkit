// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import * as fs from "fs-extra";
import * as path from "path";
import { SampledebugContext } from "./sampledebugContext";
import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { validateOneProducitvity } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { Env } from "../../utils/env";

class OneProductivityHubTestCase extends CaseFactory {
  override async onValidate(
    page: Page,
    option?: { displayName: string }
  ): Promise<void> {
    return await validateOneProducitvity(page, {
      displayName: Env.displayName,
    });
  }
  override async onAfterCreate(
    sampledebugContext: SampledebugContext
  ): Promise<void> {
    // update swa sku to standard
    const bicepPath = path.join(
      sampledebugContext.projectPath,
      "infra",
      "azure.parameters.json"
    );
    const bicep = fs.readJsonSync(bicepPath);
    bicep["parameters"]["staticWebAppSku"]["value"] = "Standard";
    fs.writeJsonSync(bicepPath, bicep);
  }
}

new OneProductivityHubTestCase(
  TemplateProject.OneProductivityHub,
  15090375,
  24121468,
  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.StartFrontend]
).test();
