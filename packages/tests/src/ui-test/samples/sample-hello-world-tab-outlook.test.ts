// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import * as fs from "fs-extra";
import * as path from "path";
import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import {
  validatePersonalTab,
  reopenPage,
} from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";

class OutlookTabTestCase extends CaseFactory {
  override async onValidate(page: Page): Promise<void> {
    return await validatePersonalTab(page);
  }
  override async onCliValidate(page: Page): Promise<void> {
    return await validatePersonalTab(page);
  }
  public override async onReopenPage(
    sampledebugContext: SampledebugContext,
    teamsAppId: string
  ): Promise<Page> {
    return await reopenPage(
      sampledebugContext.context!,
      teamsAppId,
      undefined,
      undefined,
      {
        projectPath: sampledebugContext.projectPath,
        env: "local",
      }
    );
  }
  override async onAfterCreate(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev"
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

new OutlookTabTestCase(
  TemplateProject.OutlookTab,

  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.StartFrontend],
  {
    testPlanCaseId_local: 17451443,
    testPlanCaseId_dev: 24121457,
  }
  //{ debug: "cli" }
).test();
