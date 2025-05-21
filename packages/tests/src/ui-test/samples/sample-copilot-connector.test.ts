// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { validateCopilotConnector } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { Env } from "../../utils/env";

class CopilotConnectorTestCase extends CaseFactory {
  override async onValidate(page: Page): Promise<void> {
    return await validateCopilotConnector(page, {
      displayName: Env.displayName,
    });
  }
}

new CopilotConnectorTestCase(
  TemplateProject.CopilotConnector,
  "v-ivanchen@microsoft.com",
  [
    // [BUG] warning error message block the frontend validation
    // LocalDebugTaskLabel.StartFrontend,
    LocalDebugTaskLabel.WatchBackend,
    LocalDebugTaskLabel.StartBackend,
  ],
  {
    testPlanCaseId_local: 14171510,
    testPlanCaseId_dev: 14571877,
  }
).test();
