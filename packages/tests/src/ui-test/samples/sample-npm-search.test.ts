// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { validateNpm } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";

class NpmSearchTestCase extends CaseFactory {
  override async onValidate(
    page: Page,
    options?: { npmName: string; context: SampledebugContext }
  ): Promise<void> {
    return await validateNpm(page, {
      npmName: options?.npmName,
      appName: options?.context.appName.substring(0, 10) || "",
    });
  }

  override async onCliValidate(
    page: Page,
    options?: { npmName: string; context: SampledebugContext }
  ): Promise<void> {
    return await validateNpm(page, {
      npmName: options?.npmName,
      appName: options?.context.appName.substring(0, 10) || "",
    });
  }
}

new NpmSearchTestCase(
  TemplateProject.NpmSearch,

  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.StartLocalTunnel, LocalDebugTaskLabel.StartBotApp],
  {
    npmName: "axios",
    testPlanCaseId_local: 12664761,
    testPlanCaseId_dev: 14571879,
  }
).test();
