// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

// this test case diposed due to cancel local debug
import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";
import { validateLargeNotificationBot } from "../../utils/playwrightOperation";
import { getBotSiteEndpoint } from "../../utils/commonUtils";

class LargeNotiTestCase extends CaseFactory {
  override async onValidate(
    page: Page,
    options: {
      context: SampledebugContext;
      displayName: string;
      includeFunction: boolean;
      npmName: string;
      env: "local" | "dev";
    }
  ): Promise<void> {
    const funcEndpoint = await getBotSiteEndpoint(
      options.context.projectPath,
      options.env
    );
    return await validateLargeNotificationBot(
      page,
      funcEndpoint + "/api/notification"
    );
  }
}

new LargeNotiTestCase(
  TemplateProject.LargeScaleBot,
  "v-ivanchen@microsoft.com",
  [
    LocalDebugTaskLabel.StartLocalTunnel,
    LocalDebugTaskLabel.Compile,
    LocalDebugTaskLabel.Azurite,
    LocalDebugTaskLabel.StartApplication,
  ],
  {
    skipLocal: true,
    testPlanCaseId_dev: 25960873,
  }
).test();
