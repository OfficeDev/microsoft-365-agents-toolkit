// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import {
  TemplateProject,
  LocalDebugTaskLabel,
  ValidationContent,
} from "../../utils/constants";
import { validateBot } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";

class GraphConnectorBotTestCase extends CaseFactory {
  override async onValidate(page: Page): Promise<void> {
    return await validateBot(page, {
      botCommand: "welcome",
      expected: ValidationContent.GraphBot,
    });
  }
  override async onCliValidate(page: Page): Promise<void> {
    return await validateBot(page, {
      botCommand: "welcome",
      expected: ValidationContent.GraphBot,
    });
  }
}

new GraphConnectorBotTestCase(
  TemplateProject.GraphConnectorBot,
  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.StartLocalTunnel, LocalDebugTaskLabel.StartApplication],
  {
    testPlanCaseId_local: 25178457,
    testPlanCaseId_dev: 25960851,
  }
  //{ debug: "cli" }
).test();
