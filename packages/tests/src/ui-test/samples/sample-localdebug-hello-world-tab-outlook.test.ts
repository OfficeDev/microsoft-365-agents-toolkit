// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

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
}

new OutlookTabTestCase(
  TemplateProject.OutlookTab,
  17451443,
  "v-ivanchen@microsoft.com",
  "local",
  [LocalDebugTaskLabel.StartFrontend],
  { debug: "cli" }
).test();
