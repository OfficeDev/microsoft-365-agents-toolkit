// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { validateQueryOrg, reopenPage } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { Env } from "../../utils/env";
import { SampledebugContext } from "./sampledebugContext";

class QueryOrgTestCase extends CaseFactory {
  override async onValidate(
    page: Page,
    options?: { context: SampledebugContext }
  ): Promise<void> {
    return await validateQueryOrg(page, {
      displayName: Env.displayName,
      appName: options?.context.appName.substring(0, 10) || "",
    });
  }
  override async onCliValidate(
    page: Page,
    options?: {
      context: SampledebugContext;
    }
  ): Promise<void> {
    return await validateQueryOrg(page, {
      displayName: Env.displayName,
      appName: options?.context.appName.substring(0, 10) || "",
    });
  }
  public override async onReopenPage(
    sampledebugContext: SampledebugContext,
    teamsAppId: string
  ): Promise<Page> {
    return await reopenPage(sampledebugContext.context!, teamsAppId);
  }
}

new QueryOrgTestCase(
  TemplateProject.QueryOrg,
  15554404,
  24121481,
  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.StartLocalTunnel, LocalDebugTaskLabel.StartBot]
).test();
