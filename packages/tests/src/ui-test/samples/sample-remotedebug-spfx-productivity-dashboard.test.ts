// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { initTeamsPage } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";
import { Env } from "../../utils/env";

class SpfxProductivityTestCase extends CaseFactory {
  public override async onInitPage(
    sampledebugContext: SampledebugContext,
    teamsAppId: string,
    options?: {
      teamsAppName: string;
      type: string;
    }
  ): Promise<Page> {
    return await initTeamsPage(
      sampledebugContext.context!,
      teamsAppId,
      Env.username,
      Env.password,
      {
        projectPath: sampledebugContext.projectPath,
        env: "dev",
        teamsAppName: options?.teamsAppName,
        type: options?.type,
      }
    );
  }
}

new SpfxProductivityTestCase(
  TemplateProject.SpfxProductivity,
  24753065,
  "v-ivanchen@microsoft.com",
  "dev",
  [LocalDebugTaskLabel.GulpServe],
  {
    teamsAppName: "SPFx productivity dashboard",
    skipValidation: true,
  }
).test();
