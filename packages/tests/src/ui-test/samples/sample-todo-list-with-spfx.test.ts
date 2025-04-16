// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import {
  initTeamsPage,
  validateTodoListSpfx,
} from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";
import { Env } from "../../utils/env";

class TodoListSpfxTestCase extends CaseFactory {
  public override async onInitPage(
    sampledebugContext: SampledebugContext,
    teamsAppId: string,
    options?: {
      teamsAppName: string;
      type: string;
      env: "local" | "dev";
    }
  ): Promise<Page> {
    return await initTeamsPage(
      sampledebugContext.context!,
      teamsAppId,
      Env.username,
      Env.password,
      {
        projectPath: sampledebugContext.projectPath,
        env: options?.env,
        teamsAppName: options?.teamsAppName,
        type: options?.type,
      }
    );
  }
  public override async onValidate(page: Page): Promise<void> {
    return await validateTodoListSpfx(page, "TodoLi");
  }
}

new TodoListSpfxTestCase(
  TemplateProject.TodoListSpfx,
  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.GulpServe],
  {
    teamsAppName: "TodoLi",
    type: "spfx",
    testPlanCaseId_local: 9958516,
    testPlanCaseId_dev: 24121511,
  }
).test();
