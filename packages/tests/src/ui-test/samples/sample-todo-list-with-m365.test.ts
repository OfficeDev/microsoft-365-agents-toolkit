// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import * as uuid from "uuid";
import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { validateTodoList, reopenPage } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";
import { editDotEnvFile } from "../../utils/commonUtils";
import * as path from "path";
import * as fs from "fs";
import * as os from "os";

class TodoListM365TestCase extends CaseFactory {
  public override async onAfterCreate(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev"
  ): Promise<void> {
    if (env === "local") {
      const targetPath = path.resolve(sampledebugContext.projectPath, "tabs");
      const data = "src/";
      // create .eslintignore
      fs.writeFileSync(targetPath + "/.eslintignore", data);
    } else {
      const envFilePath = path.resolve(
        sampledebugContext.projectPath,
        "env",
        ".env.dev.user"
      );
      editDotEnvFile(envFilePath, "SQL_USER_NAME", "Abc123321");
      editDotEnvFile(
        envFilePath,
        "SQL_PASSWORD",
        "Cab232332" + uuid.v4().substring(0, 6)
      );
    }
  }
  override async onValidate(
    page: Page,
    options?: { displayName: string }
  ): Promise<void> {
    return await validateTodoList(page, { displayName: options?.displayName });
  }
  override async onCliValidate(
    page: Page,
    options?: { displayName: string }
  ): Promise<void> {
    return await validateTodoList(page, { displayName: options?.displayName });
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

new TodoListM365TestCase(
  TemplateProject.TodoListM365,

  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.StartFrontend, LocalDebugTaskLabel.StartBackend],
  {
    teamsAppName: "toDoList-",
    //debug: "cli",
    testRootFolder: path.resolve(os.homedir(), "resourse"), // fix eslint error
    testPlanCaseId_local: 12664741,

    testPlanCaseId_dev: 14571883,
  }
).test();
