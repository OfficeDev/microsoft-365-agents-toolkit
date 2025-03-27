// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import {
  initTeamsPage,
  validateTodoList,
} from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { AzSqlHelper } from "../../utils/azureCliHelper";
import { SampledebugContext } from "./sampledebugContext";
import * as uuid from "uuid";
import * as fs from "fs";
import * as path from "path";
import { editDotEnvFile } from "../../utils/commonUtils";
import { Env } from "../../utils/env";
import * as os from "os";

class TodoListBackendTestCase extends CaseFactory {
  public sqlUserName: string;
  public sqlPassword: string;
  constructor(
    sampleName: TemplateProject,
    testPlanCaseId: number,
    author: string,
    env: "local" | "dev",
    validate: LocalDebugTaskLabel[] = [],
    options: {
      teamsAppName?: string;
      dashboardFlag?: boolean;
      type?: string;
      testRootFolder?: string;
      includeFunction?: boolean;
      npmName?: string;
      skipInit?: boolean;
      skipValidation?: boolean;
    } = {}
  ) {
    super(sampleName, testPlanCaseId, author, env, validate, options);
    this.sqlUserName = "Abc123321";
    this.sqlPassword = "Cab232332" + uuid.v4().substring(0, 6);
  }
  public override onAfterCreate = async (
    sampledebugContext: SampledebugContext
  ): Promise<void> => {
    const envFilePath = path.resolve(
      sampledebugContext.projectPath,
      "env",
      ".env.dev.user"
    );
    const user = this.sqlUserName;
    const password = this.sqlPassword;
    editDotEnvFile(envFilePath, "SQL_USER_NAME", user);
    editDotEnvFile(envFilePath, "SQL_PASSWORD", password);
  };
  public override onBeforeBrowerStart = async (
    sampledebugContext: SampledebugContext
  ): Promise<void> => {
    const user = this.sqlUserName;
    const password = this.sqlPassword;
    // read database from devEnvFilePath
    const devEnvFilePath = path.resolve(
      sampledebugContext.projectPath,
      "env",
      ".env.dev"
    );
    const sqlDatabaseName = fs
      .readFileSync(devEnvFilePath, "utf-8")
      .split("\n")
      .find((line) =>
        line.startsWith("PROVISIONOUTPUT__AZURESQLOUTPUT__DATABASENAME")
      )
      ?.split("=")[1];
    const sqlEndpoint = fs
      .readFileSync(devEnvFilePath, "utf-8")
      .split("\n")
      .find((line) =>
        line.startsWith("PROVISIONOUTPUT__AZURESQLOUTPUT__SQLENDPOINT")
      )
      ?.split("=")[1];

    const sqlCommands = [
      `CREATE TABLE Todo
        (
            id INT IDENTITY PRIMARY KEY,
            description NVARCHAR(128) NOT NULL,
            objectId NVARCHAR(36),
            channelOrChatId NVARCHAR(128),
            isCompleted TinyInt NOT NULL default 0,
        )`,
    ];
    const sqlHelper = new AzSqlHelper(
      `${sampledebugContext.appName}-dev-rg`,
      sqlCommands,
      sqlDatabaseName,
      sqlDatabaseName,
      user,
      password
    );
    await sqlHelper.createTable(sqlEndpoint ?? "");
  };
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
  override async onValidate(page: Page): Promise<void> {
    return await validateTodoList(page);
  }
}

new TodoListBackendTestCase(
  TemplateProject.TodoListBackend,
  14571882,
  "v-ivanchen@microsoft.com",
  "dev",
  [],
  {
    teamsAppName: "toDoList-dev",
    testRootFolder: path.resolve(os.homedir(), "resourse"), // fix eslint error
    type: "spfx",
  }
).test();
