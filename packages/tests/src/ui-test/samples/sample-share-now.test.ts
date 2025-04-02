// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { validateShareNow } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { AzSqlHelper } from "../../utils/azureCliHelper";
import { SampledebugContext } from "./sampledebugContext";
import { expect } from "chai";
import * as path from "path";
import { editDotEnvFile } from "../../utils/commonUtils";

class ShareNowTestCase extends CaseFactory {
  public override async onBefore(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev",
    azSqlHelper?: AzSqlHelper | undefined
  ): Promise<AzSqlHelper | undefined> {
    // create sql db server
    const rgName = `${sampledebugContext.appName}-dev-rg`;
    const sqlCommands = [
      `CREATE TABLE [TeamPostEntity](
          [PostID] [int] PRIMARY KEY IDENTITY,
          [ContentUrl] [nvarchar](400) NOT NULL,
          [CreatedByName] [nvarchar](50) NOT NULL,
          [CreatedDate] [datetime] NOT NULL,
          [Description] [nvarchar](500) NOT NULL,
          [IsRemoved] [bit] NOT NULL,
          [Tags] [nvarchar](100) NULL,
          [Title] [nvarchar](100) NOT NULL,
          [TotalVotes] [int] NOT NULL,
          [Type] [int] NOT NULL,
          [UpdatedDate] [datetime] NOT NULL,
          [UserID] [uniqueidentifier] NOT NULL
      );`,
      `CREATE TABLE [UserVoteEntity](
          [VoteID] [int] PRIMARY KEY IDENTITY,
          [PostID] [int] NOT NULL,
          [UserID] [uniqueidentifier] NOT NULL
      );`,
    ];
    azSqlHelper = new AzSqlHelper(rgName, sqlCommands);
    return azSqlHelper;
  }
  override async onAfter(
    sampledebugContext: SampledebugContext
  ): Promise<void> {
    await sampledebugContext.sampleAfter(
      `${sampledebugContext.appName}-dev-rg}`
    );
  }
  public override async onAfterCreate(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev",
    azSqlHelper?: AzSqlHelper | undefined
  ): Promise<void> {
    const res = await azSqlHelper?.createSql();
    expect(res).to.be.true;
    const envFilePath = path.resolve(
      sampledebugContext.projectPath,
      "env",
      `.env.${env}.user`
    );
    editDotEnvFile(envFilePath, "SQL_USER_NAME", azSqlHelper?.sqlAdmin ?? "");
    editDotEnvFile(envFilePath, "SQL_PASSWORD", azSqlHelper?.sqlPassword ?? "");
    editDotEnvFile(envFilePath, "SQL_ENDPOINT", azSqlHelper?.sqlEndpoint ?? "");
    editDotEnvFile(
      envFilePath,
      "SQL_DATABASE_NAME",
      azSqlHelper?.sqlDatabaseName ?? ""
    );
  }
  override async onValidate(page: Page): Promise<void> {
    return await validateShareNow(page);
  }
}

new ShareNowTestCase(
  TemplateProject.ShareNow,
  9958523,
  24121485,
  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.StartFrontend, LocalDebugTaskLabel.StartBackend]
).test();
