// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";
import { AzSqlHelper } from "../../utils/azureCliHelper";
import { validateIntelligentDataChart } from "../../utils/playwrightOperation";
import { SampledebugContext } from "./sampledebugContext";
import { expect } from "chai";
import * as path from "path";
import { editDotEnvFile } from "../../utils/commonUtils";
import { OpenAiKey } from "../../utils/env";
import { Page } from "playwright";

const isRealKey = OpenAiKey.azureOpenAiKey ? true : false;
class IntelligentDataChartTestCase extends CaseFactory {
  public override async onBefore(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev",
    azSqlHelper?: AzSqlHelper | undefined
  ): Promise<AzSqlHelper | undefined> {
    if (isRealKey) {
      // create sql db server
      const rgName = `${sampledebugContext.appName}-dev-rg`;
      const sqlCommands = [
        `CREATE TABLE [SalesOrderDetail](
    [SalesOrderID] [int] NOT NULL,
    [SalesOrderDetailID] [int] IDENTITY(1,1) NOT NULL,
    [OrderQty] [smallint] NOT NULL,
    [ProductID] [int] NOT NULL,
    [UnitPrice] [money] NOT NULL,
    [UnitPriceDiscount] [money] NOT NULL,
    [LineTotal]  AS (isnull(([UnitPrice]*((1.0)-[UnitPriceDiscount]))*[OrderQty],(0.0))),
    [rowguid] [uniqueidentifier] NOT NULL,
    [ModifiedDate] [datetime] NOT NULL,
   CONSTRAINT [PK_SalesOrderDetail_SalesOrderID_SalesOrderDetailID] PRIMARY KEY CLUSTERED 
  (
    [SalesOrderID] ASC,
    [SalesOrderDetailID] ASC
  )WITH (STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
  ) ON [PRIMARY]`,
        `SET IDENTITY_INSERT [SalesOrderDetail] ON;
        INSERT [SalesOrderDetail] ([SalesOrderID], [SalesOrderDetailID], [OrderQty], [ProductID], [UnitPrice], [UnitPriceDiscount], [rowguid], [ModifiedDate]) VALUES (71774, 110562, 1, 836, 356.8980, 0.0000, N'e3a1994c-7a68-4ce8-96a3-77fdd3bbd730', CAST(N'2023-04-18T19:39:54.000' AS DateTime));`,
        `SET IDENTITY_INSERT [SalesOrderDetail] ON;
        INSERT [SalesOrderDetail] ([SalesOrderID], [SalesOrderDetailID], [OrderQty], [ProductID], [UnitPrice], [UnitPriceDiscount], [rowguid], [ModifiedDate]) VALUES (71782, 110697, 2, 951, 242.9940, 0.0000, N'35d889d9-676a-4b95-a2ea-28da743c25a7', CAST(N'2023-11-30T14:22:33.000' AS DateTime));`,
      ];
      azSqlHelper = new AzSqlHelper(rgName, sqlCommands);
      return azSqlHelper;
    }
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
    if (isRealKey) {
      const res = await azSqlHelper?.createSql();
      expect(res).to.be.true;
    }
    const envFilePath = path.resolve(
      sampledebugContext.projectPath,
      "env",
      ".env.local.user"
    );
    editDotEnvFile(envFilePath, "SQL_USER", azSqlHelper?.sqlAdmin ?? "fake");
    editDotEnvFile(
      envFilePath,
      "SECRET_SQL_PASSWORD",
      azSqlHelper?.sqlPassword ?? "fake@123"
    );
    editDotEnvFile(
      envFilePath,
      "SQL_SERVER",
      azSqlHelper?.sqlEndpoint ?? "https://test.com"
    );
    editDotEnvFile(
      envFilePath,
      "SQL_DATABASE",
      azSqlHelper?.sqlDatabaseName ?? "fake"
    );

    const azureOpenAiKey = OpenAiKey.azureOpenAiKey
      ? OpenAiKey.azureOpenAiKey
      : "fake";
    const azureOpenAiEndpoint = OpenAiKey.azureOpenAiEndpoint
      ? OpenAiKey.azureOpenAiEndpoint
      : "https://test.com";
    const azureOpenAiModelDeploymentName =
      OpenAiKey.azureOpenAiModelDeploymentName
        ? OpenAiKey.azureOpenAiModelDeploymentName
        : "fake";
    editDotEnvFile(envFilePath, "SECRET_OPENAI_API_KEY", azureOpenAiKey);
    editDotEnvFile(envFilePath, "SECRET_OPENAI_ENDPOINT", azureOpenAiEndpoint);
    editDotEnvFile(
      envFilePath,
      "SECRET_OPENAI_DEPLOYMENT_NAME",
      azureOpenAiModelDeploymentName
    );
  }
  override async onValidate(page: Page): Promise<void> {
    return await validateIntelligentDataChart(page, isRealKey);
  }
}

new IntelligentDataChartTestCase(
  TemplateProject.IntelligentDataChart,
  26461509,
  27852477,
  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.WatchBackend, LocalDebugTaskLabel.StartBackend]
).test();
