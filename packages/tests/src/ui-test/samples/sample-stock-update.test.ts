// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { validateStockUpdate } from "../../utils/playwrightOperation";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";
import * as path from "path";
import * as fs from "fs";

class StockUpdateTestCase extends CaseFactory {
  public override async onAfterCreate(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev"
  ): Promise<void> {
    const envFile = path.resolve(
      sampledebugContext.projectPath,
      "env",
      `.env.${env}`
    );
    let ENDPOINT = fs.readFileSync(envFile, "utf-8");
    ENDPOINT +=
      "\nTEAMSFX_API_ALPHAVANTAGE_ENDPOINT=https://www.alphavantage.co";
    fs.writeFileSync(envFile, ENDPOINT);
    console.log(`add endpoint ${ENDPOINT} to .env.${env} file`);
    const userFile = path.resolve(
      sampledebugContext.projectPath,
      "env",
      `.env.${env}.user`
    );
    let KEY = fs.readFileSync(userFile, "utf-8");
    KEY += "\nTEAMSFX_API_ALPHAVANTAGE_API_KEY=demo";
    fs.writeFileSync(userFile, KEY);
    console.log(`add key ${KEY} to .env.${env}.user file`);
  }
  override async onValidate(page: Page): Promise<void> {
    return await validateStockUpdate(page);
  }
}

new StockUpdateTestCase(
  TemplateProject.StockUpdate,
  17303802,
  24121504,
  "v-ivanchen@microsoft.com",
  [
    LocalDebugTaskLabel.StartLocalTunnel,
    LocalDebugTaskLabel.Azurite,
    LocalDebugTaskLabel.Compile,
    LocalDebugTaskLabel.StartBotApp,
  ]
).test();
