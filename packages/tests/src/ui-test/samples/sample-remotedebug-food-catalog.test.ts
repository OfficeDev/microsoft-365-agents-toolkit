// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import path from "path";
import os from "os";
import fs from "fs";
import { TemplateProject } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";

class FoodCatalogTestCase extends CaseFactory {
  override async onAfterCreate(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev"
  ): Promise<void> {
    console.log("pre provision project");
    await sampledebugContext.createEnvFolder(
      sampledebugContext.projectPath,
      "env"
    );
    // create .env file
    const filePath = path.resolve(
      sampledebugContext.projectPath,
      "env",
      `.env.${env}`
    );
    const envContent = `TEAMSFX_ENV=dev\nAPP_NAME=${sampledebugContext.appName}`;
    fs.writeFileSync(filePath, envContent, { encoding: "utf-8" });
    console.log("env file created");
  }
}

new FoodCatalogTestCase(
  TemplateProject.FoodCatalog,
  27851823,
  "v-ivanchen@microsoft.com",
  "dev",
  [],
  {
    skipInit: true,
    repoPath: "./resource/samples",
    testRootFolder: path.resolve(os.homedir(), "resource"),
  }
).test();
