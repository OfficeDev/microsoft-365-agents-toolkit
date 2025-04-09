// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import path from "path";
import os from "os";
import fs from "fs";
import { LocalDebugTaskLabel, TemplateProject } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";
class FoodCatalogTestCase extends CaseFactory {
  override async onAfterCreate(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev"
  ): Promise<void> {
    // create folder for the test "/env/.env.dev"
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
    const envContent = `NOTIFICATION_ENDPOINT=https://test.com\nNOTIFICATION_DOMAIN=test.com\nAPP_NAME=${sampledebugContext.appName}`;
    fs.writeFileSync(filePath, envContent, { encoding: "utf-8" });
    console.log("env file created");
    console.log(fs.readFileSync(filePath, { encoding: "utf-8" }));
    // add chmod +x to the script
    if (os.platform() === "linux" || os.platform() === "darwin") {
      const scriptPath = path.resolve(
        sampledebugContext.projectPath,
        "scripts",
        "devtunnel.sh"
      );
      fs.chmodSync(scriptPath, "755");
    }
  }

  override async onProvision(
    sampledebugContext: SampledebugContext
  ): Promise<void> {
    return await sampledebugContext.provisionProject(
      sampledebugContext.appName,
      sampledebugContext.projectPath,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      "@azure/data-tables"
    );
  }
}

new FoodCatalogTestCase(
  TemplateProject.FoodCatalog,
  undefined,
  27851823,
  "v-ivanchen@microsoft.com",
  [
    LocalDebugTaskLabel.Azurite,
    LocalDebugTaskLabel.EnsureDevTunnnel,
    LocalDebugTaskLabel.RunWatch,
    LocalDebugTaskLabel.FuncStart,
  ],
  {
    skipInit: true,
    repoPath: "./resource/samples",
    testRootFolder: path.resolve(os.homedir(), "resource"),
    skipLocal: true,
  }
).test();
