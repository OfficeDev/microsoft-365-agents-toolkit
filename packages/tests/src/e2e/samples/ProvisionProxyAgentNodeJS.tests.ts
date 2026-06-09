// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author quke@microsoft.com
 */

import * as path from "path";
import { editDotEnvFile } from "../commonUtils";
import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";

class ProxyAgentNodeJSTestCase extends CaseFactory {
  override async onBeforeProvision(projectPath: string): Promise<void> {
    // This sample is different so set the env in .env.dev other than .env.dev.user.
    const envFilePath = path.resolve(projectPath, "env", ".env.dev");
    editDotEnvFile(
      envFilePath,
      "AZURE_AI_FOUNDRY_PROJECT_ENDPOINT",
      "https://fake.services.ai.azure.com/api/projects/fake"
    );
    editDotEnvFile(envFilePath, "AGENT_ID", "fake-agent-id");
  }
}

new ProxyAgentNodeJSTestCase(
  TemplateProjectFolder.ProxyAgentNodeJS,
  0,
  "quke@microsoft.com",
  ["bot"]
).test();
