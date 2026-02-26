// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author quke@microsoft.com
 */

import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";
import * as path from "path";
import { editDotEnvFile } from "../commonUtils";

class DataAnalystAgentV2TestCase extends CaseFactory {
  override async onBeforeProvision(projectPath: string): Promise<void> {
    const envFilePath = path.resolve(projectPath, "env", ".env.dev.user");
    editDotEnvFile(
      envFilePath,
      "SECRET_AZURE_OPENAI_API_KEY",
      "fake-openai-key"
    );
    editDotEnvFile(
      envFilePath,
      "AZURE_OPENAI_ENDPOINT",
      "https://fake.openai.azure.com/"
    );
    editDotEnvFile(
      envFilePath,
      "AZURE_OPENAI_DEPLOYMENT_NAME",
      "fake-deployment"
    );
  }
}

new DataAnalystAgentV2TestCase(
  TemplateProjectFolder.DataAnalystAgentV2,
  0,
  "quke@microsoft.com",
  ["bot"]
).test();
