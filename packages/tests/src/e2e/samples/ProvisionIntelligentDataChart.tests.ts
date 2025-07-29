// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";
import * as path from "path";
import { editDotEnvFile } from "../commonUtils";
import { getUuid } from "../../commonlib/utilities";

class IntelligentDataChartCase extends CaseFactory {
  override async onBeforeProvision(projectPath: string): Promise<void> {
    const envFilePath = path.resolve(projectPath, "env", ".env.dev.user");
    editDotEnvFile(envFilePath, "SQL_USER", "Abc123321");
    editDotEnvFile(
      envFilePath,
      "SECRET_SQL_PASSWORD",
      "Cab232332" + getUuid().substring(0, 6)
    );
    editDotEnvFile(envFilePath, "SQL_SERVER", "https://test.com");
    editDotEnvFile(envFilePath, "SQL_DATABASE", "fake");
    editDotEnvFile(envFilePath, "SECRET_OPENAI_ENDPOINT", "https://test.com");
    editDotEnvFile(envFilePath, "SECRET_OPENAI_DEPLOYMENT_NAME", "fake");
    editDotEnvFile(envFilePath, "SECRET_OPENAI_API_KEY", "fake");
  }
}

new IntelligentDataChartCase(
  TemplateProjectFolder.IntelligentDataChart,
  27852475,
  "Ning.Tang@microsoft.com"
).test();
