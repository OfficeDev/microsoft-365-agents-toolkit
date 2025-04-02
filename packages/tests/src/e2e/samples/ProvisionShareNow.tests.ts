// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";
import * as path from "path";
import { editDotEnvFile, setBotSkuNameToB1Bicep } from "../commonUtils";
import { getUuid } from "../../commonlib/utilities";

class ShareNowTestCase extends CaseFactory {
  override async onBeforeProvision(projectPath: string): Promise<void> {
    // fix quota issue
    await setBotSkuNameToB1Bicep(projectPath);
    const envFilePath = path.resolve(projectPath, "env", ".env.dev.user");
    editDotEnvFile(envFilePath, "SQL_USER_NAME", "Abc123321");
    editDotEnvFile(
      envFilePath,
      "SQL_PASSWORD",
      "Cab232332" + getUuid().substring(0, 6)
    );
  }
}

// Skip the test case for security reason
// new ShareNowTestCase(
//   TemplateProjectFolder.ShareNow,
//   15277467,
//   "zhaofengxu@microsoft.com",
//   ["sql", "tab & bot"]
// ).test();
