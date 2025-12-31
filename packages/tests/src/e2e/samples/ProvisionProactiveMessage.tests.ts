// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { TemplateProjectFolder } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";
import { Executor } from "../../utils/executor";
import { setBotSkuNameToB1Bicep } from "../commonUtils";

class ProactiveMessagingTestCase extends CaseFactory {
  public override async onCreate(
    appName: string,
    testFolder: string,
    sampleName: TemplateProjectFolder
  ): Promise<void> {
    // The sample has nodejs/ and python/ subdirectories
    // Use the nodejs subfolder path directly as specified in samples-config downloadUrlInfo.dir
    await Executor.openTemplateProject(
      appName,
      testFolder,
      "nodejs" as TemplateProjectFolder,
      undefined,
      `samples/${sampleName}`
    );
  }

  override async onBeforeProvision(projectPath: string): Promise<void> {
    await setBotSkuNameToB1Bicep(
      projectPath,
      "templates/azure/azure.parameters.dev.json"
    );
  }

  override async onAfterCreate(projectPath: string): Promise<void> {
    return Promise.resolve();
  }
}

new ProactiveMessagingTestCase(
  TemplateProjectFolder.ProactiveMessaging,
  15277473,
  "ning.tang@microsoft.com"
).test();
