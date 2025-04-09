// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { TemplateProject, LocalDebugTaskLabel } from "../../utils/constants";
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";
import * as path from "path";
import * as fs from "fs";

class IncomingWebhookTestCase extends CaseFactory {
  public override async onAfterCreate(
    sampledebugContext: SampledebugContext
  ): Promise<void> {
    // replace "<webhook-url>" to "https://test.com"
    console.log("replace webhook url");
    const targetFile = path.resolve(
      sampledebugContext.projectPath,
      "src",
      "index.ts"
    );
    let data = fs.readFileSync(targetFile, "utf-8");
    data = data.replace(/<webhook-url>/g, "https://test.com");
    fs.writeFileSync(targetFile, data);
    console.log("replace webhook url finish!");
  }

  public override async onBeforeBrowerStart(): Promise<void> {
    console.log("no need to verify in browser");
    return;
  }
}

new IncomingWebhookTestCase(
  TemplateProject.IncomingWebhook,
  14524902,
  undefined,
  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.StartWebhook],
  { skipInit: true, skipRemote: true }
).test();
