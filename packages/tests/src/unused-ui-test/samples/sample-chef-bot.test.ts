// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ivan Chen <v-ivanchen@microsoft.com>
 */

import { Page } from "playwright";
import {
  TemplateProject,
  LocalDebugTaskLabel,
  ValidationContent,
} from "../../utils/constants";
import { CaseFactory } from "../../ui-test/samples/sampleCaseFactory";
import { SampledebugContext } from "../../ui-test/samples/sampledebugContext";
import { validateWelcomeAndReplyBot } from "../../utils/playwrightOperation";
import * as path from "path";
import * as fs from "fs";
import * as os from "os";

class ChefBotTestCase extends CaseFactory {
  public override async onAfterCreate(
    sampledebugContext: SampledebugContext,
    env: "local" | "dev"
  ): Promise<void> {
    fs.mkdirSync(path.resolve(sampledebugContext.projectPath, "env"), {
      recursive: true,
    });

    const envFile = path.resolve(
      sampledebugContext.projectPath,
      "env",
      `.env.${env}.user`
    );
    // create .env.local.user file
    fs.writeFileSync(envFile, "SECRET_OPENAI_KEY=yourapikey");
    console.log(`add SECRET_OPENAI_KEY=yourapikey to .env file`);
    // await sampledebugContext.prepareDebug("yarn");
  }
  override async onValidate(page: Page): Promise<void> {
    console.log("Moked api key. Only verify happy path...");
    return await validateWelcomeAndReplyBot(page, {
      hasCommandReplyValidation: true,
      botCommand: "helloWorld",
      expectedReplyMessage: ValidationContent.AiBotErrorMessage3,
    });
  }
  public override async onCliValidate(page: Page): Promise<void> {
    console.log("Mocked api key. Only verify happy path...");
    return await validateWelcomeAndReplyBot(page, {
      hasCommandReplyValidation: true,
      botCommand: "helloWorld",
      expectedReplyMessage: ValidationContent.AiBotErrorMessage3,
    });
  }
}

new ChefBotTestCase(
  TemplateProject.ChefBot,
  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.StartLocalTunnel, LocalDebugTaskLabel.StartBotApp],
  {
    repoPath: "./resource/js/samples/04.ai-apps",
    testRootFolder: path.resolve(os.homedir(), "resource"),
    testPlanCaseId_local: 24409837,
    testPlanCaseId_dev: 24409842,
  }
).test();
