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
import { CaseFactory } from "./sampleCaseFactory";
import { SampledebugContext } from "./sampledebugContext";
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
      ".env.local"
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
      expectedReplyMessage: ValidationContent.AiBotErrorMessage,
    });
  }
  public override async onCliValidate(page: Page): Promise<void> {
    console.log("Mocked api key. Only verify happy path...");
    return await validateWelcomeAndReplyBot(page, {
      hasCommandReplyValidation: true,
      botCommand: "helloWorld",
      expectedReplyMessage: ValidationContent.AiBotErrorMessage,
    });
  }
}

new ChefBotTestCase(
  TemplateProject.ChefBot,
  24409837,
  24409842,
  "v-ivanchen@microsoft.com",
  [LocalDebugTaskLabel.StartLocalTunnel, LocalDebugTaskLabel.StartBotApp],
  {
    repoPath: "./resource/js/samples/04.ai-apps",
    testRootFolder: path.resolve(os.homedir(), "resource"),
  }
).test();
