// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */
import { validateBot } from "../../utils/playwrightOperation";
import {
  Timeout,
  LocalDebugTaskLabel,
  LocalDebugTaskInfo,
  ValidationContent,
  Lang,
} from "../../utils/constants";
import { it } from "../../utils/it";
import {
  botHappyPathTestForLocalDebug,
  botHappyPathTestForRemoteDebug,
} from "./BotHappyPath";
import { Page } from "playwright";

describe("Local Debug Tests", function () {
  this.timeout(Timeout.localAndRemoteTestCase);
  const successFlag = {
    successFlagForLocal: false,
    successFlagForRemote: false,
  };

  async function validationCrbot(
    page: Page,
    options?: {
      botCommand?: string;
      expected?: ValidationContent;
    }
  ) {
    await validateBot(page, {
      botCommand: options?.botCommand,
      expected: options?.expected,
    });
  }

  after(async function () {
    this.timeout(Timeout.finishTestCase);
    setTimeout(() => {
      if (successFlag.successFlagForLocal && successFlag.successFlagForRemote)
        process.exit(0);
      else process.exit(1);
    }, 30000);
  });

  it(
    "[auto] Local debug Command and Response Bot App",
    {
      testPlanCaseId: 13999814,
      author: "aochengwang@microsoft.com",
    },
    async function () {
      await botHappyPathTestForLocalDebug("crbot", {
        lang: Lang.TS,
        successFlag: successFlag,
        localDebugTaskLabel: LocalDebugTaskLabel.StartBotApp,
        localDebugTaskInfo: LocalDebugTaskInfo.StartBotInfo,
        validationFn: validationCrbot,
        validationBotCommand: "helloWorld",
        validationExpectedResponse: "Your Hello World App is Running",
      });
    }
  );

  it(
    "[auto] Remote debug for command and response project Tests",
    {
      testPlanCaseId: 14112765,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      await botHappyPathTestForRemoteDebug("crbot", {
        lang: Lang.TS,
        successFlag: successFlag,
        validationFn: validationCrbot,
        validationBotCommand: "helloWorld",
        validationExpectedResponse: "Your Hello World App is Running",
      });
    }
  );
});
