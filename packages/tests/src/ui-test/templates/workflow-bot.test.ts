// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */
import {
  validateBot,
  validateWorkFlowBot,
} from "../../utils/playwrightOperation";
import {
  Timeout,
  LocalDebugTaskLabel,
  LocalDebugTaskInfo,
  ValidationContent,
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

  async function validationWorkflowBot(
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
    await validateWorkFlowBot(page);
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
    "[auto] [JavaScript] Local debug workflow app",
    {
      testPlanCaseId: 15638255,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      await botHappyPathTestForLocalDebug("workflow", {
        lang: "javascript",
        successFlag: successFlag,
        localDebugTaskLabel: LocalDebugTaskLabel.StartBotApp,
        localDebugTaskInfo: LocalDebugTaskInfo.StartBotInfo,
        validationFn: validationWorkflowBot,
        validationBotCommand: "helloWorld",
        validationExpectedResponse: "Your Hello World Bot is Running",
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
      await botHappyPathTestForRemoteDebug("workflow", {
        lang: "JavaScript",
        successFlag: successFlag,
        validationFn: validationWorkflowBot,
        validationBotCommand: "helloWorld",
        validationExpectedResponse: "Your Hello World Bot is Running",
      });
    }
  );
});
