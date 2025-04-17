// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Anne Fu <v-annefu@microsoft.com>
 */
import {
  validateNotificationBot,
  validateNotificationTimeBot,
} from "../../utils/playwrightOperation";
import {
  Timeout,
  LocalDebugTaskLabel,
  LocalDebugTaskInfo,
} from "../../utils/constants";
import { it } from "../../utils/it";
import { Page } from "playwright";
import {
  notiBotHappyPathTestForLocalDebug,
  notiBotHappyPathTestForRemoteDebug,
} from "./NotiBotHappyPath";

describe("Time-trigger Notification Bot Tests", function () {
  this.timeout(Timeout.localAndRemoteTestCase);
  const successFlag = {
    successFlagForLocal: false,
    successFlagForRemote: false,
  };

  async function validationNotiBot(
    page: Page,
    httpTrigger?: boolean,
    timerTrigger?: boolean,
    funcEndpoint?: string
  ) {
    if (httpTrigger) {
      await validateNotificationBot(page, funcEndpoint);
    }
    if (timerTrigger) {
      await validateNotificationTimeBot(page);
    }
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
    "[auto] [Typescript] Local debug Func Hosted and Timer-trigger Notification Bot App",
    {
      testPlanCaseId: 17431807,
      author: "v-annefu@microsoft.com",
    },
    async function () {
      await notiBotHappyPathTestForLocalDebug("timeNoti", {
        lang: "typescript",
        successFlag: successFlag,
        fileValidation: "src/timerTrigger.ts",
        localDebugTaskLabel: LocalDebugTaskLabel.StartBotApp,
        localDebugTaskInfo: LocalDebugTaskInfo.BackendStartedInfo,
        haveTaskCompileTypescript: true,
        validationFn: validationNotiBot,
        httpTrigger: false,
        timerTrigger: true,
      });
    }
  );

  it(
    "[auto] [Typescript] Remote debug Time-trigger Notification Bot App",
    {
      testPlanCaseId: 16551762,
      author: "v-annefu@microsoft.com",
    },
    async function () {
      await notiBotHappyPathTestForRemoteDebug("timenoti", {
        lang: "TypeScript",
        successFlag: successFlag,
        fileValidation: "src/timerTrigger.ts",
        validationFn: validationNotiBot,
        httpTrigger: false,
        timerTrigger: true,
      });
    }
  );
});
