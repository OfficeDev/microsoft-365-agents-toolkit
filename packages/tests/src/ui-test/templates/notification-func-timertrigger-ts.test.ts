// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Aocheng Wang <aochengwang@microsoft.com>
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

describe("Func Hosted and Timer-trigger Notification Bot App Tests", function () {
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
      await notiBotHappyPathTestForLocalDebug("ftNoti", {
        lang: "typescript",
        successFlag: successFlag,
        fileValidation: "src/httpTrigger.ts",
        localDebugTaskLabel: LocalDebugTaskLabel.StartBotApp,
        localDebugTaskInfo: LocalDebugTaskInfo.BackendStartedInfo,
        haveTaskCompileTypescript: true,
        validationFn: validationNotiBot,
        httpTrigger: true,
        timerTrigger: true,
      });
    }
  );

  it(
    "[auto] [Typescript] Remote debug for Func Hosted and Timer-trigger Notification Bot App",
    {
      testPlanCaseId: 17431811,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      await notiBotHappyPathTestForRemoteDebug("functimernoti", {
        lang: "TypeScript",
        successFlag: successFlag,
        fileValidation: "src/httpTrigger.ts",
        validationFn: validationNotiBot,
        httpTrigger: true,
        timerTrigger: true,
      });
    }
  );
});
