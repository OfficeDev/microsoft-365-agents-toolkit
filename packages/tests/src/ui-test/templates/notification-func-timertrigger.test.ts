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
    "[auto] [Javascript] Local debug Func Hosted and Timer-trigger Notification Bot App",
    {
      testPlanCaseId: 17431806,
      author: "v-annefu@microsoft.com",
    },
    async function () {
      await notiBotHappyPathTestForLocalDebug("ftNoti", {
        lang: "javascript",
        successFlag: successFlag,
        fileValidation: "src/httpTrigger.js",
        localDebugTaskLabel: LocalDebugTaskLabel.StartBotApp,
        localDebugTaskInfo: LocalDebugTaskInfo.BackendStartedInfo,
        validationFn: validationNotiBot,
        httpTrigger: true,
        timerTrigger: true,
      });
    }
  );

  it(
    "[auto] [Javascript] Remote debug for Func Hosted and Timer-trigger Notification Bot App",
    {
      testPlanCaseId: 17431813,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      await notiBotHappyPathTestForRemoteDebug("functimernoti", {
        lang: "JavaScript",
        successFlag: successFlag,
        validationFn: validationNotiBot,
        fileValidation: "src/httpTrigger.js",
        httpTrigger: true,
        timerTrigger: true,
      });
    }
  );
});
