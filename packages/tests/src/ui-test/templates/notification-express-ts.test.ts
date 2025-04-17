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

describe("Express Notification Bot Tests", function () {
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
    "[auto] Local debug Express Notification Bot App",
    {
      testPlanCaseId: 13999815,
      author: "aochengwang@microsoft.com",
    },
    async function () {
      await notiBotHappyPathTestForLocalDebug("expressnoti", {
        lang: "typescript",
        successFlag: successFlag,
        localDebugTaskLabel: LocalDebugTaskLabel.StartBotApp,
        localDebugTaskInfo: LocalDebugTaskInfo.StartBotInfo,
        validationFn: validationNotiBot,
        httpTrigger: true,
        timerTrigger: false,
      });
    }
  );

  it(
    "[auto] Remote debug for express notification project Tests",
    {
      testPlanCaseId: 14112917,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      await notiBotHappyPathTestForRemoteDebug("expressnoti", {
        lang: "TypeScript",
        successFlag: successFlag,
        validationFn: validationNotiBot,
        httpTrigger: true,
        timerTrigger: false,
      });
    }
  );
});
