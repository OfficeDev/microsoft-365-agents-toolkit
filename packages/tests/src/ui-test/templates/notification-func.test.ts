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

describe("Func Hosted Notification Bot Local Debug Tests", function () {
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
    "[auto] Local debug Func Hosted Notification Bot App",
    {
      testPlanCaseId: 13999812,
      author: "aochengwang@microsoft.com",
    },
    async function () {
      await notiBotHappyPathTestForLocalDebug("funcNoti", {
        lang: "javascript",
        successFlag: successFlag,
        fileValidation: "src/httpTrigger.js",
        localDebugTaskLabel: LocalDebugTaskLabel.StartBotApp,
        localDebugTaskInfo: LocalDebugTaskInfo.BackendStartedInfo,
        validationFn: validationNotiBot,
        httpTrigger: true,
        timerTrigger: false,
      });
    }
  );

  it(
    "[auto] Remote debug for func hosted notification project Tests",
    {
      testPlanCaseId: 14112811,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      await notiBotHappyPathTestForRemoteDebug("funcnoti", {
        lang: "JavaScript",
        successFlag: successFlag,
        validationFn: validationNotiBot,
        fileValidation: "src/httpTrigger.js",
        httpTrigger: true,
        timerTrigger: false,
      });
    }
  );
});
