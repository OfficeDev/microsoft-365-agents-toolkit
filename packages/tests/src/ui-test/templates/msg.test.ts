// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Anne Fu <v-annefu@microsoft.com>
 */
import { validateCreatedCard } from "../../utils/playwrightOperation";
import {
  Timeout,
  LocalDebugTaskLabel,
  LocalDebugTaskInfo,
  Lang,
} from "../../utils/constants";
import { it } from "../../utils/it";
import {
  msgHappyPathTestForLocalDebug,
  msgHappyPathTestForRemoteDebug,
} from "./MsgHappyPath";
import { Page } from "playwright";

describe("Debug Tests", function () {
  this.timeout(Timeout.localAndRemoteTestCase);
  const successFlag = {
    successFlagForLocal: false,
    successFlagForRemote: false,
  };
  async function validationCreatedCard(
    page: Page,
    options: {
      appName: string;
    }
  ) {
    await validateCreatedCard(page, options.appName);
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
    "[Javascript] Local Debug for Message Extension project",
    {
      testPlanCaseId: 9729550,
      author: "v-annefu@microsoft.com",
    },
    async function () {
      await msgHappyPathTestForLocalDebug("msg", {
        lang: Lang.JS,
        successFlag: successFlag,
        localDebugTaskLabel: LocalDebugTaskLabel.StartApplication,
        localDebugTaskInfo: LocalDebugTaskInfo.StartBotInfo2,
        validationFn: validationCreatedCard,
      });
    }
  );

  it(
    "[auto] Remote debug for Message Extension project Tests",
    {
      testPlanCaseId: 24739650,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      await msgHappyPathTestForRemoteDebug("msg", {
        lang: Lang.JS,
        successFlag: successFlag,
        validationFn: validationCreatedCard,
      });
    }
  );
});
