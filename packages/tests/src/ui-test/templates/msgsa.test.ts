// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Anne Fu <v-annefu@microsoft.com>
 */
import { validateNpm } from "../../utils/playwrightOperation";
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
  async function validationNpm(
    page: Page,
    options: {
      appName: string;
    }
  ) {
    await validateNpm(page, {
      npmName: "axios",
      appName: options.appName,
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
    "[Javascript] Local debug for Search-based message extension project",
    {
      testPlanCaseId: 14483079,
      author: "xiaofu.huang@microsoft.com",
    },
    async function () {
      await msgHappyPathTestForLocalDebug("msgsa", {
        lang: Lang.JS,
        successFlag: successFlag,
        localDebugTaskLabel: LocalDebugTaskLabel.StartApplication,
        localDebugTaskInfo: LocalDebugTaskInfo.StartBotInfo2,
        validationFn: validationNpm,
      });
    }
  );

  it(
    "[auto] Remote debug for Search-based message extension project Tests",
    {
      testPlanCaseId: 14907803,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      await msgHappyPathTestForRemoteDebug("msgsa", {
        lang: Lang.JS,
        successFlag: successFlag,
        validationFn: validationNpm,
      });
    }
  );
});
