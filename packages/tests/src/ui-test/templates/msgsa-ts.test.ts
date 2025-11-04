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
    "[Typescript] Local debug for Search-based message extension project",
    {
      testPlanCaseId: 15277314,
      author: "xiaofu.huang@microsoft.com",
    },
    async function () {
      await msgHappyPathTestForLocalDebug("msgsa", {
        lang: Lang.TS,
        successFlag: successFlag,
        localDebugTaskLabel: LocalDebugTaskLabel.StartApplication,
        localDebugTaskInfo: LocalDebugTaskInfo.ListeningOn,
        validationFn: validationNpm,
      });
    }
  );

  it(
    "[auto] [TypeScript] Remote debug for Search-based message extension typescript project Tests",
    {
      testPlanCaseId: 14907800,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      await msgHappyPathTestForRemoteDebug("msgsa", {
        lang: Lang.TS,
        successFlag: successFlag,
        validationFn: validationNpm,
      });
    }
  );
});
