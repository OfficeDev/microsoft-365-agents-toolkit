// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Helly Zhang <v-helzha@microsoft.com>
 */
import { validatePrompt } from "../../utils/playwrightOperation";
import { Timeout, Lang } from "../../utils/constants";
import { it } from "../../utils/it";
import { Page } from "playwright";
import {
  daActionHappPathTestForLocalDebug,
  daActionHappPathTestForRemoteDebug,
} from "./DaActionHappPath";

describe("Debug Tests", function () {
  this.timeout(Timeout.localAndRemoteTestCase);
  const successFlag = {
    successFlagForLocal: false,
    successFlagForRemote: false,
  };

  async function validateFn(
    page: Page,
    options: {
      appName: string;
      prompt?: string;
      expected?: string;
    }
  ) {
    await validatePrompt(page, options.appName, {
      prompt: options?.prompt || "Show repair records assigned to Karin Blair",
      expected: options?.expected || "Oil",
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
    "[auto][VSC][JS] Local debug for basic declarative agent with API plugin from scratch (API key)",
    {
      testPlanCaseId: 34628854,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      await daActionHappPathTestForLocalDebug("daaction", {
        lang: Lang.JS,
        apiAuth: "api-key",
        successFlag: successFlag,
        fileValidation: "src/functions/repair.js",
        validationFn: validateFn,
      });
    }
  );
  it(
    "[auto][VSC][JS] Remote debug for basic declarative agent with API plugin from scratch (API key)",
    {
      testPlanCaseId: 34628865,
      author: "v-helzha@microsoft.com",
    },
    async function () {
      await daActionHappPathTestForRemoteDebug("daAction", {
        lang: Lang.JS,
        authOption: "API Key",
        successFlag: successFlag,
        fileValidation: "src/functions/repair.js",
        validationFn: validateFn,
      });
    }
  );
});
