// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Yuqi Zhou <yuqzho@microsoft.com>
 **/
import { happyPathTest } from "../BotHappyPathCommon";
import { Runtime } from "../../../commonlib/constants";
import { it } from "@microsoft/extra-shot-mocha";

describe("Provision message extension used for TDP integration", () => {
  it(
    "Provision Resource: message extension for TDP integration",
    { testPlanCaseId: 26547090, author: "yuqzho@microsoft.com" },
    async function () {
      await happyPathTest(
        Runtime.Node,
        "collect-form-message-extension",
        undefined,
        true
      );
    }
  );
});
