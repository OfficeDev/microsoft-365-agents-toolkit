// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yukun-dong <yukundong@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import { Runtime } from "../../commonlib/constants";
import { happyPathTest } from "../bot/BotHappyPathCommon";

describe("Remote happy path for message extension search dotnet", () => {
  it(
    "Remote happy path for message extension search dotnet",
    { testPlanCaseId: 24916471, author: "yukundong@microsoft.com" },
    async function () {
      await happyPathTest(Runtime.Dotnet, "search-message-extension");
    }
  );
});
