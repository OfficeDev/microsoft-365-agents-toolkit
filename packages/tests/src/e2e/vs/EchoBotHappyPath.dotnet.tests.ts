// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yukun-dong <yukundong@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import { Runtime } from "../../commonlib/constants";
import { happyPathTest } from "../bot/BotHappyPathCommon";

describe("Remote happy path for echo bot dotnet", () => {
  it(
    "Remote happy path for echo bot dotnet",
    { testPlanCaseId: 24916323, author: "yukundong@microsoft.com" },
    async function () {
      await happyPathTest(Runtime.Dotnet, "bot");
    }
  );
});
