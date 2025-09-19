// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Siglud <fanhu@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import { Runtime } from "../../commonlib/constants";
import { happyPathTest } from "../bot/CommandBotHappyPathCommon";

describe("Provision CommandBot Dotnet", () => {
  it(
    "Provision Resource: CommandBot Dotnet",
    { testPlanCaseId: 15685857, author: "fanhu@microsoft.com" },
    async function () {
      await happyPathTest(Runtime.Dotnet);
    }
  );
});
