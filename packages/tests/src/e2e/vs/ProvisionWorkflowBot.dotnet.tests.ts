// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author dol <dol@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import { Runtime } from "../../commonlib/constants";
import { happyPathTest } from "../bot/WorkflowBotHappyPathCommon";

describe("Provision workflow Dotnet", () => {
  it(
    "Provision Resource: Workflow Dotnet",
    { testPlanCaseId: 24692255, author: "dol@microsoft.com" },
    async function () {
      await happyPathTest(Runtime.Dotnet);
    }
  );
});
