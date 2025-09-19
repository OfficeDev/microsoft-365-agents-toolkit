// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Siglud <fanhu@microsoft.com>
 **/
import { it } from "@microsoft/extra-shot-mocha";
import { Runtime } from "../../commonlib/constants";
import { happyPathTest } from "../bot/BotHappyPathCommon";

describe("Provision message extension action Dotnet", () => {
  it(
    "Provision Resource: message extension action dotnet",
    { testPlanCaseId: 15685646, author: "fanhu@microsoft.com" },
    async function () {
      await happyPathTest(Runtime.Dotnet, "collect-form-message-extension");
    }
  );
});
