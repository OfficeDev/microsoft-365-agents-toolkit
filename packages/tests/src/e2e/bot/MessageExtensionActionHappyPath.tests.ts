// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Siglud <fanhu@microsoft.com>
 **/
import { it } from "@microsoft/extra-shot-mocha";
import { Runtime } from "../../commonlib/constants";
import { happyPathTest } from "./BotHappyPathCommon";

describe("Provision message extension Node template", () => {
  it(
    "Provision Template: message extension for NodeJS",
    { testPlanCaseId: 15685647, author: "fanhu@microsoft.com" },
    async function () {
      await happyPathTest(Runtime.Node, "basic-message-extension");
    }
  );
});
