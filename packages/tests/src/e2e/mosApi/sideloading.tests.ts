// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Huajie Zhang <huajiezhang@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import path from "path";
import { M365TitleHelper } from "../../commonlib/m365TitleHelper";
import { assert } from "chai";

describe("MOS3 API", function () {
  it(`sideloading`, async function () {
    // }, //   author: "huajiezhang@microsoft.com", //   testPlanCaseId:9426074, // {
    const m365TitleHelper = await M365TitleHelper.init(
      "https://titles.prod.mos.microsoft.com",
      "https://titles.prod.mos.microsoft.com/.default"
    );
    const packageFilePath = path.join(__dirname, ".", "appPackage.local.zip");
    const res = await m365TitleHelper.acquire(packageFilePath);
    assert.isDefined(res[0], res[1]);
    await m365TitleHelper.unacquire(res[0]);
  });
});
