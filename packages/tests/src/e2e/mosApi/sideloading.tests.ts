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
  it(
    "install & uninstall for non DA app",
    { testPlanCaseId: 31339285, author: "huajiezhang@microsoft.com" },
    async function () {
      const m365TitleHelper = await M365TitleHelper.init(
        "https://titles.prod.mos.microsoft.com",
        "https://titles.prod.mos.microsoft.com/.default"
      );
      const packageFilePath = path.join(__dirname, ".", "appPackage.local.zip");
      let success = false;
      try {
        const res = await m365TitleHelper.acquire(packageFilePath);
        assert.isDefined(res[0], res[1]);
        await m365TitleHelper.unacquire(res[0]);
        success = true;
        console.log(
          "Successfully call MOS3 API (acquire/unacquire) for non DA app"
        );
      } catch (e) {
        console.error(
          "Failed to call MOS3 API (acquire/unacquire) for non DA app"
        );
        console.error(JSON.stringify(e, Object.getOwnPropertyNames(e)));
      }
      assert.isTrue(success);
    }
  );

  it(
    "install & uninstall for DA app",
    { testPlanCaseId: 31339285, author: "huajiezhang@microsoft.com" },
    async function () {
      const m365TitleHelper = await M365TitleHelper.init(
        "https://titles.prod.mos.microsoft.com",
        "https://titles.prod.mos.microsoft.com/.default"
      );
      const packageFilePath = path.join(
        __dirname,
        ".",
        "appPackage-da.local.zip"
      );
      let success = false;
      try {
        const res = await m365TitleHelper.acquire(packageFilePath);
        assert.isDefined(res[0], res[1]);
        await m365TitleHelper.unacquire(res[0]);
        success = true;
        console.log(
          "Successfully call MOS3 API (acquire/unacquire) for DA app"
        );
      } catch (e) {
        console.error("Failed to call MOS3 API (acquire/unacquire) for DA app");
        console.error(JSON.stringify(e, Object.getOwnPropertyNames(e)));
      }
      assert.isTrue(success);
    }
  );
});
