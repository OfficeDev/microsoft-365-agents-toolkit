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
    const m365TitleHelper = await M365TitleHelper.init(
      "https://titles.prod.mos.microsoft.com",
      "https://titles.prod.mos.microsoft.com/.default"
    );
    const packageFilePath = path.join(__dirname, ".", "appPackage.local.zip");
    const errors: any[] = [];
    let success = false;
    for (let i = 0; i < 5; i++) {
      try {
        if (i > 0) {
          console.log(`Retry ${i} time`);
        }
        const res = await m365TitleHelper.acquire(packageFilePath);
        assert.isDefined(res[0], res[1]);
        await m365TitleHelper.unacquire(res[0]);
        success = true;
        break;
      } catch (e) {
        console.error(`Failed to call MOS3 API (acquire/unacquire): ${e}`);
        errors.push(e);
        await new Promise((resolve) => setTimeout(resolve, 5000));
      }
    }
    if (success) {
      console.log("Successfully call MOS3 API (acquire/unacquire)");
    } else {
      console.error(
        "Failed to call MOS3 API (acquire/unacquire) after 5 retries"
      );
      console.error(errors);
    }
    assert.isTrue(success);
  });
});
