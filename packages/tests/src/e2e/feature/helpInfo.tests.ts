// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Zhiyu You <zhiyou@microsoft.com>
 */

import { it } from "@microsoft/extra-shot-mocha";
import { expect } from "chai";
import { execAsync } from "../commonUtils";

describe("atk command help", function () {
  it(
    `atk auth list -h`,
    { testPlanCaseId: 15685961, author: "zhiyou@microsoft.com" },
    async function () {
      const result = await execAsync(`atk auth list -h`, {
        env: process.env,
        timeout: 0,
      });
      expect(result.stdout).not.includes("--action");
    }
  );
});
