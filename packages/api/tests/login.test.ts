// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
"use strict";

import { assert } from "chai";
import "mocha";
import { Result, ok } from "neverthrow";
import { FxError } from "../src/error";
import { BasicLogin, LoginStatus, M365TokenProvider, TokenRequest } from "../src/utils/login";

class M365Provider extends BasicLogin implements M365TokenProvider {
  async getAccessToken(tokenRequest: TokenRequest): Promise<Result<string, FxError>> {
    return ok("fakeToken");
  }
  async getJsonObject(
    tokenRequest: TokenRequest
  ): Promise<Result<Record<string, unknown>, FxError>> {
    return ok({});
  }
  async getStatus(tokenRequest: TokenRequest): Promise<Result<LoginStatus, FxError>> {
    return ok({ status: "SignedIn" });
  }

  async switchTenant(tenantId: string): Promise<Result<string, FxError>> {
    return ok("fakeToken");
  }

  async signout(): Promise<boolean> {
    return true;
  }
}

describe("m365Login", function () {
  describe("withM365Provider", function () {
    it("happy path", async () => {
      const m365 = new M365Provider();
      assert.exists(await m365.getAccessToken({ scopes: ["test"] }));
      assert.exists(await m365.getJsonObject({ scopes: ["test"] }));
      assert.exists(await m365.getStatus({ scopes: ["test"] }));
      assert.exists(
        await m365.setStatusChangeMap(
          "a",
          { scopes: ["test"] },
          async (status, token, accountInfo) => {
            console.log(status);
          },
          true
        )
      );
      assert.exists(await m365.removeStatusChangeMap("a"));
    });
  });
});
