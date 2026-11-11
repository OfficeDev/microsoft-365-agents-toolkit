// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect, vi } from "vitest";
import { FeatureFlagName } from "../../../../src/common/featureFlags";
import {
  mapPersonalMicrosoftAccountsOption,
  validateSupportedAccountTypesCloud,
} from "../../../../src/component/driver/oauth/utility/supportedAccountTypes";

describe("personal Microsoft accounts option", () => {
  afterEach(() => {
    vi.unstubAllEnvs();
  });

  for (const testCase of [
    { input: undefined, expected: undefined },
    { input: false, expected: "Enterprise" },
    { input: true, expected: "Enterprise, Consumer" },
  ]) {
    it(`OAUTH-MSA-02: maps ${String(testCase.input)} to the TGS contract`, () => {
      const result = mapPersonalMicrosoftAccountsOption(testCase.input, "oauth/register");

      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value).toBe(testCase.expected);
      }
    });
  }

  for (const input of ["true", 1, null, {}]) {
    it(`OAUTH-MSA-03: rejects ${JSON.stringify(input)}`, () => {
      const result = mapPersonalMicrosoftAccountsOption(input, "oauth/register");

      expect(result.isErr()).toBe(true);
      if (result.isErr()) {
        expect(result.error.name).toBe("InvalidActionInputError");
      }
    });
  }

  it("OAUTH-MSA-04: rejects the parameter outside Public cloud", () => {
    vi.stubEnv(FeatureFlagName.SovereignCloudEnvironment, "GCC H");

    const result = validateSupportedAccountTypesCloud("Enterprise, Consumer", "oauth/register");

    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.name).toBe("OauthSupportedAccountTypesUnsupportedCloud");
    }
  });
});
