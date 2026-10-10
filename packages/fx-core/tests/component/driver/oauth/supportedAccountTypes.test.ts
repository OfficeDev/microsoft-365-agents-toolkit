// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect, vi } from "vitest";
import { FeatureFlagName } from "../../../../src/common/featureFlags";
import {
  normalizeSupportedAccountTypes,
  validateSupportedAccountTypesCloud,
} from "../../../../src/component/driver/oauth/utility/supportedAccountTypes";

describe("supportedAccountTypes", () => {
  afterEach(() => {
    vi.unstubAllEnvs();
  });

  for (const testCase of [
    { input: undefined, expected: undefined },
    { input: "Enterprise", expected: "Enterprise" },
    { input: "Enterprise, Consumer", expected: "Enterprise, Consumer" },
    { input: "Consumer, Enterprise", expected: "Enterprise, Consumer" },
  ]) {
    it(`OAUTH-MSA-02: normalizes ${String(testCase.input)}`, () => {
      const result = normalizeSupportedAccountTypes(testCase.input, "oauth/register");

      expect(result.isOk()).toBe(true);
      if (result.isOk()) {
        expect(result.value).toBe(testCase.expected);
      }
    });
  }

  for (const input of [
    "Consumer",
    "",
    "Enterprise, Enterprise",
    "Enterprise, Unknown",
    "Enterprise,, Consumer",
    1,
  ]) {
    it(`OAUTH-MSA-03: rejects ${JSON.stringify(input)}`, () => {
      const result = normalizeSupportedAccountTypes(input, "oauth/register");

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
