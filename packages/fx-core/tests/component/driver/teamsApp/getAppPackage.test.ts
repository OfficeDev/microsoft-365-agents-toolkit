// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { M365TokenProvider, ok } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import mockedEnv from "mocked-env";
import { afterEach, chai, describe, it, vi } from "vitest";
import { teamsDevPortalClient } from "../../../../src/client/teamsDevPortalClientProvider";
import { FeatureFlagName } from "../../../../src/common/featureFlags";
import { getAppPackage } from "../../../../src/component/driver/teamsApp/appStudio";

describe("getAppPackage", () => {
  let restoreEnv: (() => void) | undefined;

  afterEach(() => {
    vi.restoreAllMocks();
    restoreEnv?.();
  });

  it("reads a non-Buffer package returned by the new Developer Portal API", async () => {
    restoreEnv = mockedEnv({ [FeatureFlagName.NewDeveloperPortalApis]: "true" });
    const zip = new AdmZip();
    zip.addFile("manifest.json", Buffer.from("{}"));
    const packageBytes = new Uint8Array(zip.toBuffer());
    vi.spyOn(teamsDevPortalClient, "getAppPackage").mockResolvedValue(packageBytes);
    const tokenProvider = {
      getAccessToken: vi.fn().mockResolvedValue(ok("token")),
    } as unknown as M365TokenProvider;

    const result = await getAppPackage("app-id", tokenProvider);

    chai.assert.isTrue(result.isOk());
    if (result.isOk()) {
      chai.assert.equal(result.value.manifest?.toString(), "{}");
    }
  });
});
