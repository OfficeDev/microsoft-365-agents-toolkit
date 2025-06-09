// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { assert } from "chai";
import { getTemplateZipUrlByVersion, getTemplateUrl } from "../../../src/component/generator/utils";
import { Platform } from "@microsoft/teamsfx-api";
import mockedEnv from "mocked-env";
describe("utils unit test cases", () => {
  it("should return the correct URL for a given version", () => {
    const version = "6.0.0";
    const perfix = "templates@";
    const name = "js";
    const expectedUrl =
      "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download/templates@6.0.0/js.zip";
    const result = getTemplateZipUrlByVersion(name, version, perfix);
    assert.strictEqual(result, expectedUrl);
  });

  it("should return the correct URL for getTemplateVSUrl", async () => {
    const restore = mockedEnv({
      TEAMSFX_TEMPLATE_PRERELEASE: "rc",
    });
    const getLatestVersion = () => Promise.resolve("0.0.0-rc");
    const result = await getTemplateUrl("csharp", getLatestVersion, Platform.VS);
    const expectedUrl =
      "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download/templates-vs@0.0.0-rc/csharp.zip";
    assert.strictEqual(result, expectedUrl);
    restore();
  });

  it("should return the correct URL for getTemplateVSCUrl", async () => {
    const restore = mockedEnv({
      TEAMSFX_TEMPLATE_PRERELEASE: "rc",
    });
    const getLatestVersion = () => Promise.resolve("0.0.0-rc");
    const result = await getTemplateUrl("ts", getLatestVersion, Platform.VSCode);
    const expectedUrl =
      "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download/templates@0.0.0-rc/ts.zip";
    assert.strictEqual(result, expectedUrl);
    restore();
  });
});
