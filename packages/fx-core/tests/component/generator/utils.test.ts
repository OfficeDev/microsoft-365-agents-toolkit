// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { assert } from "chai";
import { getTemplateZipUrlByVersion, getTemplateUrl } from "../../../src/component/generator/utils";
import { Platform } from "@microsoft/teamsfx-api";
import mockedEnv from "mocked-env";
import proxyquire from "proxyquire";
import sinon from "sinon";
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

  it("should return undefined for alpha getTemplateVSUrl", async () => {
    const getLatestVersion = () => Promise.resolve("0.0.0-rc");
    const result = await getTemplateUrl("csharp", getLatestVersion, Platform.VS);
    assert.isUndefined(result);
  });

  it("should return the correct URL for RC getTemplateVSUrl", async () => {
    const restore = mockedEnv({
      TEAMSFX_TEMPLATE_PRERELEASE: "vs",
    });
    const mockSettings = {
      version: "~6.0",
      localVersion: "6.0.0",
      tagPrefix: "templates@",
      vstagPrefix: "templates-vs@",
      vsversion: "17.0.0",
      tagListURL:
        "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download/template-tag-list/template-tags.txt",
      templateDownloadBaseURL:
        "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download",
      templateReleaseURL:
        "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/expanded_assets",
      templateDownloadBasePath: "/OfficeDev/microsoft-365-agents-toolkit/releases/download",
      templateExt: ".zip",
      useLocalTemplate: false,
    };
    const dUtils = proxyquire("../../../src/component/generator/utils", {
      "../../common/templates-config.json": mockSettings,
    });
    const getLatestVersion = () => Promise.resolve("0.0.0-rc");
    const result = await dUtils.getTemplateUrl("csharp", getLatestVersion, Platform.VS);
    const expectedUrl =
      "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download/templates-vs@0.0.0-rc/csharp.zip";
    assert.strictEqual(result, expectedUrl);
    restore();
  });

  it("should return the stable URL for getTemplateVSUrl", async () => {
    const mockSettings = {
      version: "~6.0",
      localVersion: "6.0.0",
      tagPrefix: "templates@",
      vstagPrefix: "templates-vs@",
      vsversion: "17.0.0",
      tagListURL:
        "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download/template-tag-list/template-tags.txt",
      templateDownloadBaseURL:
        "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download",
      templateReleaseURL:
        "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/expanded_assets",
      templateDownloadBasePath: "/OfficeDev/microsoft-365-agents-toolkit/releases/download",
      templateExt: ".zip",
      useLocalTemplate: false,
    };
    const dUtils = proxyquire("../../../src/component/generator/utils", {
      "../../common/templates-config.json": mockSettings,
    });
    const getLatestVersion = () => Promise.resolve("17.0.0");
    const result = await dUtils.getTemplateUrl("csharp", getLatestVersion, Platform.VS);
    const expectedUrl =
      "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download/templates-vs@17.0.0/csharp.zip";
    assert.strictEqual(result, expectedUrl);
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

  it("should return undefined for use local template for getTemplateVSCUrl", async () => {
    const restore = mockedEnv({
      TEAMSFX_TEMPLATE_PRERELEASE: "",
    });
    const getLatestVersion = () => Promise.resolve("0.0.0-rc");
    const result = await getTemplateUrl("ts", getLatestVersion, Platform.VSCode);
    assert.isUndefined(result);
    restore();
  });
});
