// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { DeclarativeAgentManifest, Platform, err, ok, signedIn } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import mockedEnv from "mocked-env";
import proxyquire from "proxyquire";
import * as sinon from "sinon";
import { GraphClient } from "../../../src/client/graphClient";
import { createContext, setTools } from "../../../src/common/globalVars";
import * as requestUtils from "../../../src/common/requestUtils";
import { copilotGptManifestUtils } from "../../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import { useLocalTemplate } from "../../../src/component/generator/templateHelper";
import {
  getTemplateUrl,
  getTemplateVSLatestVersion,
  getTemplateZipUrlByVersion,
  setGeneralSensitivityLabel,
} from "../../../src/component/generator/utils";
import { MockTools } from "../../core/utils";
import templateConfig from "../../../src/common/templates-config.json";

describe("utils unit test cases", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

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
      vsversion: templateConfig.vsversion,
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
    // Stub HTTP so getTemplateVSLatestVersion() can resolve without network
    sandbox
      .stub(requestUtils, "sendRequestWithTimeout")
      .resolves({ data: "templates-vs@18.0.0\ntemplates@6.0.0\n" } as any);
    const mockSettings = {
      version: "~6.0",
      localVersion: "6.0.0",
      tagPrefix: "templates@",
      vstagPrefix: "templates-vs@",
      vsversion: "18.0.0",
      vsVersionPattern: "~18.0",
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
      "../../../package.json": { version: "3.0.0" }, // stable, not beta
    });
    const getLatestVersion = () => Promise.resolve("18.0.0");
    const result = await dUtils.getTemplateUrl("csharp", getLatestVersion, Platform.VS);
    const expectedUrl =
      "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download/templates-vs@18.0.0/csharp.zip";
    assert.strictEqual(result, expectedUrl);
  });

  it("should return rc URL for beta fx-core version (VS pre-release test build)", async () => {
    const mockPackageJson = { version: "3.0.0-beta.1" };
    const mockSettings = {
      version: "~6.0",
      localVersion: "6.0.0",
      tagPrefix: "templates@",
      vstagPrefix: "templates-vs@",
      vsversion: templateConfig.vsversion,
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
      "../../../package.json": mockPackageJson,
    });
    const getLatestVersion = () => Promise.resolve(templateConfig.vsversion);
    const result = await dUtils.getTemplateUrl("csharp", getLatestVersion, Platform.VS);
    const expectedUrl =
      "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download/templates-vs@0.0.0-rc/csharp.zip";
    assert.strictEqual(result, expectedUrl);
  });

  it("should return the correct URL for getTemplateVSCUrl", async () => {
    const restore = mockedEnv({
      TEMPLATE_VERSION: "0.0.0-rc",
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
      TEMPLATE_VERSION: "local",
    });
    const getLatestVersion = () => Promise.resolve("0.0.0-rc");
    const result = await getTemplateUrl("ts", getLatestVersion, Platform.VSCode);
    assert.isUndefined(result);
    restore();
  });

  it("should return undefined for alpha version in package.json", async () => {
    const mockPackageJson = {
      version: "3.0.0-alpha.1",
    };
    const dUtils = proxyquire("../../../src/component/generator/utils", {
      "../../../package.json": mockPackageJson,
    });
    const getLatestVersion = () => Promise.resolve("6.0.0");
    const result = await dUtils.getTemplateUrl("ts", getLatestVersion, Platform.VSCode);
    assert.isUndefined(result);
  });

  it("should return rc URL for rc version in package.json", async () => {
    const mockPackageJson = {
      version: "3.0.0-rc.1",
    };
    const dUtils = proxyquire("../../../src/component/generator/utils", {
      "../../../package.json": mockPackageJson,
    });
    const getLatestVersion = () => Promise.resolve("6.0.0");
    const result = await dUtils.getTemplateUrl("ts", getLatestVersion, Platform.VSCode);
    const expectedUrl =
      "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download/templates@0.0.0-rc/ts.zip";
    assert.strictEqual(result, expectedUrl);
  });

  it("should use latest version for beta version in package.json when latest is higher", async () => {
    const mockPackageJson = {
      version: "3.0.0-beta.1",
    };
    const mockSettings = {
      version: "~6.0",
      localVersion: "5.0.0",
      tagPrefix: "templates@",
      vstagPrefix: "templates-vs@",
      vsversion: "18.0.0",
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
      "../../../package.json": mockPackageJson,
      "../../common/templates-config.json": mockSettings,
    });
    const getLatestVersion = () => Promise.resolve("6.0.0");
    const result = await dUtils.getTemplateUrl("ts", getLatestVersion, Platform.VSCode);
    const expectedUrl =
      "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download/templates@6.0.0/ts.zip";
    assert.strictEqual(result, expectedUrl);
  });

  it("should return undefined for stable version when latest is not higher than local", async () => {
    const mockPackageJson = {
      version: "3.0.0",
    };
    const mockSettings = {
      version: "~6.0",
      localVersion: "6.0.0",
      tagPrefix: "templates@",
      vstagPrefix: "templates-vs@",
      vsversion: "18.0.0",
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
      "../../../package.json": mockPackageJson,
      "../../common/templates-config.json": mockSettings,
    });
    const getLatestVersion = () => Promise.resolve("5.0.0");
    const result = await dUtils.getTemplateUrl("ts", getLatestVersion, Platform.VSCode);
    assert.isUndefined(result);
  });

  it("setGeneralSensitivityLabel happy path", async () => {
    const gtools = new MockTools();
    setTools(gtools);
    const context = createContext();
    const manifestPath = "test/manifest.json";

    const tokenStub = sandbox.stub(context.tokenProvider!.m365TokenProvider, "getStatus").resolves(
      ok({
        status: signedIn,
        token: "fake-token",
      })
    );
    const getLabelStub = sandbox.stub(GraphClient.prototype, "getGeneralSentivityLabel").resolves(
      ok({
        id: "general-label-id",
        displayName: "General",
        name: "General Label",
        description: "General Label Description",
      })
    );
    const DAManifest = {
      version: "v1.4" as const,
      name: "test-agent",
      description: "test agent description",
    };

    const readStub = sandbox
      .stub(copilotGptManifestUtils, "readDeclarativeAgentManifestFile")
      .resolves(ok(DAManifest as DeclarativeAgentManifest));

    const writeStub = sandbox
      .stub(copilotGptManifestUtils, "writeDeclarativeAgentManifestFile")
      .resolves(ok(undefined));

    await setGeneralSensitivityLabel(context, manifestPath);

    assert.isTrue(tokenStub.calledOnce);
    assert.isTrue(getLabelStub.calledOnceWith("fake-token"));
    assert.isTrue(readStub.calledOnceWith(manifestPath));
    assert.isTrue(writeStub.calledOnce);

    // Verify the manifest was updated by checking the writeStub was called correctly
    assert.equal(writeStub.firstCall.args[1], manifestPath);

    const sensitivityLabel = (DAManifest as any).sensitivity_label;
    assert.equal(sensitivityLabel?.id, "general-label-id");
  });

  it("setGeneralSensitivityLabel failed", async () => {
    const gtools = new MockTools();
    setTools(gtools);
    const context = createContext();
    const manifestPath = "test/manifest.json";

    sandbox.stub(context, "tokenProvider").value(undefined);
    const getLabelStub = sandbox
      .stub(GraphClient.prototype, "getGeneralSentivityLabel")
      .resolves(err(new Error("Failed to get label") as any));
    const DAManifest = {
      version: "v1.4" as const,
      name: "test-agent",
      description: "test agent description",
    };

    const readStub = sandbox
      .stub(copilotGptManifestUtils, "readDeclarativeAgentManifestFile")
      .resolves(ok(DAManifest as DeclarativeAgentManifest));

    const writeStub = sandbox
      .stub(copilotGptManifestUtils, "writeDeclarativeAgentManifestFile")
      .resolves(ok(undefined));

    await setGeneralSensitivityLabel(context, manifestPath);

    assert.isTrue(readStub.notCalled);
    assert.isTrue(writeStub.notCalled);

    const sensitivityLabel = (DAManifest as any).sensitivity_label;
    assert.isUndefined(sensitivityLabel);
  });
});

describe("templateHelper unit test cases", () => {
  it("should return true when TEMPLATE_VERSION is set to local", () => {
    const restore = mockedEnv({
      TEMPLATE_VERSION: "local",
    });
    const result = useLocalTemplate();
    assert.isTrue(result);
    restore();
  });

  it("should return false when TEMPLATE_VERSION is not set to local", () => {
    const restore = mockedEnv({
      TEMPLATE_VERSION: "1.0.0",
    });
    const mockPackageJson = {
      version: "3.0.0",
    };
    const templateHelper = proxyquire("../../../src/component/generator/templateHelper", {
      "../../../package.json": mockPackageJson,
    });
    const result = templateHelper.useLocalTemplate();
    assert.isFalse(result);
    restore();
  });

  it("should return true when package version contains alpha", () => {
    const restore = mockedEnv({
      TEMPLATE_VERSION: "",
    });
    const mockPackageJson = {
      version: "3.0.0-alpha.1",
    };
    const templateHelper = proxyquire("../../../src/component/generator/templateHelper", {
      "../../../package.json": mockPackageJson,
    });
    const result = templateHelper.useLocalTemplate();
    assert.isTrue(result);
    restore();
  });

  it("should return false when package version is stable and TEMPLATE_VERSION is not local", () => {
    const restore = mockedEnv({
      TEMPLATE_VERSION: "",
    });
    const mockPackageJson = {
      version: "3.0.0",
    };
    const templateHelper = proxyquire("../../../src/component/generator/templateHelper", {
      "../../../package.json": mockPackageJson,
    });
    const result = templateHelper.useLocalTemplate();
    assert.isFalse(result);
    restore();
  });

  it("should return false when package version contains beta", () => {
    const restore = mockedEnv({
      TEMPLATE_VERSION: "",
    });
    const mockPackageJson = {
      version: "3.0.0-beta.1",
    };
    const templateHelper = proxyquire("../../../src/component/generator/templateHelper", {
      "../../../package.json": mockPackageJson,
    });
    const result = templateHelper.useLocalTemplate();
    assert.isFalse(result);
    restore();
  });

  it("should return false when package version contains rc", () => {
    const restore = mockedEnv({
      TEMPLATE_VERSION: "",
    });
    const mockPackageJson = {
      version: "3.0.0-rc.1",
    };
    const templateHelper = proxyquire("../../../src/component/generator/templateHelper", {
      "../../../package.json": mockPackageJson,
    });
    const result = templateHelper.useLocalTemplate();
    assert.isFalse(result);
    restore();
  });
});

describe("getTemplateVSLatestVersion", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("should return the max satisfying version matching vsVersionPattern", async () => {
    // Build tag list from the actual vsVersionPattern so the test stays valid after version bumps
    const base = templateConfig.vsVersionPattern.replace("~", ""); // e.g. "18.6"
    const [major, minor] = base.split(".");
    const nextMinor = `${major}.${parseInt(minor) + 1}`;
    // shared tag list contains both VSC and VS tags
    const tagList = `templates@6.6.0\ntemplates@6.6.1\ntemplates-vs@${base}.0\ntemplates-vs@${base}.1\ntemplates-vs@${nextMinor}.0\n`;
    sandbox.stub(requestUtils, "sendRequestWithTimeout").resolves({ data: tagList } as any);

    const result = await getTemplateVSLatestVersion();
    // ~base matches base.x only, not nextMinor.x; VSC tags are ignored
    assert.strictEqual(result, `${base}.1`);
  });

  it("should handle CRLF line endings in tag list", async () => {
    const base = templateConfig.vsVersionPattern.replace("~", "");
    const tagList = `templates@6.6.1\r\ntemplates-vs@${base}.0\r\ntemplates-vs@${base}.1\r\n`;
    sandbox.stub(requestUtils, "sendRequestWithTimeout").resolves({ data: tagList } as any);

    const result = await getTemplateVSLatestVersion();
    assert.strictEqual(result, `${base}.1`);
  });

  it("should throw when no version satisfies vsVersionPattern", async () => {
    // only non-VS tags and old VS tags — none match ~18.6
    const tagList = "templates@6.6.1\ntemplates-vs@17.0.0\ntemplates-vs@17.1.0\n";
    sandbox.stub(requestUtils, "sendRequestWithTimeout").resolves({ data: tagList } as any);

    try {
      await getTemplateVSLatestVersion();
      assert.fail("Expected error to be thrown");
    } catch (e: any) {
      assert.include(e.message, "Failed to find valid VS template version");
    }
  });

  it("should propagate network errors from fetchTagList", async () => {
    sandbox.stub(requestUtils, "sendRequestWithTimeout").rejects(new Error("Network timeout"));

    try {
      await getTemplateVSLatestVersion();
      assert.fail("Expected error to be thrown");
    } catch (e: any) {
      assert.include(e.message, "Network timeout");
    }
  });
});
