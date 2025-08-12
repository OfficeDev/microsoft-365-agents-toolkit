// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { assert } from "chai";
import * as sinon from "sinon";
import "mocha";
import {
  getTemplateZipUrlByVersion,
  getTemplateUrl,
  setGeneralSensitivityLabel,
} from "../../../src/component/generator/utils";
import { Platform, ok, signedIn, DeclarativeAgentManifest, err } from "@microsoft/teamsfx-api";
import mockedEnv from "mocked-env";
import proxyquire from "proxyquire";
import { createContext, setTools } from "../../../src/common/globalVars";
import { copilotGptManifestUtils } from "../../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import { GraphClient } from "../../../src/client/graphClient";
import { MockTools } from "../../core/utils";

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
      "../../common/templates-config.json": mockSettings,
    });
    const getLatestVersion = () => Promise.resolve("18.0.0");
    const result = await dUtils.getTemplateUrl("csharp", getLatestVersion, Platform.VS);
    const expectedUrl =
      "https://github.com/OfficeDev/microsoft-365-agents-toolkit/releases/download/templates-vs@18.0.0/csharp.zip";
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
