// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import AdmZip from "adm-zip";
import { assert } from "chai";
import fs from "fs-extra";
import { RestoreFn } from "mocked-env";
import * as os from "os";
import * as path from "path";
import sinon from "sinon";
import { FxCore } from "../../src";
import { setTools } from "../../src/common/globalVars";
import templateConfigModule from "../../src/common/templates-config.json";
import "../../src/component/feature/sso";
import { fxCoreDeps } from "../../src/core/FxCore";
import { MockTools, randomAppName } from "./utils";

const tools = new MockTools();

async function mockV3Project(): Promise<string> {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  // await fs.move(path.join(__dirname, "../sampleV3"), path.join(os.tmpdir(), appName));
  await fs.copy(path.join(__dirname, "../samples/sampleV3/"), path.join(projectPath));
  return appName;
}

async function mockCliUninstallProject(): Promise<string> {
  const appName = randomAppName();
  const projectPath = path.join(os.tmpdir(), appName);
  await fs.copy(path.join(__dirname, "../samples/uninstall/"), path.join(projectPath));
  return appName;
}

async function deleteTestProject(appName: string) {
  await fs.remove(path.join(os.tmpdir(), appName));
}

describe("fetchOnlineTemplateMetadata", () => {
  const sandbox = sinon.createSandbox();
  let core: FxCore;
  let mockedEnvRestore: RestoreFn | undefined;

  beforeEach(() => {
    setTools(tools);
    core = new FxCore(tools);
  });

  afterEach(() => {
    sandbox.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
      mockedEnvRestore = undefined;
    }
  });

  it("should skip download when using local template", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(true);

    const result = await core.fetchOnlineTemplateMetadata();

    assert.isTrue(result.isOk());
    if (result.isOk()) {
      assert.isUndefined(result.value);
    }
  });

  it("should download metadata for rc version when coreVersion contains 'rc'", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "tagPrefix").value("templates@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getCoreVersion").returns("1.0.0-rc.1");

    const mockZip = new AdmZip();
    const fetchZipStub = sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    const unzipStub = sandbox.stub(fxCoreDeps, "unzip").resolves();

    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "ensureDir").resolves();
    const writeFileStub = sandbox.stub(fs, "writeFile").resolves();

    const result = await core.fetchOnlineTemplateMetadata();

    assert.isTrue(result.isOk());
    assert.isTrue(fetchZipStub.calledOnce);
    assert.isTrue(
      fetchZipStub.calledWith(
        "https://example.com/releases/download/templates@0.0.0-rc/metadata.zip"
      )
    );
    assert.isTrue(unzipStub.calledOnce);
    assert.isTrue(writeFileStub.calledOnce);
  });

  it("should download metadata for stable version", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "tagPrefix").value("templates@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getCoreVersion").returns("1.0.0");

    const getTemplateLatestVersionStub = sandbox
      .stub(fxCoreDeps, "getTemplateLatestVersion")
      .resolves("2.0.0");
    const mockZip = new AdmZip();
    const fetchZipStub = sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    const unzipStub = sandbox.stub(fxCoreDeps, "unzip").resolves();

    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "ensureDir").resolves();
    const writeFileStub = sandbox.stub(fs, "writeFile").resolves();

    const result = await core.fetchOnlineTemplateMetadata();

    assert.isTrue(result.isOk());
    assert.isTrue(getTemplateLatestVersionStub.calledOnce);
    assert.isTrue(fetchZipStub.calledOnce);
    assert.isTrue(
      fetchZipStub.calledWith("https://example.com/releases/download/templates@2.0.0/metadata.zip")
    );
    assert.isTrue(unzipStub.calledOnce);
    assert.isTrue(writeFileStub.calledWith(sinon.match.string, "2.0.0", { encoding: "utf-8" }));
  });

  it("should skip download when cached version matches latest version", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(fxCoreDeps, "getCoreVersion").returns("1.0.0");

    sandbox.stub(fxCoreDeps, "getTemplateLatestVersion").resolves("2.0.0");
    const fetchZipStub = sandbox.stub(fxCoreDeps, "fetchZipFromUrl");
    const unzipStub = sandbox.stub(fxCoreDeps, "unzip");

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "readFile").resolves("2.0.0" as any);
    sandbox.stub(fs, "writeFile").resolves();

    const result = await core.fetchOnlineTemplateMetadata();

    assert.isTrue(result.isOk());
    assert.equal(fetchZipStub.called, false);
    assert.equal(unzipStub.called, false);
  });

  it("should download when cached version file does not exist", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "tagPrefix").value("templates@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getCoreVersion").returns("1.0.0");

    sandbox.stub(fxCoreDeps, "getTemplateLatestVersion").resolves("2.0.0");
    const mockZip = new AdmZip();
    const fetchZipStub = sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    const unzipStub = sandbox.stub(fxCoreDeps, "unzip").resolves();

    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").resolves();

    const result = await core.fetchOnlineTemplateMetadata();

    assert.isTrue(result.isOk());
    assert.isTrue(fetchZipStub.calledOnce);
    assert.isTrue(unzipStub.calledOnce);
  });

  it("should download when cached version differs from latest version", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "tagPrefix").value("templates@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getCoreVersion").returns("1.0.0");

    sandbox.stub(fxCoreDeps, "getTemplateLatestVersion").resolves("2.0.0");
    const mockZip = new AdmZip();
    const fetchZipStub = sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    const unzipStub = sandbox.stub(fxCoreDeps, "unzip").resolves();

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "readFile").resolves("1.0.0" as any); // Old cached version
    sandbox.stub(fs, "writeFile").resolves();

    const result = await core.fetchOnlineTemplateMetadata();

    assert.isTrue(result.isOk());
    assert.isTrue(fetchZipStub.calledOnce);
    assert.isTrue(unzipStub.calledOnce);
  });

  it("should re-download when cached version file is corrupted", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "tagPrefix").value("templates@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getCoreVersion").returns("1.0.0");

    sandbox.stub(fxCoreDeps, "getTemplateLatestVersion").resolves("2.0.0");
    const mockZip = new AdmZip();
    const fetchZipStub = sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    const unzipStub = sandbox.stub(fxCoreDeps, "unzip").resolves();

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "readFile").rejects(new Error("File read error"));
    sandbox.stub(fs, "writeFile").resolves();

    const result = await core.fetchOnlineTemplateMetadata();

    assert.isTrue(result.isOk());
    assert.isTrue(fetchZipStub.calledOnce);
    assert.isTrue(unzipStub.calledOnce);
  });

  it("should handle alpha version correctly", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "tagPrefix").value("templates@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getCoreVersion").returns("1.0.0-alpha.1");

    const mockZip = new AdmZip();
    const fetchZipStub = sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    const unzipStub = sandbox.stub(fxCoreDeps, "unzip").resolves();

    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").resolves();

    const result = await core.fetchOnlineTemplateMetadata();

    assert.isTrue(result.isOk());
    assert.isTrue(
      fetchZipStub.calledWith(
        "https://example.com/releases/download/templates@0.0.0-rc/metadata.zip"
      )
    );
  });

  it("should handle beta version correctly", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "tagPrefix").value("templates@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getCoreVersion").returns("1.0.0-beta.1");

    const mockZip = new AdmZip();
    const fetchZipStub = sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    const unzipStub = sandbox.stub(fxCoreDeps, "unzip").resolves();

    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").resolves();

    const result = await core.fetchOnlineTemplateMetadata();

    assert.isTrue(result.isOk());
    assert.isTrue(
      fetchZipStub.calledWith(
        "https://example.com/releases/download/templates@0.0.0-rc/metadata.zip"
      )
    );
  });

  it("should return error when fetchZipFromUrl fails", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "tagPrefix").value("templates@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getCoreVersion").returns("1.0.0");

    sandbox.stub(fxCoreDeps, "getTemplateLatestVersion").resolves("2.0.0");
    sandbox
      .stub(fxCoreDeps, "fetchZipFromUrl")
      .rejects(new Error("Network error: Failed to fetch"));

    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "ensureDir").resolves();

    const result = await core.fetchOnlineTemplateMetadata();

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.source, "FetchOnlineTemplateMetadata");
      assert.equal(result.error.name, "DownloadFailed");
      assert.include(result.error.message, "Network error: Failed to fetch");
    }
  });

  it("should return error when unzip fails", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "tagPrefix").value("templates@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getCoreVersion").returns("1.0.0");

    sandbox.stub(fxCoreDeps, "getTemplateLatestVersion").resolves("2.0.0");
    const mockZip = new AdmZip();
    sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    sandbox.stub(fxCoreDeps, "unzip").rejects(new Error("Unzip failed: Invalid archive"));

    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "ensureDir").resolves();

    const result = await core.fetchOnlineTemplateMetadata();

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.source, "FetchOnlineTemplateMetadata");
      assert.equal(result.error.name, "DownloadFailed");
      assert.include(result.error.message, "Unzip failed: Invalid archive");
    }
  });

  it("should return error when fs.writeFile fails", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "tagPrefix").value("templates@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getCoreVersion").returns("1.0.0");

    sandbox.stub(fxCoreDeps, "getTemplateLatestVersion").resolves("2.0.0");
    const mockZip = new AdmZip();
    sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    sandbox.stub(fxCoreDeps, "unzip").resolves();

    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").rejects(new Error("Permission denied"));

    const result = await core.fetchOnlineTemplateMetadata();

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.source, "FetchOnlineTemplateMetadata");
      assert.equal(result.error.name, "DownloadFailed");
      assert.include(result.error.message, "Permission denied");
    }
  });

  it("should use correct metadata directory path", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "tagPrefix").value("templates@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getCoreVersion").returns("1.0.0");

    sandbox.stub(fxCoreDeps, "getTemplateLatestVersion").resolves("2.0.0");
    const mockZip = new AdmZip();
    sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    const unzipStub = sandbox.stub(fxCoreDeps, "unzip").resolves();

    sandbox.stub(fs, "pathExists").resolves(false);
    const ensureDirStub = sandbox.stub(fs, "ensureDir").resolves();
    const writeFileStub = sandbox.stub(fs, "writeFile").resolves();

    const expectedMetadataDir = path.join(os.homedir(), ".fx");

    const result = await core.fetchOnlineTemplateMetadata();

    assert.isTrue(result.isOk());
    assert.isTrue(ensureDirStub.calledWith(expectedMetadataDir));
    assert.isTrue(unzipStub.calledWith(mockZip, expectedMetadataDir));
    assert.isTrue(
      writeFileStub.calledWith(path.join(expectedMetadataDir, "template-version.txt"), "2.0.0", {
        encoding: "utf-8",
      })
    );
  });
});

describe("fetchOnlineTemplateMetadataForVS", () => {
  const sandbox = sinon.createSandbox();
  let core: FxCore;
  let mockedEnvRestore: RestoreFn | undefined;

  beforeEach(() => {
    setTools(tools);
    core = new FxCore(tools);
  });

  afterEach(() => {
    sandbox.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
      mockedEnvRestore = undefined;
    }
  });

  it("should skip download when using local template", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(true);

    const result = await core.fetchOnlineTemplateMetadataForVS();

    assert.isTrue(result.isOk());
    if (result.isOk()) {
      assert.isUndefined(result.value);
    }
  });

  it("should download metadata when version file does not exist (stable fx-core)", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(fxCoreDeps, "getCoreVersion").returns("1.0.0"); // stable
    sandbox.stub(templateConfigModule, "vstagPrefix").value("templates-vs@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getTemplateVSLatestVersion").resolves("18.4.1");
    const mockZip = new AdmZip();
    const fetchZipStub = sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    const unzipStub = sandbox.stub(fxCoreDeps, "unzip").resolves();

    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "ensureDir").resolves();
    const writeFileStub = sandbox.stub(fs, "writeFile").resolves();

    const result = await core.fetchOnlineTemplateMetadataForVS();

    assert.isTrue(result.isOk());
    assert.isTrue(fetchZipStub.calledOnce);
    assert.isTrue(
      fetchZipStub.calledWith(
        "https://example.com/releases/download/templates-vs@18.4.1/metadata.zip"
      )
    );
    assert.isTrue(unzipStub.calledOnce);
    assert.isTrue(writeFileStub.calledWith(sinon.match.string, "18.4.1", { encoding: "utf-8" }));
  });

  it("should use rc templates for beta fx-core (VS pre-release test build)", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(fxCoreDeps, "getCoreVersion").returns("1.0.0-beta.1"); // beta = pre-stable test
    sandbox.stub(templateConfigModule, "vstagPrefix").value("templates-vs@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    const getVSLatestStub = sandbox.stub(fxCoreDeps, "getTemplateVSLatestVersion");
    const mockZip = new AdmZip();
    const fetchZipStub = sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    sandbox.stub(fxCoreDeps, "unzip").resolves();

    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "writeFile").resolves();

    const result = await core.fetchOnlineTemplateMetadataForVS();

    assert.isTrue(result.isOk());
    // beta should NOT call getTemplateVSLatestVersion
    assert.isFalse(getVSLatestStub.called);
    assert.isTrue(
      fetchZipStub.calledWith(
        "https://example.com/releases/download/templates-vs@0.0.0-rc/metadata.zip"
      )
    );
  });

  it("should skip download when cached version matches latest", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(fxCoreDeps, "getTemplateVSLatestVersion").resolves("18.4.1");
    const fetchZipStub = sandbox.stub(fxCoreDeps, "fetchZipFromUrl");
    const unzipStub = sandbox.stub(fxCoreDeps, "unzip");

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "readFile").resolves("18.4.1" as any);
    sandbox.stub(fs, "writeFile").resolves();

    const result = await core.fetchOnlineTemplateMetadataForVS();

    assert.isTrue(result.isOk());
    assert.isFalse(fetchZipStub.called);
    assert.isFalse(unzipStub.called);
  });

  it("should download when cached version differs from latest", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "vstagPrefix").value("templates-vs@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getTemplateVSLatestVersion").resolves("18.4.1");
    const mockZip = new AdmZip();
    const fetchZipStub = sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    const unzipStub = sandbox.stub(fxCoreDeps, "unzip").resolves();

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "readFile").resolves("18.4.0" as any);
    sandbox.stub(fs, "writeFile").resolves();

    const result = await core.fetchOnlineTemplateMetadataForVS();

    assert.isTrue(result.isOk());
    assert.isTrue(fetchZipStub.calledOnce);
    assert.isTrue(unzipStub.calledOnce);
  });

  it("should re-download when version file read throws", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "vstagPrefix").value("templates-vs@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getTemplateVSLatestVersion").resolves("18.4.1");
    const mockZip = new AdmZip();
    const fetchZipStub = sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    const unzipStub = sandbox.stub(fxCoreDeps, "unzip").resolves();

    sandbox.stub(fs, "pathExists").resolves(true);
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "readFile").rejects(new Error("File read error"));
    sandbox.stub(fs, "writeFile").resolves();

    const result = await core.fetchOnlineTemplateMetadataForVS();

    assert.isTrue(result.isOk());
    assert.isTrue(fetchZipStub.calledOnce);
    assert.isTrue(unzipStub.calledOnce);
  });

  it("should return error when getTemplateVSLatestVersion fails", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox
      .stub(fxCoreDeps, "getTemplateVSLatestVersion")
      .rejects(new Error("Failed to find valid VS template version"));

    sandbox.stub(fs, "ensureDir").resolves();

    const result = await core.fetchOnlineTemplateMetadataForVS();

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.source, "FetchOnlineTemplateMetadataForVS");
      assert.equal(result.error.name, "DownloadFailed");
      assert.include(result.error.message, "Failed to find valid VS template version");
    }
  });

  it("should return error when fetchZipFromUrl fails", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "vstagPrefix").value("templates-vs@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getTemplateVSLatestVersion").resolves("18.4.1");
    sandbox
      .stub(fxCoreDeps, "fetchZipFromUrl")
      .rejects(new Error("Download failed: 404 Not Found"));

    sandbox.stub(fs, "pathExists").resolves(false);
    sandbox.stub(fs, "ensureDir").resolves();

    const result = await core.fetchOnlineTemplateMetadataForVS();

    assert.isTrue(result.isErr());
    if (result.isErr()) {
      assert.equal(result.error.source, "FetchOnlineTemplateMetadataForVS");
      assert.equal(result.error.name, "DownloadFailed");
      assert.include(result.error.message, "Download failed: 404 Not Found");
    }
  });

  it("should use vs-metadata directory with correct version file path", async () => {
    sandbox.stub(fxCoreDeps, "useLocalTemplate").returns(false);
    sandbox.stub(templateConfigModule, "vstagPrefix").value("templates-vs@");
    sandbox
      .stub(templateConfigModule, "templateDownloadBaseURL")
      .value("https://example.com/releases/download");
    sandbox.stub(fxCoreDeps, "getTemplateVSLatestVersion").resolves("18.4.1");
    const mockZip = new AdmZip();
    sandbox.stub(fxCoreDeps, "fetchZipFromUrl").resolves(mockZip);
    const unzipStub = sandbox.stub(fxCoreDeps, "unzip").resolves();

    sandbox.stub(fs, "pathExists").resolves(false);
    const ensureDirStub = sandbox.stub(fs, "ensureDir").resolves();
    const writeFileStub = sandbox.stub(fs, "writeFile").resolves();

    const expectedMetadataDir = path.join(os.homedir(), ".fx", "vs-metadata");

    const result = await core.fetchOnlineTemplateMetadataForVS();

    assert.isTrue(result.isOk());
    assert.isTrue(ensureDirStub.calledWith(expectedMetadataDir));
    assert.isTrue(unzipStub.calledWith(mockZip, expectedMetadataDir));
    assert.isTrue(
      writeFileStub.calledWith(
        path.join(expectedMetadataDir, "template-vs-version.txt"),
        "18.4.1",
        { encoding: "utf-8" }
      )
    );
  });
});
