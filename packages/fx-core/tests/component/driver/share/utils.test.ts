import sinon from "sinon";
import "mocha";
import chai from "chai";
import os from "os";
import path, { parse } from "path";
import AdmZip from "adm-zip";
import fs from "fs-extra";
import { Constants } from "../../../../src/component/driver/teamsApp/constants";
import { pathUtils } from "../../../../src/component/utils/pathUtils";
import { metadataUtil } from "../../../../src/component/utils/metadataUtil";
import { err, ok, UserError } from "@microsoft/teamsfx-api";
import { envUtil } from "../../../../src/component/utils/envUtil";
import { parseShareAppActionYamlConfig } from "../../../../src/component/driver/share/utils";
// Helper function to create a temporary ZIP file
function createMockZipFile(): string {
  const tempDir = os.tmpdir();
  const tempZipPath = path.join(tempDir, "mockAppPackage.zip");
  const zip = new AdmZip();
  zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify({ id: "mockManifestId" })));
  zip.writeZip(tempZipPath);
  return tempZipPath;
}

describe("parseShareAppActionYamlConfig", () => {
  const sandbox = sinon.createSandbox();
  afterEach(() => {
    sandbox.restore();
  });

  it("should return manifestId, sharedTitleId, and sharedAppId when config is valid", async () => {
    const mockZipPath = createMockZipFile();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        share: {
          driverDefs: [
            {
              uses: "teamsApp/shareToOthers",
              with: { appPackagePath: mockZipPath } as any,
              writeToEnvironmentFile: {
                titleId: "parseShareAppActionYamlConfigMockTitleIdName",
                appId: "parseShareAppActionYamlConfigMockAppIdName",
              } as any,
            },
          ],
        },
      } as any)
    );
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(fs, "existsSync").returns(true);
    const admZipInstance = new AdmZip();
    sandbox.stub(admZipInstance, "getEntries").returns([
      {
        entryName: Constants.MANIFEST_FILE,
        getData: () => Buffer.from(JSON.stringify({ id: "mockManifestId" })),
      },
    ] as any);
    sandbox.stub(AdmZip, "prototype").returns(admZipInstance);
    process.env["parseShareAppActionYamlConfigMockTitleIdName"] = "mockTitleId";
    process.env["parseShareAppActionYamlConfigMockAppIdName"] = "mockAppId";
    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isOk());
    if (result.isOk()) {
      chai.assert.deepEqual(result.value, ["mockManifestId", "mockTitleId", "mockAppId"]);
    }

    process.env["parseShareAppActionYamlConfigMockTitleIdName"] = undefined;
    process.env["parseShareAppActionYamlConfigMockAppIdName"] = undefined;
  });

  it("should return error when yaml config is invalid", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox
      .stub(metadataUtil, "parse")
      .resolves(err(new UserError("FxCore", "InvalidYaml", "Invalid yaml config")));

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.instanceOf(result.error, UserError);
      chai.assert.equal(result.error.name, "InvalidYaml");
    }
  });

  it("should return error when appPackagePath is missing", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        share: {
          driverDefs: [
            {
              uses: "teamsApp/shareToOthers",
            },
          ],
        },
      } as any)
    );

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.instanceOf(result.error, UserError);
      chai.assert.equal(result.error.name, "Share to Users");
    }
  });

  it("should return error when manifest file is missing", async () => {
    const mockZipPath = createMockZipFile();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        share: {
          driverDefs: [
            {
              uses: "teamsApp/shareToOthers",
              with: { appPackagePath: mockZipPath },
            },
          ],
        },
      } as any)
    );
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(fs, "existsSync").returns(true);
    const admZipInstance = new AdmZip();
    sandbox.stub(admZipInstance, "getEntries").returns([]);
    sandbox.stub(AdmZip, "prototype").returns(admZipInstance);

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.instanceOf(result.error, UserError);
      chai.assert.equal(result.error.name, "Share to Users");
    }
  });

  it("should return error when sharedTitleId or sharedAppId is missing", async () => {
    const mockZipPath = createMockZipFile();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        share: {
          driverDefs: [
            {
              uses: "teamsApp/shareToOthers",
              with: { appPackagePath: mockZipPath },
              writeToEnvironmentFile: {},
            },
          ],
        },
      } as any)
    );
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(fs, "existsSync").returns(true);
    const admZipInstance = new AdmZip();
    sandbox.stub(admZipInstance, "getEntries").returns([
      {
        entryName: Constants.MANIFEST_FILE,
        getData: () => Buffer.from(JSON.stringify({ id: "mockManifestId" })),
      },
    ] as any);
    sandbox.stub(AdmZip, "prototype").returns(admZipInstance);

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.instanceOf(result.error, UserError);
      chai.assert.equal(result.error.name, "Share to Users");
    }
  });
});
