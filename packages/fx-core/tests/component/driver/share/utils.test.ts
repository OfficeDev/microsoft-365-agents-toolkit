import { err, ok, UserError } from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import chai from "chai";
import fs from "fs-extra";
import "mocha";
import os from "os";
import path from "path";
import sinon from "sinon";
import { parseShareAppActionYamlConfig } from "../../../../src/component/driver/share/utils";
import { Constants } from "../../../../src/component/driver/teamsApp/constants";
import { envUtil } from "../../../../src/component/utils/envUtil";
import { metadataUtil } from "../../../../src/component/utils/metadataUtil";
import { pathUtils } from "../../../../src/component/utils/pathUtils";

// Helper function to create a temporary ZIP file
function createMockZipFile(): string {
  const tempDir = os.tmpdir();
  const tempZipPath = path.join(tempDir, "mockAppPackage.zip");
  const zip = new AdmZip();
  zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify({ id: "mockManifestId" })));
  zip.writeZip(tempZipPath);
  return tempZipPath;
}

function createMockZipFileWithoutManifest(): string {
  const tempDir = os.tmpdir();
  const tempZipPath = path.join(tempDir, "mockAppPackageWithoutManifest.zip");
  const zip = new AdmZip();
  zip.writeZip(tempZipPath);
  return tempZipPath;
}

function createMockZipFileWithoutManifestId(): string {
  const tempDir = os.tmpdir();
  const tempZipPath = path.join(tempDir, "mockAppPackageWithoutManifestId.zip");
  const zip = new AdmZip();
  zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify({})));
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
        deploy: {
          driverDefs: [
            {
              uses: "teamsApp/extendToM365",
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
    process.env["parseShareAppActionYamlConfigMockTitleIdName"] = "mockTitleId";
    process.env["parseShareAppActionYamlConfigMockAppIdName"] = "mockAppId";
    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isOk());
    if (result.isOk()) {
      chai.assert.deepEqual(result.value, {
        teamsappId: "mockManifestId",
        titleId: "mockTitleId",
        appId: "mockAppId",
      });
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
        deploy: {
          driverDefs: [
            {
              uses: "teamsApp/extendToM365",
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
    const mockZipPath = createMockZipFileWithoutManifest();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        deploy: {
          driverDefs: [
            {
              uses: "teamsApp/extendToM365",
              with: { appPackagePath: mockZipPath },
            },
          ],
        },
      } as any)
    );
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(fs, "existsSync").returns(true);

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
        deploy: {
          driverDefs: [
            {
              uses: "teamsApp/extendToM365",
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
      chai.assert.equal(result.error.name, "Share");
    }
  });

  it("should return error when projectModel.share or driverDefs is missing", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(ok({} as any));

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(result.error.name, "Share");
    }
  });

  it("should return error when provision has no driverDefs and deploy is missing", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        provision: {},
      } as any)
    );

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(result.error.name, "Share");
      chai.assert.include(result.error.message, 'Unable to find the "provision"');
    }
  });

  it("should return error when deploy has no driverDefs and provision is missing", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        deploy: {},
      } as any)
    );

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(result.error.name, "Share");
      chai.assert.include(result.error.message, 'Unable to find the "provision"');
    }
  });

  it("should return error when both provision and deploy have no driverDefs", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        provision: {},
        deploy: {},
      } as any)
    );

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(result.error.name, "Share");
      chai.assert.include(result.error.message, 'Unable to find the "provision"');
    }
  });

  it("should find extendToM365Action in provision when deploy doesn't have it", async () => {
    const mockZipPath = createMockZipFile();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        provision: {
          driverDefs: [
            {
              uses: "teamsApp/extendToM365",
              with: { appPackagePath: mockZipPath } as any,
              writeToEnvironmentFile: {
                titleId: "parseShareAppActionYamlConfigMockTitleIdName",
                appId: "parseShareAppActionYamlConfigMockAppIdName",
              } as any,
            },
          ],
        },
        deploy: {
          driverDefs: [
            {
              uses: "someOtherAction",
            },
          ],
        },
      } as any)
    );
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(fs, "existsSync").returns(true);
    process.env["parseShareAppActionYamlConfigMockTitleIdName"] = "mockTitleId";
    process.env["parseShareAppActionYamlConfigMockAppIdName"] = "mockAppId";

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isOk());
    if (result.isOk()) {
      chai.assert.deepEqual(result.value, {
        teamsappId: "mockManifestId",
        titleId: "mockTitleId",
        appId: "mockAppId",
      });
    }

    process.env["parseShareAppActionYamlConfigMockTitleIdName"] = undefined;
    process.env["parseShareAppActionYamlConfigMockAppIdName"] = undefined;
  });

  it("should return error when share action is not found in either provision or deploy", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        provision: {
          driverDefs: [
            {
              uses: "someOtherAction",
            },
          ],
        },
        deploy: {
          driverDefs: [
            {
              uses: "anotherAction",
            },
          ],
        },
      } as any)
    );

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(result.error.name, "Share");
      chai.assert.include(result.error.message, 'Unable to find the "copilotAgent/publish"');
    }
  });

  it("should find copilotAgent/publish action in provision", async () => {
    const mockZipPath = createMockZipFile();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        provision: {
          driverDefs: [
            {
              uses: "copilotAgent/publish",
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
    process.env["parseShareAppActionYamlConfigMockTitleIdName"] = "mockTitleId";
    process.env["parseShareAppActionYamlConfigMockAppIdName"] = "mockAppId";

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isOk());
    if (result.isOk()) {
      chai.assert.deepEqual(result.value, {
        teamsappId: "mockManifestId",
        titleId: "mockTitleId",
        appId: "mockAppId",
      });
    }

    process.env["parseShareAppActionYamlConfigMockTitleIdName"] = undefined;
    process.env["parseShareAppActionYamlConfigMockAppIdName"] = undefined;
  });

  it("should return error when extendToM365Action is missing", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(ok({ deploy: { driverDefs: [] } } as any));

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(result.error.name, "Share");
    }
  });

  it("should return error when readEnv fails", async () => {
    const mockZipPath = createMockZipFile();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        deploy: {
          driverDefs: [
            {
              uses: "teamsApp/extendToM365",
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
    sandbox.stub(fs, "existsSync").returns(true);
    sandbox
      .stub(envUtil, "readEnv")
      .resolves(err(new UserError("FxCore", "EnvError", "Failed to read env")));

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(result.error.name, "EnvError");
    }
  });

  it("should return error when resolvedAppPackagePath does not exist", async () => {
    const mockZipPath = createMockZipFile();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        deploy: {
          driverDefs: [
            {
              uses: "teamsApp/extendToM365",
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
    sandbox.stub(fs, "existsSync").returns(false);

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(result.error.name, "Share");
    }
  });

  it("should return error when manifestId is missing", async () => {
    const mockZipPath = createMockZipFileWithoutManifestId();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        deploy: {
          driverDefs: [
            {
              uses: "teamsApp/extendToM365",
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
    process.env["parseShareAppActionYamlConfigMockTitleIdName"] = "mockTitleId";
    process.env["parseShareAppActionYamlConfigMockAppIdName"] = "mockAppId";

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(result.error.name, "Share to Users");
    }
  });

  it("should return error when shared ids are missing in environment", async () => {
    const mockZipPath = createMockZipFile();
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(
      ok({
        deploy: {
          driverDefs: [
            {
              uses: "teamsApp/extendToM365",
              with: { appPackagePath: mockZipPath } as any,
              writeToEnvironmentFile: {
                titleId: "parseShareAppActionYamlConfigMissingTitleIdName",
                appId: "parseShareAppActionYamlConfigMissingAppIdName",
              } as any,
            },
          ],
        },
      } as any)
    );
    sandbox.stub(envUtil, "readEnv").resolves(ok({}));
    sandbox.stub(fs, "existsSync").returns(true);
    delete process.env["parseShareAppActionYamlConfigMissingTitleIdName"];
    delete process.env["parseShareAppActionYamlConfigMissingAppIdName"];

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(result.error.name, "Share");
      chai.assert.include(result.error.message, "Unable to get title id or app id");
    }
  });

  it("should error with yamlConfigNotSupported for version < 1.10.0", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    sandbox.stub(metadataUtil, "parse").resolves(ok({ version: "v1.9" } as any));

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      chai.assert.instanceOf(result.error, UserError);
      chai.assert.equal(result.error.name, "Share");
      // message should include the unsupported version string
      chai.assert.include(
        result.error.message,
        "Share feature only supports m365agents.yml version v1.10 or above"
      );
    }
  });

  it("should proceed when version >= 1.10.0", async () => {
    sandbox.stub(pathUtils, "getYmlFilePath").returns("mockTemplatePath");
    // minimal model with only version; subsequent checks will return yamlConfigNotFound
    sandbox.stub(metadataUtil, "parse").resolves(ok({ version: "v1.10" } as any));

    const result = await parseShareAppActionYamlConfig("mockProjectPath");

    chai.assert.isTrue(result.isErr());
    if (result.isErr()) {
      // ensure not failing due to version gate by checking generic Share error
      chai.assert.instanceOf(result.error, UserError);
      chai.assert.equal(result.error.name, "Share");
    }
  });
});
