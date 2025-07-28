// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  DeclarativeCopilotManifestSchema,
  err,
  ok,
  Platform,
  TeamsManifest,
  TeamsManifestV1D19,
  TeamsManifestVDevPreview,
  UserError,
} from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import chai from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as sinon from "sinon";
import { featureFlagManager, FeatureFlags } from "../../../../src/common/featureFlags";
import { DriverContext } from "../../../../src/component/driver/interface/commonArgs";
import { CreateAppPackageDriver } from "../../../../src/component/driver/teamsApp/createAppPackage";
import { CreateAppPackageArgs } from "../../../../src/component/driver/teamsApp/interfaces/CreateAppPackageArgs";
import { copilotGptManifestUtils } from "../../../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import { manifestUtils } from "../../../../src/component/driver/teamsApp/utils/ManifestUtils";
import * as utils from "../../../../src/component/driver/util/utils";
import * as envFunctionUtils from "../../../../src/component/utils/envFunctionUtils";
import { ManifestType } from "../../../../src/component/utils/envFunctionUtils";
import { FileNotFoundError, JSONSyntaxError } from "../../../../src/error/common";
import { InvalidFileOutsideOfTheDirectotryError } from "../../../../src/error/teamsApp";
import { MockedM365Provider } from "../../../core/utils";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";

describe("teamsApp/createAppPackage", async () => {
  const teamsAppDriver = new CreateAppPackageDriver();
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    projectPath: "./",
    platform: Platform.VSCode,
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
    addTelemetryProperties: () => {},
  };
  let mockedEnvRestore: RestoreFn;
  const fakeUrl = "https://fake.com";
  const openapiServerPlaceholder = "TEAMSFX_TEST_API_URL";
  beforeEach(() => {
    mockedEnvRestore = mockedEnv({
      ["CONFIG_TEAMS_APP_NAME"]: "fakeName",
      [openapiServerPlaceholder]: fakeUrl,
      ["APP_NAME_SUFFIX"]: "test",
    });
  });

  afterEach(() => {
    sinon.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
  });

  it("happy path - with .generated folder", async () => {
    mockedEnvRestore = mockedEnv({
      TEAMSFX_TYPESPEC: "true",
      ["CONFIG_TEAMS_APP_NAME"]: "fakeName",
      [openapiServerPlaceholder]: fakeUrl,
    });

    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
    };

    const manifest = {
      manifestVersion: "1.19",
      icons: {
        color: "resources/color.png",
        outline: "resources/outline.png",
      },
      copilotAgents: {
        declarativeAgents: [
          {
            file: "resources/declarativeAgent.json",
            id: "dc1",
          },
        ],
      },
    } as TeamsManifest;
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
    sinon.stub(fs, "chmod").callsFake(async () => {});
    sinon.stub(fs, "existsSync").returns(true);
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(utils, "updateVersionForTeamsAppYamlFile").resolves();
    const writeFileStub = sinon.stub(fs, "writeFile").callsFake(async () => {});

    const driverContext: any = {
      m365TokenProvider: new MockedM365Provider(),
      projectPath: "./tests/plugins/resource/appstudio/resources-multi-env/templates/",
      platform: Platform.VSCode,
      logProvider: new MockedLogProvider(),
      ui: new MockedUserInteraction(),
      addTelemetryProperties: () => {},
    };
    const result = (await teamsAppDriver.execute(args, driverContext)).result;
    if (result.isErr()) {
      console.log(result.error);
    }
    chai.assert.isTrue(result.isOk());
    delete process.env["APP_NAME_SUFFIX"];
    await fs.remove(args.outputZipPath);
  });

  it("happy path - with .generated folder and ac in .generated folder", async () => {
    mockedEnvRestore = mockedEnv({
      TEAMSFX_TYPESPEC: "true",
      ["CONFIG_TEAMS_APP_NAME"]: "fakeName",
      [openapiServerPlaceholder]: fakeUrl,
    });

    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
    };

    const manifest = {
      manifestVersion: "1.19",
      icons: {
        color: "resources/color.png",
        outline: "resources/outline.png",
      },
      copilotAgents: {
        declarativeAgents: [
          {
            file: "resources/declarativeAgent.json",
            id: "dc1",
          },
        ],
      },
    } as TeamsManifest;
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
    sinon.stub(fs, "chmod").callsFake(async () => {});
    sinon.stub(fs, "existsSync").returns(true);
    sinon.stub(fs, "pathExists").callsFake((filePath) => {
      if (filePath.includes("adaptiveCards") && !filePath.includes(".generated")) {
        return false;
      } else {
        return true;
      }
    });
    sinon.stub(utils, "updateVersionForTeamsAppYamlFile").resolves();
    const writeFileStub = sinon.stub(fs, "writeFile").callsFake(async () => {});

    const driverContext: any = {
      m365TokenProvider: new MockedM365Provider(),
      projectPath: "./tests/plugins/resource/appstudio/resources-multi-env/templates/",
      platform: Platform.VSCode,
      logProvider: new MockedLogProvider(),
      ui: new MockedUserInteraction(),
      addTelemetryProperties: () => {},
    };
    const result = (await teamsAppDriver.execute(args, driverContext)).result;
    if (result.isErr()) {
      console.log(result.error);
    }
    chai.assert.isTrue(result.isOk());
    delete process.env["APP_NAME_SUFFIX"];
    await fs.remove(args.outputZipPath);
  });

  it("should throw error if file not exists case 1", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath: "fakepath",
      outputZipPath: "fakePath",
      outputJsonPath: "fakePath",
    };
    sinon.stub(manifestUtils, "getManifestV3").resolves(
      ok({
        manifestVersion: "1.0",
        icons: {
          color: "",
          outline: "",
        },
      } as TeamsManifest)
    );
    sinon.stub(fs, "pathExists").onFirstCall().resolves(false);
    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof FileNotFoundError);
    }
  });

  it("should throw error if file not exists case 2", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath: "fakepath",
      outputZipPath: "fakePath",
      outputJsonPath: "fakePath",
    };
    sinon.stub(manifestUtils, "getManifestV3").resolves(
      ok({
        manifestVersion: "1.0",
        icons: {
          color: "",
          outline: "",
        },
      } as TeamsManifest)
    );
    sinon.stub(fs, "pathExists").onFirstCall().resolves(true).onSecondCall().resolves(false);
    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof FileNotFoundError);
    }
  });

  it("should throw error if file not exists case 3", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath: "fakepath",
      outputZipPath: "fakePath",
      outputJsonPath: "fakePath",
    };
    const manifest = {
      manifestVersion: "1.19",
      icons: {
        color: "",
        outline: "",
      },
      localizationInfo: {
        additionalLanguages: [{ file: "aaa", languageTag: "zh" }],
        defaultLanguageTag: "en",
      },
    } as TeamsManifest;
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
    sinon
      .stub(fs, "pathExists")
      .onFirstCall()
      .resolves(true)
      .onSecondCall()
      .resolves(true)
      .onThirdCall()
      .resolves(false);
    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof FileNotFoundError);
    }
  });

  it("should throw error if color32x32 does not exist", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath: "fakepath",
      outputZipPath: "fakePath",
      outputJsonPath: "fakePath",
    };
    const manifest = {
      manifestVersion: "1.21",
      icons: {
        color: "",
        outline: "",
        color32x32: "notExist.png",
      },
    } as TeamsManifest;
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
    sinon.stub(fs, "pathExists").callsFake((filePath) => {
      if (filePath.includes("notExist.png")) {
        return false;
      }
      return true;
    });
    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof FileNotFoundError);
    }
  });

  it("should throw error if file not exists case 4", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputJsonPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
    };

    const manifest = {
      manifestVersion: "1.19",
      composeExtensions: [
        {
          composeExtensionType: "apiBased",
          apiSpecificationFile: "resources/openai.yml",
          commands: [
            {
              id: "GET /repairs",
              apiResponseRenderingTemplateFile: "resources/repairs.json",
              title: "fake",
            },
          ],
          botId: "",
        },
      ],
      icons: {
        color: "resources/color.png",
        outline: "resources/outline.png",
      },
    } as TeamsManifest;
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));

    sinon.stub(fs, "pathExists").callsFake((filePath) => {
      if (filePath.includes("openai.yml")) {
        return false;
      } else {
        return true;
      }
    });
    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof FileNotFoundError);
    }
  });

  it("should throw error if file not exists case 5", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputJsonPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
    };
    sinon.stub(fs, "pathExists").callsFake((filePath) => {
      if (filePath.includes("repairs.json")) {
        return false;
      } else {
        return true;
      }
    });

    const manifest = {
      manifestVersion: "1.19",
      composeExtensions: [
        {
          composeExtensionType: "apiBased",
          apiSpecificationFile: "resources/openai.yml",
          commands: [
            {
              id: "GET /repairs",
              apiResponseRenderingTemplateFile: "resources/repairs.json",
              title: "fake",
            },
          ],
          botId: "",
        },
      ],
      icons: {
        color: "resources/color.png",
        outline: "resources/outline.png",
      },
    } as TeamsManifest;
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof FileNotFoundError);
    }
  });

  it("should throw error if file not exists case 6", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputJsonPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
    };
    sinon.stub(fs, "pathExists").callsFake((filePath) => {
      if (filePath.includes("fake.json")) {
        return false;
      } else {
        return true;
      }
    });

    const manifest = {
      manifestVersion: "1.19",
      icons: {
        color: "",
        outline: "",
      },
      localizationInfo: {
        additionalLanguages: [{ file: "aaa", languageTag: "zh" }],
        defaultLanguageTag: "en",
        defaultLanguageFile: "fake.json",
      },
    } as TeamsManifest;
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof FileNotFoundError);
    }
  });

  it("invalid param error", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath: "",
      outputZipPath: "",
      outputJsonPath: "",
    };
    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal("InvalidActionInputError", result.error.name);
    }
  });

  it("version <= 1.6: happy path", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputJsonPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
    };

    const manifest = {
      manifestVersion: "1.17",
      localizationInfo: {
        defaultLanguageTag: "en",
        additionalLanguages: [
          {
            languageTag: "de",
            file: "resources/de.json",
          },
        ],
      },
      composeExtensions: [
        {
          composeExtensionType: "apiBased",
          apiSpecificationFile: "resources/openai.yml",
          commands: [
            {
              id: "GET /repairs",
              apiResponseRenderingTemplateFile: "resources/repairs.json",
              title: "fake",
            },
          ],
          botId: "",
        },
      ],
      icons: {
        color: "resources/color.png",
        outline: "resources/outline.png",
      },
    } as TeamsManifest;
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));

    sinon.stub(fs, "chmod").callsFake(async () => {});
    const writeFileStub = sinon.stub(fs, "writeFile").callsFake(async () => {});

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
    chai.assert(writeFileStub.calledOnce);
    if (await fs.pathExists(args.outputZipPath)) {
      const zip = new AdmZip(args.outputZipPath);

      let openapiContent = "";

      const entries = zip.getEntries();
      for (const e of entries) {
        const name = e.entryName;

        if (name.endsWith("openai.yml")) {
          const data = e.getData();
          openapiContent = data.toString("utf8");
          break;
        }
      }

      chai.assert(
        openapiContent != undefined &&
          openapiContent.length > 0 &&
          openapiContent.search(fakeUrl) >= 0 &&
          openapiContent.search(openapiServerPlaceholder) < 0
      );
      await fs.remove(args.outputZipPath);
    }
  });

  it("version > 1.6: happy path", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
    };

    const manifest = {
      manifestVersion: "1.19",
      localizationInfo: {
        defaultLanguageTag: "en",
        additionalLanguages: [
          {
            languageTag: "de",
            file: "resources/de.json",
          },
        ],
        defaultLanguageFile: "resources/de.json",
      },
      composeExtensions: [
        {
          composeExtensionType: "apiBased",
          apiSpecificationFile: "resources/openai.yml",
          commands: [
            {
              id: "GET /repairs",
              apiResponseRenderingTemplateFile: "resources/repairs.json",
              title: "fake",
            },
          ],
          botId: "",
        },
      ],
      icons: {
        color: "resources/color.png",
        outline: "resources/outline.png",
      },
    } as TeamsManifest;
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));

    sinon.stub(fs, "chmod").callsFake(async () => {});
    const writeFileStub = sinon.stub(fs, "writeFile").callsFake(async () => {});

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
    chai.assert(writeFileStub.calledOnce);
    if (await fs.pathExists(args.outputZipPath)) {
      const zip = new AdmZip(args.outputZipPath);

      let openapiContent = "";

      const entries = zip.getEntries();
      for (const e of entries) {
        const name = e.entryName;

        if (name.endsWith("openai.yml")) {
          const data = e.getData();
          openapiContent = data.toString("utf8");
          break;
        }
      }

      chai.assert(
        openapiContent != undefined &&
          openapiContent.length > 0 &&
          openapiContent.search(fakeUrl) >= 0 &&
          openapiContent.search(openapiServerPlaceholder) < 0
      );
      await fs.remove(args.outputZipPath);
    }
  });

  it("version > 1.6:should return error when placeholder is not resolved in openapi.yml", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
    };

    const manifest = {
      manifestVersion: "1.19",
      localizationInfo: {
        defaultLanguageTag: "en",
        additionalLanguages: [
          {
            languageTag: "de",
            file: "resources/de.json",
          },
        ],
      },
      composeExtensions: [
        {
          composeExtensionType: "apiBased",
          apiSpecificationFile: "resources/openai.yml",
          commands: [
            {
              id: "GET /repairs",
              apiResponseRenderingTemplateFile: "resources/repairs.json",
              title: "fake",
            },
          ],
          botId: "",
        },
      ],
      icons: {
        color: "resources/color.png",
        outline: "resources/outline.png",
      },
    } as TeamsManifest;
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));

    sinon.stub(fs, "chmod").callsFake(async () => {});
    sinon.stub(fs, "writeFile").callsFake(async () => {});

    delete process.env[openapiServerPlaceholder];
    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(
      result.isErr() &&
        result.error.name === "MissingEnvironmentVariablesError" &&
        result.error.message.includes(openapiServerPlaceholder)
    );
  });

  it("version > 1.6: happy path - CLI", async () => {
    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
    };

    const manifest = {
      manifestVersion: "1.19",
      localizationInfo: {
        defaultLanguageTag: "en",
        additionalLanguages: [
          {
            languageTag: "de",
            file: "resources/de.json",
          },
        ],
      },
    } as TeamsManifest;
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "resources/openai.yml",
        commands: [
          {
            id: "GET /repairs",
            apiResponseRenderingTemplateFile: "resources/repairs.json",
            title: "fake",
          },
        ],
        botId: "",
      },
    ];
    manifest.icons = {
      color: "resources/color.png",
      outline: "resources/outline.png",
    };
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));

    sinon.stub(fs, "chmod").callsFake(async () => {});
    sinon.stub(fs, "writeFile").callsFake(async () => {});

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isOk());
    if (await fs.pathExists(args.outputZipPath)) {
      await fs.remove(args.outputZipPath);
    }
  });

  it("happy path - relative path", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputJsonPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
    };

    const manifest = {
      manifestVersion: "1.19",
    } as TeamsManifestV1D19.TeamsManifestV1D19;
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "manifest.template.json",
        commands: [
          {
            id: "GET /repairs",
            apiResponseRenderingTemplateFile: "manifest.template.json",
            title: "fake",
          },
        ],
        botId: "",
      },
    ];
    manifest.icons = {
      color: "resources/color.png",
      outline: "resources/outline.png",
    };
    manifest.localizationInfo = {
      defaultLanguageTag: "en",
      additionalLanguages: [
        {
          languageTag: "de",
          file: "resources/de.json",
        },
      ],
    };
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));

    sinon.stub(fs, "chmod").callsFake(async () => {});
    sinon.stub(fs, "writeFile").callsFake(async () => {});

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
    if (await fs.pathExists(args.outputZipPath)) {
      await fs.remove(args.outputZipPath);
    }

    const executeResult = await teamsAppDriver.execute(args, mockedDriverContext);
    chai.assert.isTrue(executeResult.result.isOk());
  });

  it("happy path - no AC template", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputJsonPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
    };

    const manifest = {
      manifestVersion: "1.19",
    } as TeamsManifestV1D19.TeamsManifestV1D19;
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "manifest.template.json",
        commands: [
          {
            id: "GET /repairs",
            title: "fake",
          },
        ],
        botId: "",
      },
    ];
    manifest.icons = {
      color: "resources/color.png",
      outline: "resources/outline.png",
    };
    manifest.localizationInfo = {
      defaultLanguageTag: "en",
      additionalLanguages: [
        {
          languageTag: "de",
          file: "resources/de.json",
        },
      ],
    };
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));

    sinon.stub(fs, "chmod").callsFake(async () => {});
    sinon.stub(fs, "writeFile").callsFake(async () => {});

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
    if (await fs.pathExists(args.outputZipPath)) {
      await fs.remove(args.outputZipPath);
    }

    const executeResult = await teamsAppDriver.execute(args, mockedDriverContext);
    chai.assert.isTrue(executeResult.result.isOk());
  });

  it("version >= 1.9: happy path - API plugin", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
    };

    const manifest = {
      manifestVersion: "1.19",
    } as TeamsManifestV1D19.TeamsManifestV1D19;
    manifest.copilotAgents = {
      declarativeAgents: [
        {
          file: "resources/de.json",
          id: "dc1",
        },
      ],
    };
    manifest.icons = {
      color: "resources/color.png",
      outline: "resources/outline.png",
    };
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
    sinon.stub(fs, "chmod").callsFake(async () => {});
    const writeFileStub = sinon.stub(fs, "writeFile").callsFake(async () => {});

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    if (result.isErr()) {
      console.log(result.error);
    }
    chai.assert.isTrue(result.isOk());
    const outputExist = await fs.pathExists(args.outputZipPath);
    chai.assert.isTrue(outputExist);
    chai.assert.isTrue(writeFileStub.calledTwice);
    if (outputExist) {
      const zip = new AdmZip(args.outputZipPath);
      const openapiContent = "";
      let declarativeAgentsContent = "";

      const entries = zip.getEntries();
      entries.forEach((e) => {
        const name = e.entryName;

        if (name.endsWith("de.json")) {
          const data = e.getData();
          declarativeAgentsContent = data.toString("utf8");
        }
      });

      chai.assert(declarativeAgentsContent);
      await fs.remove(args.outputZipPath);
    }
  });

  it("happy path - Plugin file with underscore namespace", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
    };

    const manifest = {
      manifestVersion: "1.19",
    } as TeamsManifestV1D19.TeamsManifestV1D19;
    manifest.copilotAgents = {
      declarativeAgents: [
        {
          file: "resources/declarativeAgent-namespace.json",
          id: "dc1",
        },
      ],
    };
    manifest.icons = {
      color: "resources/color.png",
      outline: "resources/outline.png",
    };
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
    sinon.stub(fs, "chmod").callsFake(async () => {});
    const writeFileStub = sinon.stub(fs, "writeFile").callsFake(async () => {});

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    if (result.isErr()) {
      console.log(result.error);
    }
    chai.assert.isTrue(result.isOk());
    const outputExist = await fs.pathExists(args.outputZipPath);
    chai.assert.isTrue(outputExist);
    chai.assert.isTrue(writeFileStub.calledThrice);
    if (outputExist) {
      const zip = new AdmZip(args.outputZipPath);
      let aiPluginContent = "";
      let openapiContent = "";
      let declarativeAgentsContent = "";

      const entries = zip.getEntries();
      entries.forEach((e) => {
        const name = e.entryName;
        if (name.endsWith("ai-plugin-with-underscore-namespace.json")) {
          const data = e.getData();
          aiPluginContent = data.toString("utf8");
        }

        if (name.endsWith("openai.yml")) {
          const data = e.getData();
          openapiContent = data.toString("utf8");
        }

        if (name.endsWith("declarativeAgent-namespace.json")) {
          const data = e.getData();
          declarativeAgentsContent = data.toString("utf8");
        }
      });

      chai.assert(openapiContent, "openapi.yml not found in the zip file");
      chai.assert(aiPluginContent, "ai-plugin.json not found in the zip file");
      chai.assert(declarativeAgentsContent, "declarativeAgent.json not found in the zip file");
      chai.assert(
        aiPluginContent.search(openapiServerPlaceholder) < 0,
        "openapiServerPlaceholder not replaced"
      );
      chai.assert.include(aiPluginContent, "pluginnamespace", "plugin_namespace not replaced");
      chai.assert(openapiContent.search("APP_NAME_SUFFIX") < 0, "APP_NAME_SUFFIX not replaced");
      chai.assert(aiPluginContent.search("file") < 0, "file not replaced");

      await fs.remove(args.outputZipPath);
    }
  });

  it("Plugin file processed error when expandVariableWithFunction failed ", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
    };

    const manifest = {
      manifestVersion: "1.19",
    } as TeamsManifestV1D19.TeamsManifestV1D19;
    manifest.copilotAgents = {
      declarativeAgents: [
        {
          file: "resources/declarativeAgent-namespace.json",
          id: "dc1",
        },
      ],
    };
    manifest.icons = {
      color: "resources/color.png",
      outline: "resources/outline.png",
    };

    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
    sinon.stub(fs, "chmod").callsFake(async () => {});
    sinon.stub(fs, "writeFile").callsFake(async () => {});

    sinon
      .stub(envFunctionUtils, "expandVariableWithFunction")
      .callsFake(
        async (
          content: string,
          ctx: DriverContext,
          envs: { [key in string]: string } | undefined,
          isJson: boolean,
          manifestType: ManifestType,
          fromPath: string
        ) => {
          if (fromPath.endsWith("ai-plugin-with-underscore-namespace.json")) {
            return err(new UserError("source", "name", "message"));
          } else {
            return ok(content);
          }
        }
      );

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert.isTrue(result.isErr());

    await fs.remove(args.outputZipPath);
  });

  it("happy path - Declarative Agent with external adaptive cards", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
    };

    const manifest = {
      manifestVersion: "1.19",
    } as TeamsManifestV1D19.TeamsManifestV1D19;
    manifest.copilotAgents = {
      declarativeAgents: [
        {
          file: "resources/declarativeAgent.json",
          id: "dc1",
        },
      ],
    };
    manifest.icons = {
      color: "resources/color.png",
      outline: "resources/outline.png",
    };
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
    sinon.stub(fs, "chmod").callsFake(async () => {});
    const writeFileStub = sinon.stub(fs, "writeFile").callsFake(async () => {});

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    if (result.isErr()) {
      console.log(result.error);
    }
    chai.assert.isTrue(result.isOk());
    const outputExist = await fs.pathExists(args.outputZipPath);
    chai.assert.isTrue(outputExist);
    chai.assert.isTrue(writeFileStub.calledThrice);
    if (outputExist) {
      const zip = new AdmZip(args.outputZipPath);
      let aiPluginContent = "";
      let openapiContent = "";
      let declarativeAgentsContent = "";

      const entries = zip.getEntries();
      entries.forEach((e) => {
        const name = e.entryName;
        if (name.endsWith("ai-plugin-with-external-ac.json")) {
          const data = e.getData();
          aiPluginContent = data.toString("utf8");
        }

        if (name.endsWith("openai.yml")) {
          const data = e.getData();
          openapiContent = data.toString("utf8");
        }

        if (name.endsWith("declarativeAgent.json")) {
          const data = e.getData();
          declarativeAgentsContent = data.toString("utf8");
        }
      });

      chai.assert(openapiContent, "openapi.yml not found in the zip file");
      chai.assert(aiPluginContent, "ai-plugin.json not found in the zip file");
      chai.assert(declarativeAgentsContent, "declarativeAgent.json not found in the zip file");
      chai.assert(
        aiPluginContent.search(openapiServerPlaceholder) < 0,
        "openapiServerPlaceholder not replaced"
      );
      chai.assert(openapiContent.search("APP_NAME_SUFFIX") < 0, "APP_NAME_SUFFIX not replaced");
      chai.assert(aiPluginContent.search("file") < 0, "file not replaced");

      await fs.remove(args.outputZipPath);
    }
  });

  it("invalid color file", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputJsonPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
    };

    const manifest = {
      manifestVersion: "1.19",
    } as TeamsManifestV1D19.TeamsManifestV1D19;
    manifest.icons = {
      color: "../color.png",
      outline: "resources/outline.png",
    };
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
    sinon.stub(fs, "pathExists").callsFake(() => {
      return true;
    });
    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof InvalidFileOutsideOfTheDirectotryError);
    }
  });

  it("invalid outline file", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputJsonPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
    };

    const manifest = {
      manifestVersion: "1.19",
    } as TeamsManifestV1D19.TeamsManifestV1D19;
    manifest.icons = {
      color: "resources/color.png",
      outline: "../outline.png",
    };
    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
    sinon.stub(fs, "pathExists").callsFake((filePath) => {
      return true;
    });
    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof InvalidFileOutsideOfTheDirectotryError);
    }
  });

  it("invalid api spec file", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputJsonPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
    };

    const manifest = {
      manifestVersion: "1.19",
    } as TeamsManifestV1D19.TeamsManifestV1D19;
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "../openai.yml",
        commands: [
          {
            id: "GET /repairs",
            apiResponseRenderingTemplateFile: "resources/repairs.json",
            title: "fake",
          },
        ],
        botId: "",
      },
    ];
    manifest.icons = {
      color: "resources/color.png",
      outline: "resources/outline.png",
    };

    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
    sinon.stub(fs, "pathExists").callsFake((filePath) => {
      return true;
    });
    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof InvalidFileOutsideOfTheDirectotryError);
    }
  });

  it("invalid response template file", async () => {
    const args: CreateAppPackageArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      outputZipPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
      outputJsonPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
    };

    const manifest = {
      manifestVersion: "1.19",
    } as TeamsManifestV1D19.TeamsManifestV1D19;
    manifest.composeExtensions = [
      {
        composeExtensionType: "apiBased",
        apiSpecificationFile: "resources/openai.yml",
        commands: [
          {
            id: "GET /repairs",
            apiResponseRenderingTemplateFile: "../repairs.json",
            title: "fake",
          },
        ],
        botId: "",
      },
    ];
    manifest.icons = {
      color: "resources/color.png",
      outline: "resources/outline.png",
    };

    sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
    sinon.stub(fs, "pathExists").callsFake((filePath) => {
      return true;
    });
    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof InvalidFileOutsideOfTheDirectotryError);
    }
  });

  describe("copilotGpt", async () => {
    it("version <= 1.6: happy path ", async () => {
      // Enable builder API flag.
      sinon
        .stub(featureFlagManager, "getBooleanValue")
        .withArgs(FeatureFlags.BuilderAPIEnabled)
        .returns(true);
      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
        outputJsonPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
      };

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;
      manifest.copilotAgents = {
        declarativeAgents: [
          {
            file: "resources/gpt.json",
            id: "action_1",
          },
        ],
      };
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };
      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "chmod").callsFake(async () => {});
      const writeFileStub = sinon.stub(fs, "writeFile").callsFake(async () => {});

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      if (result.isErr()) {
        console.log(result.error);
      }
      chai.assert.isTrue(result.isOk());
      chai.assert.isTrue(writeFileStub.calledOnce);
      const outputExist = await fs.pathExists(args.outputZipPath);
      chai.assert.isTrue(outputExist);
      if (outputExist) {
        const zip = new AdmZip(args.outputZipPath);
        let gptManifestContent = "";
        let plugin = "";
        let apiSpec = "";

        const entries = zip.getEntries();
        entries.forEach((e) => {
          const name = e.entryName;
          if (name.endsWith("gpt.json")) {
            const data = e.getData();
            gptManifestContent = data.toString("utf8");
          } else if (name.endsWith("ai-plugin.json")) {
            const data = e.getData();
            plugin = data.toString("utf8");
          } else if (name.endsWith("openai.yml")) {
            const data = e.getData();
            apiSpec = data.toString("utf8");
          }
        });

        chai.assert(
          plugin &&
            apiSpec &&
            gptManifestContent &&
            gptManifestContent.search("APP_NAME_SUFFIX") < 0 &&
            gptManifestContent.search("test") > 0
        );
        await fs.remove(args.outputZipPath);
      }
    });

    it("version > 1.6: happy path ", async () => {
      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
        outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
      };

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;
      manifest.copilotAgents = {
        declarativeAgents: [
          {
            file: "resources/gpt.json",
            id: "action_1",
          },
        ],
      };
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };
      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "chmod").callsFake(async () => {});
      const writeFileStub = sinon.stub(fs, "writeFile").callsFake(async () => {});

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      if (result.isErr()) {
        console.log(result.error);
      }
      chai.assert.isTrue(result.isOk());
      chai.assert.isTrue(writeFileStub.calledThrice);
      const outputExist = await fs.pathExists(args.outputZipPath);
      chai.assert.isTrue(outputExist);
      if (outputExist) {
        const zip = new AdmZip(args.outputZipPath);
        let gptManifestContent = "";
        let plugin = "";
        let apiSpec = "";

        const entries = zip.getEntries();
        entries.forEach((e) => {
          const name = e.entryName;
          if (name.endsWith("gpt.json")) {
            const data = e.getData();
            gptManifestContent = data.toString("utf8");
          } else if (name.endsWith("ai-plugin.json")) {
            const data = e.getData();
            plugin = data.toString("utf8");
          } else if (name.endsWith("openai.yml")) {
            const data = e.getData();
            apiSpec = data.toString("utf8");
          }
        });

        chai.assert(
          plugin &&
            apiSpec &&
            gptManifestContent &&
            gptManifestContent.search("APP_NAME_SUFFIX") < 0 &&
            gptManifestContent.search("test") > 0
        );
        await fs.remove(args.outputZipPath);
      }
    });

    it("error if gpt manifest does not exist ", async () => {
      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
        outputJsonPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
      };

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;
      manifest.copilotAgents = {
        declarativeAgents: [
          {
            file: "resources/gpt.json",
            id: "action_1",
          },
        ],
      };
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };
      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "chmod").callsFake(async () => {});
      sinon.stub(fs, "writeFile").callsFake(async () => {});
      sinon.stub(fs, "pathExists").callsFake(async (path: string) => {
        if (path.endsWith("gpt.json")) {
          return false;
        } else {
          return true;
        }
      });

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;

      chai.assert.isTrue(result.isErr());

      if (result.isErr()) {
        chai.assert.isTrue(result.error instanceof FileNotFoundError);
      }
    });

    it("error if parse gpt manifest error ", async () => {
      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
        outputJsonPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
      };

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;

      manifest.copilotAgents = {
        declarativeAgents: [
          {
            file: "resources/gpt.json",
            id: "action_1",
          },
        ],
      };
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };
      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "chmod").callsFake(async () => {});
      sinon.stub(fs, "writeFile").callsFake(async () => {});
      sinon.stub(fs, "readFile").callsFake(async (file: fs.PathLike | number) => {
        if (file.toString().includes("gpt.json")) {
          return "" as any;
        } else {
          return JSON.stringify({});
        }
      });

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;

      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert.isTrue(result.error instanceof JSONSyntaxError);
      }
    });

    it("error when placeholder is not resolved in gpt manifest", async () => {
      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
        outputJsonPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
      };
      sinon.stub(fs, "pathExists").callsFake((filePath) => {
        return true;
      });

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };
      manifest.copilotAgents = {
        declarativeAgents: [
          {
            file: "resources/gpt.json",
            id: "action_1",
          },
        ],
      };
      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "chmod").callsFake(async () => {});
      sinon.stub(fs, "writeFile").callsFake(async () => {});

      delete process.env["APP_NAME_SUFFIX"];
      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;

      chai.assert(
        result.isErr() &&
          result.error.name === "MissingEnvironmentVariablesError" &&
          result.error.message.includes("APP_NAME_SUFFIX")
      );
    });

    it("error when add files for plugin failed", async () => {
      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
        outputJsonPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.dev.json",
      };

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;
      manifest.copilotAgents = {
        declarativeAgents: [
          {
            file: "resources/gpt.json",
            id: "action_1",
          },
        ],
      };
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };
      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "chmod").callsFake(async () => {});
      sinon.stub(fs, "writeFile").callsFake(async () => {});
      delete process.env[openapiServerPlaceholder];

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;

      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert(
          result.isErr() &&
            result.error.name === "MissingEnvironmentVariablesError" &&
            result.error.message.includes(openapiServerPlaceholder)
        );
      }
    });

    it("relative path error 1", async () => {
      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
        outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
      };

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;
      manifest.localizationInfo = {
        defaultLanguageTag: "en",
        additionalLanguages: [
          {
            languageTag: "de",
            file: "../migrate.manifest.json",
          },
        ],
        defaultLanguageFile: "resources/de.json",
      };
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };
      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "chmod").callsFake(async () => {});
      const writeFileStub = sinon.stub(fs, "writeFile").callsFake(async () => {});

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      if (result.isErr()) {
        chai.assert.isTrue(result.error instanceof InvalidFileOutsideOfTheDirectotryError);
      }
    });

    it("resolve additional localization file error", async () => {
      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
        outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
      };

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;
      manifest.localizationInfo = {
        defaultLanguageTag: "en",
        additionalLanguages: [
          {
            languageTag: "de",
            file: "migrate.manifest.json",
          },
        ],
        defaultLanguageFile: "de.json",
      };
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };
      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "chmod").callsFake(async () => {});
      sinon.stub(fs, "writeFile").callsFake(async () => {});
      sinon
        .stub(manifestUtils, "resolveLocFile")
        .resolves(err(new FileNotFoundError("teamsapp", "faked_loc_path")));

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      if (result.isErr()) {
        chai.assert.isTrue(result.error instanceof FileNotFoundError);
      }
    });

    it("resolve default localization file error", async () => {
      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
        outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
      };

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;
      manifest.localizationInfo = {
        defaultLanguageTag: "en",
        additionalLanguages: [
          {
            languageTag: "de",
            file: "migrate.manifest.json",
          },
        ],
        defaultLanguageFile: "de.json",
      };
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };
      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "chmod").callsFake(async () => {});
      sinon.stub(fs, "writeFile").callsFake(async () => {});
      sinon.stub(manifestUtils, "resolveLocFile").callsFake(async (path) => {
        if (path.includes("migrate.manifest.json")) {
          return ok("{}");
        } else {
          return err(new FileNotFoundError("teamsapp", "faked_loc_path"));
        }
      });

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      if (result.isErr()) {
        chai.assert.isTrue(result.error instanceof FileNotFoundError);
      }
    });

    it("relative path error 2", async () => {
      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
        outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
      };

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;
      manifest.localizationInfo = {
        defaultLanguageTag: "en",
        additionalLanguages: [
          {
            languageTag: "de",
            file: "resources/de.json",
          },
        ],
        defaultLanguageFile: "../migrate.manifest.json",
      };
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };
      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "chmod").callsFake(async () => {});
      const writeFileStub = sinon.stub(fs, "writeFile").callsFake(async () => {});

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      if (result.isErr()) {
        chai.assert.isTrue(result.error instanceof InvalidFileOutsideOfTheDirectotryError);
      }
    });

    it("zip same level dir", async () => {
      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.dev.zip",
        outputFolder: "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage",
      };

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;
      manifest.composeExtensions = [
        {
          composeExtensionType: "apiBased",
          apiSpecificationFile: "resources/openai.yml",
          commands: [
            {
              id: "GET /repairs",
              apiResponseRenderingTemplateFile: "resources/repairs.json",
              title: "fake",
            },
          ],
          botId: "",
        },
      ];
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };
      manifest.localizationInfo = {
        defaultLanguageTag: "en",
        additionalLanguages: [
          {
            languageTag: "de",
            file: "de.json",
          },
        ],
        defaultLanguageFile: "de.json",
      };
      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));

      sinon.stub(fs, "chmod").callsFake(async () => {});
      const writeFileStub = sinon.stub(fs, "writeFile").callsFake(async () => {});
      sinon.stub(manifestUtils, "resolveLocFile").resolves(ok("{}"));

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert(result.isOk());
      chai.assert(writeFileStub.calledOnce);
      if (await fs.pathExists(args.outputZipPath)) {
        const zip = new AdmZip(args.outputZipPath);

        let openapiContent = "";

        const entries = zip.getEntries();
        for (const e of entries) {
          const name = e.entryName;

          if (name.endsWith("openai.yml")) {
            const data = e.getData();
            openapiContent = data.toString("utf8");
            break;
          }
        }

        chai.assert(
          openapiContent != undefined &&
            openapiContent.length > 0 &&
            openapiContent.search(fakeUrl) >= 0 &&
            openapiContent.search(openapiServerPlaceholder) < 0
        );
        await fs.remove(args.outputZipPath);
      }
    });

    it("should add embedded knowledge files for Declarative Agent", async () => {
      sinon
        .stub(featureFlagManager, "getBooleanValue")
        .withArgs(FeatureFlags.EmbeddedKnowledgeEnabled)
        .returns(true);

      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.embedded.zip",
        outputJsonPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.embedded.json",
      };

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;
      manifest.copilotAgents = {
        declarativeAgents: [{ file: "resources/declarativeAgent.json", id: "1" }],
      };
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };

      // Updated gpt manifest stub with required properties.
      const declarativeAgentManifest = {
        name: "TestDeclarativeCopilot",
        description: "Test declarative copilot manifest",
        actions: [],
        capabilities: [
          {
            name: "EmbeddedKnowledge",
            files: [{ file: "EmbeddedKnowledge/knowledge.docx" }],
          },
        ],
      } as DeclarativeCopilotManifestSchema;

      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "chmod").callsFake(async () => {});
      sinon.stub(fs, "writeFile").callsFake(async () => {});
      sinon.stub(copilotGptManifestUtils, "getManifest").resolves(ok(declarativeAgentManifest));
      sinon.stub(fs, "pathExists").callsFake(async (filePath) => {
        // Return true for all required files including declarativeAgent.json, color/outline files and knowledge file.
        if (
          filePath.includes("knowledge.docx") ||
          filePath.includes("declarativeAgent.json") ||
          filePath.includes("color.png") ||
          filePath.includes("outline.png")
        ) {
          return true;
        }
        return true;
      });

      const mockedDriverContext: any = {
        m365TokenProvider: {},
        projectPath: "./",
        platform: 0,
        logProvider: { info: () => {} },
        ui: {},
        addTelemetryProperties: () => {},
      };

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert.isTrue(result.isOk());

      if (await fs.pathExists(args.outputZipPath)) {
        const AdmZip = require("adm-zip");
        const zip = new AdmZip(args.outputZipPath);
        const knowledgeEntry = zip.getEntry("EmbeddedKnowledge/knowledge.docx");
        chai.assert.exists(knowledgeEntry, "Embedded knowledge file should be added");
        await fs.remove(args.outputZipPath);
      }
    });

    it("should add embedded knowledge files for Declarative Agent of MetaOS", async () => {
      sinon
        .stub(featureFlagManager, "getBooleanValue")
        .withArgs(FeatureFlags.EmbeddedKnowledgeEnabled)
        .returns(true);

      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.embedded.zip",
        outputJsonPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.embedded.json",
      };

      const manifest = {
        manifestVersion: "devPreview",
      } as TeamsManifestVDevPreview.TeamsManifestVDevPreview;
      manifest.copilotAgents = {
        declarativeAgents: [{ file: "resources/declarativeAgent.json", id: "1" }],
      };
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };

      // Updated gpt manifest stub with required properties.
      const declarativeAgentManifest = {
        name: "TestDeclarativeCopilot",
        description: "Test declarative copilot manifest",
        actions: [],
        capabilities: [
          {
            name: "EmbeddedKnowledge",
            files: [{ file: "EmbeddedKnowledge/knowledge.docx" }],
          },
        ],
      } as DeclarativeCopilotManifestSchema;

      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "chmod").callsFake(async () => {});
      sinon.stub(fs, "writeFile").callsFake(async () => {});
      sinon.stub(copilotGptManifestUtils, "getManifest").resolves(ok(declarativeAgentManifest));
      sinon.stub(fs, "pathExists").callsFake(async (filePath) => {
        // Return true for all required files including declarativeAgent.json, color/outline files and knowledge file.
        if (
          filePath.includes("knowledge.docx") ||
          filePath.includes("declarativeAgent.json") ||
          filePath.includes("color.png") ||
          filePath.includes("outline.png")
        ) {
          return true;
        }
        return true;
      });

      const mockedDriverContext: any = {
        m365TokenProvider: {},
        projectPath: "./",
        platform: 0,
        logProvider: { info: () => {} },
        ui: {},
        addTelemetryProperties: () => {},
      };

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert.isTrue(result.isOk());

      if (await fs.pathExists(args.outputZipPath)) {
        const AdmZip = require("adm-zip");
        const zip = new AdmZip(args.outputZipPath);
        const knowledgeEntry = zip.getEntry("EmbeddedKnowledge/knowledge.docx");
        chai.assert.exists(knowledgeEntry, "Embedded knowledge file should be added");
        await fs.remove(args.outputZipPath);
      }
    });

    it("should skip if there is no embedded knowledge capability for Declarative Agent", async () => {
      sinon
        .stub(featureFlagManager, "getBooleanValue")
        .withArgs(FeatureFlags.EmbeddedKnowledgeEnabled)
        .returns(true);

      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.embedded.zip",
        outputJsonPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.embedded.json",
      };

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;
      manifest.copilotAgents = {
        declarativeAgents: [{ file: "resources/declarativeAgent.json", id: "1" }],
      };
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };

      // Updated gpt manifest stub with required properties.
      const declarativeAgentManifest = {
        name: "TestDeclarativeCopilot",
        description: "Test declarative copilot manifest",
        actions: [],
        capabilities: [
          {
            name: "WebSearch",
          },
        ],
      } as DeclarativeCopilotManifestSchema;

      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "chmod").callsFake(async () => {});
      sinon.stub(fs, "writeFile").callsFake(async () => {});
      sinon.stub(copilotGptManifestUtils, "getManifest").resolves(ok(declarativeAgentManifest));
      sinon.stub(fs, "pathExists").callsFake(async (filePath) => {
        // Return true for all required files including declarativeAgent.json, color/outline files.
        if (
          filePath.includes("declarativeAgent.json") ||
          filePath.includes("color.png") ||
          filePath.includes("outline.png")
        ) {
          return true;
        }
        return true;
      });

      const mockedDriverContext: any = {
        m365TokenProvider: {},
        projectPath: "./",
        platform: 0,
        logProvider: { info: () => {} },
        ui: {},
        addTelemetryProperties: () => {},
      };

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert.isTrue(result.isOk());

      if (await fs.pathExists(args.outputZipPath)) {
        await fs.remove(args.outputZipPath);
      }
    });

    it("should handle undefined embedded knowledge files for Declarative Agent", async () => {
      sinon
        .stub(featureFlagManager, "getBooleanValue")
        .withArgs(FeatureFlags.EmbeddedKnowledgeEnabled)
        .returns(true);

      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.embedded.zip",
        outputJsonPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.embedded.json",
      };

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;
      manifest.copilotAgents = {
        declarativeAgents: [{ file: "resources/declarativeAgent.json", id: "1" }],
      };
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };

      // Updated gpt manifest stub with required properties.
      const declarativeAgentManifest = {
        name: "TestDeclarativeCopilot",
        description: "Test declarative copilot manifest",
        actions: [],
        capabilities: [
          {
            name: "EmbeddedKnowledge",
            files: [{}],
          },
        ],
      } as DeclarativeCopilotManifestSchema;

      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "chmod").callsFake(async () => {});
      sinon.stub(fs, "writeFile").callsFake(async () => {});
      sinon.stub(copilotGptManifestUtils, "getManifest").resolves(ok(declarativeAgentManifest));
      sinon.stub(fs, "pathExists").callsFake(async (filePath) => {
        // Return true for all required files including declarativeAgent.json, color/outline files.
        if (
          filePath.includes("declarativeAgent.json") ||
          filePath.includes("color.png") ||
          filePath.includes("outline.png")
        ) {
          return true;
        }
        return true;
      });

      const mockedDriverContext: any = {
        m365TokenProvider: {},
        projectPath: "./",
        platform: 0,
        logProvider: { info: () => {} },
        ui: {},
        addTelemetryProperties: () => {},
      };

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert.isTrue(result.isOk());

      if (await fs.pathExists(args.outputZipPath)) {
        await fs.remove(args.outputZipPath);
      }
    });

    it("should throw error if embedded knowledge file does not exist for Declarative Agent", async () => {
      sinon
        .stub(featureFlagManager, "getBooleanValue")
        .withArgs(FeatureFlags.EmbeddedKnowledgeEnabled)
        .returns(true);
      const args: CreateAppPackageArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        outputZipPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/appPackage.embedded.missing.zip",
        outputJsonPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/build/appPackage/manifest.embedded.missing.json",
      };

      const manifest = {
        manifestVersion: "1.19",
      } as TeamsManifestV1D19.TeamsManifestV1D19;
      manifest.copilotAgents = {
        declarativeAgents: [{ file: "resources/declarativeAgent.json", id: "1" }],
      };
      manifest.icons = {
        color: "resources/color.png",
        outline: "resources/outline.png",
      };

      // Prepare a minimal declarative agent manifest with an embedded knowledge file.
      const declarativeAgentManifest = {
        name: "TestDeclarativeCopilot",
        description: "Missing knowledge file test",
        actions: [],
        capabilities: [
          {
            name: "EmbeddedKnowledge",
            files: [{ file: "EmbeddedKnowledge/knowledgeMissing.docx" }],
          },
        ],
      } as DeclarativeCopilotManifestSchema;

      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(manifest));
      sinon.stub(fs, "chmod").callsFake(async () => {});
      sinon.stub(fs, "writeFile").callsFake(async () => {});
      sinon.stub(copilotGptManifestUtils, "getManifest").resolves(ok(declarativeAgentManifest));
      // Simulate missing knowledge file.
      sinon.stub(fs, "pathExists").callsFake(async (filePath) => {
        if (filePath.toString().includes("knowledgeMissing.docx")) {
          return false;
        }
        return true;
      });

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert.isTrue(result.isErr());
      if (result.isErr()) {
        chai.assert.isTrue(result.error instanceof FileNotFoundError);
      }
    });
  });
});
