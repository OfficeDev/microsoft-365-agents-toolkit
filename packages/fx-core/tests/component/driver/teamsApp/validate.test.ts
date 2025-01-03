// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ManifestUtil,
  Platform,
  SystemError,
  TeamsAppManifest,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import chai from "chai";
import fs from "fs-extra";
import "mocha";
import * as sinon from "sinon";
import { teamsDevPortalClient } from "../../../../src/client/teamsDevPortalClient";
import { setTools } from "../../../../src/common/globalVars";
import * as commonTools from "../../../../src/common/utils";
import {
  Constants,
  GeneralValidationErrorId,
} from "../../../../src/component/driver/teamsApp/constants";
import { AppStudioError } from "../../../../src/component/driver/teamsApp/errors";
import {
  AsyncAppValidationResponse,
  AsyncAppValidationStatus,
} from "../../../../src/component/driver/teamsApp/interfaces/AsyncAppValidationResponse";
import { AsyncAppValidationResultsResponse } from "../../../../src/component/driver/teamsApp/interfaces/AsyncAppValidationResultsResponse";
import { ValidateAppPackageArgs } from "../../../../src/component/driver/teamsApp/interfaces/ValidateAppPackageArgs";
import { ValidateManifestArgs } from "../../../../src/component/driver/teamsApp/interfaces/ValidateManifestArgs";
import { ValidateWithTestCasesArgs } from "../../../../src/component/driver/teamsApp/interfaces/ValidateWithTestCasesArgs";
import { IAppValidationNote } from "../../../../src/component/driver/teamsApp/interfaces/appdefinitions/IValidationResult";
import { teamsappMgr } from "../../../../src/component/driver/teamsApp/teamsappMgr";
import { copilotGptManifestUtils } from "../../../../src/component/driver/teamsApp/utils/CopilotGptManifestUtils";
import { manifestUtils } from "../../../../src/component/driver/teamsApp/utils/ManifestUtils";
import { pluginManifestUtils } from "../../../../src/component/driver/teamsApp/utils/PluginManifestUtils";
import { ValidateManifestDriver } from "../../../../src/component/driver/teamsApp/validate";
import { ValidateAppPackageDriver } from "../../../../src/component/driver/teamsApp/validateAppPackage";
import { ValidateWithTestCasesDriver } from "../../../../src/component/driver/teamsApp/validateTestCases";
import { metadataUtil } from "../../../../src/component/utils/metadataUtil";
import { InvalidActionInputError, UserCancelError } from "../../../../src/error/common";
import { MockedM365Provider, MockTools } from "../../../core/utils";
import { MockedLogProvider, MockedUserInteraction } from "../../../plugins/solution/util";

describe("teamsApp/validateManifest", async () => {
  const teamsAppDriver = new ValidateManifestDriver();
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
    projectPath: "./",
  };

  afterEach(() => {
    sinon.restore();
  });

  it("file not found - manifest", async () => {
    const args: ValidateManifestArgs = {
      manifestPath: "fakepath",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(AppStudioError.FileNotFoundError.name, result.error.name);
    }
  });

  it("invalid param error", async () => {
    const args: ValidateManifestArgs = {
      manifestPath: "",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof InvalidActionInputError);
    }
  });

  it("happy path - validate against schema", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
    };

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("execute", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
    };

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = await teamsAppDriver.execute(args, mockedDriverContext);
    chai.assert(result.result.isOk());
  });

  it("happy path - VS", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
    };

    mockedDriverContext.platform = Platform.VS;

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("happy path- CLI", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      showMessage: true,
    };

    mockedDriverContext.platform = Platform.CLI;

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("happy path- VSC", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
      showMessage: true,
    };

    mockedDriverContext.platform = Platform.VSCode;

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("validation error - no schema", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.noSchema.manifest.json",
    };

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert(result.error.name, AppStudioError.ValidationFailedError.name);
    }
  });

  it("validation error - invalid", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.invalid.manifest.json",
    };

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert(result.error.name, AppStudioError.ValidationFailedError.name);
    }
  });

  it("validation error - cli", async () => {
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.invalid.manifest.json",
    };

    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert(result.error.name, AppStudioError.ValidationFailedError.name);
    }
  });

  it("validation error - download failed", async () => {
    sinon.stub(ManifestUtil, "validateManifest").resolves([]);
    sinon.stub(ManifestUtil, "validateManifestAgainstSchema").throws("error");
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
    };

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert(result.error.name, AppStudioError.ValidationFailedError.name);
    }
  });

  it("validation error - localization file validation failed", async () => {
    sinon
      .stub(ManifestUtil, "validateManifest")
      .throws(new Error(`Failed to get manifest at url due to: unknown error`));
    const args: ValidateManifestArgs = {
      manifestPath:
        "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
    };

    process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert(result.error.name, AppStudioError.ValidationFailedError.name);
    }
  });

  describe("validateLocalizatoinFiles", async () => {
    const teamsAppDriver = new ValidateManifestDriver();
    const mockedDriverContext: any = {
      projectPath: "./",
    };

    afterEach(() => {
      sinon.restore();
    });

    it("should return ok when no additionalLanguages in manifest", async () => {
      const args: ValidateManifestArgs = { manifestPath: "fakepath" };
      const manifest = { localizationInfo: { additionalLanguages: [] } } as any;

      const result = await teamsAppDriver.validateLocalizatoinFiles(
        args,
        mockedDriverContext,
        manifest
      );
      chai.assert(result.isOk());
    });

    it("should return error when language file path is not defined", async () => {
      const args: ValidateManifestArgs = { manifestPath: "fakepath" };
      const manifest = { localizationInfo: { additionalLanguages: [{ file: undefined }] } } as any;

      const result = await teamsAppDriver.validateLocalizatoinFiles(
        args,
        mockedDriverContext,
        manifest
      );
      chai.assert(result.isErr());
      if (result.isErr()) {
        chai.assert.equal(result.error.name, AppStudioError.ValidationFailedError.name);
      }
    });

    it("should return error when manifest file cannot be found", async () => {
      const args: ValidateManifestArgs = { manifestPath: "fakepath" };
      const manifest = { localizationInfo: { additionalLanguages: [{ file: "filePath" }] } } as any;

      sinon
        .stub(manifestUtils, "resolveLocFile")
        .resolves(err(new SystemError("error", "error", "", "")));

      const result = await teamsAppDriver.validateLocalizatoinFiles(
        args,
        mockedDriverContext,
        manifest
      );
      chai.assert(result.isErr());
      if (result.isErr()) {
        chai.assert.equal(result.error.name, "error");
      }
    });

    it("should return error when validation fails", async () => {
      const args: ValidateManifestArgs = { manifestPath: "fakepath" };
      const manifest = { localizationInfo: { additionalLanguages: [{ file: "filePath" }] } } as any;
      const fakeLocalizationFile = {
        $schema:
          "https://developer.microsoft.com/en-us/json-schemas/teams/v1.16/MicrosoftTeams.Localization.schema.json",
      };

      sinon
        .stub(manifestUtils, "resolveLocFile")
        .resolves(ok(JSON.stringify(fakeLocalizationFile)));
      sinon.stub(ManifestUtil, "validateManifestAgainstSchema").resolves(["Validation error"]);

      const result = await teamsAppDriver.validateLocalizatoinFiles(
        args,
        mockedDriverContext,
        manifest
      );
      chai.assert(result.isOk());
      if (result.isOk()) {
        chai.assert.isTrue(result.value.error[0].includes("Validation error"));
      }
    });

    it("should output errors when validation fails", async () => {
      const args: ValidateManifestArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.invalid.localization.manifest.json",
      };
      process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert(result.isErr());
      if (result.isErr()) {
        chai.assert.isTrue(result.error.message.includes("2 failed"));
      }
    });

    it("should output errors when validation fails - CLI", async () => {
      const args: ValidateManifestArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.invalid.localization.manifest.json",
      };
      const mockedCLIDriverContext: any = {
        m365TokenProvider: new MockedM365Provider(),
        logProvider: new MockedLogProvider(),
        ui: new MockedUserInteraction(),
        projectPath: "./",
        platform: Platform.CLI,
      };
      process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

      const result = (await teamsAppDriver.execute(args, mockedCLIDriverContext)).result;
      chai.assert(result.isErr());
      if (result.isErr()) {
        chai.assert.isTrue(result.error.message.includes("2 failed"));
      }
    });

    it("should output errors when default language file validation fails", async () => {
      const args: ValidateManifestArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.invalid.default.localization.manifest.json",
      };
      process.env.CONFIG_TEAMS_APP_NAME = "fakeName";

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert(result.isErr());
      if (result.isErr()) {
        chai.assert.isTrue(result.error.message.includes("2 failed"));
      }
    });

    it("should return error when validation throws exception", async () => {
      const args: ValidateManifestArgs = { manifestPath: "fakepath" };
      const manifest = { localizationInfo: { additionalLanguages: [{ file: "filePath" }] } } as any;
      const fakeLocalizationFile = {};

      sinon
        .stub(manifestUtils, "resolveLocFile")
        .resolves(ok(JSON.stringify(fakeLocalizationFile)));
      sinon
        .stub(ManifestUtil, "validateManifestAgainstSchema")
        .throws(new Error("validation exception"));

      const result = await teamsAppDriver.validateLocalizatoinFiles(
        args,
        mockedDriverContext,
        manifest
      );
      chai.assert(result.isErr());
      if (result.isErr()) {
        chai.assert.equal(result.error.name, AppStudioError.ValidationFailedError.name);
      }
    });

    it("should not throw error if schema does not have patternProperties", async () => {
      const args: ValidateManifestArgs = { manifestPath: "fakepath" };
      const manifest = { localizationInfo: { additionalLanguages: [{ file: "filePath" }] } } as any;
      sinon.stub(ManifestUtil, "fetchSchema").resolves({} as any);
      sinon.stub(manifestUtils, "resolveLocFile").resolves(ok("{}"));
      sinon.stub(ManifestUtil, "validateManifestAgainstSchema").resolves([] as any);
      const result = await teamsAppDriver.validateLocalizatoinFiles(
        args,
        mockedDriverContext,
        manifest
      );
      chai.assert(result.isOk());
    });

    it("should return ok when localization file is valid", async () => {
      const args: ValidateManifestArgs = { manifestPath: "fakepath" };
      const manifest = { localizationInfo: { additionalLanguages: [{ file: "filePath" }] } } as any;
      const fakeLocalizationFile = {
        $schema:
          "https://developer.microsoft.com/en-us/json-schemas/teams/v1.16/MicrosoftTeams.Localization.schema.json",
        "name.short": "name short",
        "name.full": "name full",
        "description.short": "desp short",
        "description.full": "desp full",
        "staticTabs[0].name": "static tab name",
        "activities.activityTypes[0].description": "aa",
      };

      sinon
        .stub(manifestUtils, "resolveLocFile")
        .resolves(ok(JSON.stringify(fakeLocalizationFile)));
      const result = await teamsAppDriver.validateLocalizatoinFiles(
        args,
        mockedDriverContext,
        manifest
      );
      chai.assert(result.isOk());
    });
  });

  describe("validate Copilot extensions", async () => {
    it("validate with errors returned", async () => {
      const teamsManifest: TeamsAppManifest = new TeamsAppManifest();
      teamsManifest.copilotExtensions = {
        declarativeCopilots: [
          {
            id: "fakeId",
            file: "fakeFile",
          },
        ],
        plugins: [
          {
            id: "fakeId",
            file: "fakeFile",
          },
        ],
      };

      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(teamsManifest));
      sinon.stub(ManifestUtil, "validateManifest").resolves([]);
      sinon.stub(pluginManifestUtils, "validateAgainstSchema").resolves(
        ok({
          id: "fakeId",
          filePath: "fakeFile",
          validationResult: ["error1"],
        })
      );
      sinon.stub(pluginManifestUtils, "logValidationErrors").returns("errorMessage1");

      sinon.stub(copilotGptManifestUtils, "validateAgainstSchema").resolves(
        ok({
          id: "fakeId",
          filePath: "fakeFile",
          validationResult: ["error2"],
          actionValidationResult: [
            {
              id: "fakeId",
              filePath: "fakeFile",
              validationResult: ["error3"],
            },
          ],
        })
      );
      sinon.stub(copilotGptManifestUtils, "logValidationErrors").returns("errorMessage2");

      const args: ValidateManifestArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        showMessage: true,
      };

      mockedDriverContext.platform = Platform.VSCode;
      mockedDriverContext.projectPath = "test";

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert(result.isErr());
      if (result.isErr()) {
        chai.assert.equal(result.error.name, AppStudioError.ValidationFailedError.name);
      }
    });

    it("validate with errors returned - copilot agent", async () => {
      const teamsManifest: TeamsAppManifest = new TeamsAppManifest();
      teamsManifest.copilotAgents = {
        declarativeAgents: [
          {
            id: "fakeId",
            file: "fakeFile",
          },
        ],
        plugins: [
          {
            id: "fakeId",
            file: "fakeFile",
          },
        ],
      };

      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(teamsManifest));
      sinon.stub(ManifestUtil, "validateManifest").resolves([]);
      sinon.stub(pluginManifestUtils, "validateAgainstSchema").resolves(
        ok({
          id: "fakeId",
          filePath: "fakeFile",
          validationResult: ["error1"],
        })
      );
      sinon.stub(pluginManifestUtils, "logValidationErrors").returns("errorMessage1");

      sinon.stub(copilotGptManifestUtils, "validateAgainstSchema").resolves(
        ok({
          id: "fakeId",
          filePath: "fakeFile",
          validationResult: ["error2"],
          actionValidationResult: [
            {
              id: "fakeId",
              filePath: "fakeFile",
              validationResult: ["error3"],
            },
          ],
        })
      );
      sinon.stub(copilotGptManifestUtils, "logValidationErrors").returns("errorMessage2");

      const args: ValidateManifestArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        showMessage: true,
      };

      mockedDriverContext.platform = Platform.VSCode;
      mockedDriverContext.projectPath = "test";

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert(result.isErr());
      if (result.isErr()) {
        chai.assert.equal(result.error.name, AppStudioError.ValidationFailedError.name);
      }
    });

    it("skip plugin validation", async () => {
      const teamsManifest: TeamsAppManifest = new TeamsAppManifest();
      teamsManifest.copilotAgents = {};

      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(teamsManifest));
      sinon.stub(ManifestUtil, "validateManifest").resolves([]);
      sinon.stub(pluginManifestUtils, "validateAgainstSchema").resolves(
        ok({
          id: "fakeId",
          filePath: "fakeFile",
          validationResult: ["error1"],
        })
      );
      sinon.stub(pluginManifestUtils, "logValidationErrors").returns("errorMessage1");

      sinon.stub(copilotGptManifestUtils, "validateAgainstSchema").resolves(
        ok({
          id: "fakeId",
          filePath: "fakeFile",
          validationResult: ["error2"],
          actionValidationResult: [
            {
              id: "fakeId",
              filePath: "fakeFile",
              validationResult: ["error3"],
            },
          ],
        })
      );
      sinon.stub(copilotGptManifestUtils, "logValidationErrors").returns("errorMessage2");

      const args: ValidateManifestArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        showMessage: true,
      };

      mockedDriverContext.platform = Platform.VSCode;
      mockedDriverContext.projectPath = "test";

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert(result.isOk());
    });
    it("plugin manifest validation error", async () => {
      const teamsManifest: TeamsAppManifest = new TeamsAppManifest();
      teamsManifest.copilotExtensions = {
        declarativeCopilots: [
          {
            id: "fakeId",
            file: "fakeFile",
          },
        ],
        plugins: [
          {
            id: "fakeId",
            file: "fakeFile",
          },
        ],
      };

      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(teamsManifest));
      sinon.stub(ManifestUtil, "validateManifest").resolves([]);
      sinon
        .stub(pluginManifestUtils, "validateAgainstSchema")
        .resolves(err(new SystemError("testError", "testError", "", "")));
      sinon.stub(pluginManifestUtils, "logValidationErrors").returns("errorMessage1");

      sinon.stub(copilotGptManifestUtils, "validateAgainstSchema").resolves(
        ok({
          id: "fakeId",
          filePath: "fakeFile",
          validationResult: ["error2"],
          actionValidationResult: [
            {
              id: "fakeId",
              filePath: "fakeFile",
              validationResult: ["error3"],
            },
          ],
        })
      );
      sinon.stub(copilotGptManifestUtils, "logValidationErrors").returns("errorMessage2");

      const args: ValidateManifestArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        showMessage: true,
      };

      mockedDriverContext.platform = Platform.VSCode;
      mockedDriverContext.projectPath = "test";

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert(result.isErr());
      if (result.isErr()) {
        chai.assert.equal(result.error.name, "testError");
      }
    });

    it("declarative copilot manifest validation error", async () => {
      const teamsManifest: TeamsAppManifest = new TeamsAppManifest();
      teamsManifest.copilotExtensions = {
        declarativeCopilots: [
          {
            id: "fakeId",
            file: "fakeFile",
          },
        ],
        plugins: [
          {
            id: "fakeId",
            file: "fakeFile",
          },
        ],
      };

      sinon.stub(manifestUtils, "getManifestV3").resolves(ok(teamsManifest));
      sinon.stub(ManifestUtil, "validateManifest").resolves([]);
      sinon.stub(pluginManifestUtils, "validateAgainstSchema").resolves(
        ok({
          id: "fakeId",
          filePath: "fakeFile",
          validationResult: ["error1"],
        })
      );
      sinon.stub(pluginManifestUtils, "logValidationErrors").returns("errorMessage1");

      sinon
        .stub(copilotGptManifestUtils, "validateAgainstSchema")
        .resolves(err(new SystemError("testError", "testError", "", "")));
      sinon.stub(copilotGptManifestUtils, "logValidationErrors").returns("errorMessage2");

      const args: ValidateManifestArgs = {
        manifestPath:
          "./tests/plugins/resource/appstudio/resources-multi-env/templates/appPackage/v3.manifest.template.json",
        showMessage: true,
      };

      mockedDriverContext.platform = Platform.VSCode;
      mockedDriverContext.projectPath = "test";

      const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
      chai.assert(result.isErr());
      if (result.isErr()) {
        chai.assert.equal(result.error.name, "testError");
      }
    });
  });
});

describe("teamsApp/validateAppPackage", async () => {
  const teamsAppDriver = new ValidateAppPackageDriver();
  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
    projectPath: "./",
  };
  const contextWithoutUI: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    projectPath: "./",
  };

  afterEach(() => {
    sinon.restore();
    (mockedDriverContext.logProvider as MockedLogProvider).msg = "";
  });

  it("file not found - app package", async () => {
    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakepath",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(AppStudioError.FileNotFoundError.name, result.error.name);
    }
  });

  it("validate app package - error", async () => {
    sinon.stub(teamsDevPortalClient, "partnerCenterAppPackageValidation").resolves({
      errors: [
        {
          id: "fakeId",
          content: "Reserved Tab Name property should not be specified.",
          filePath: "",
          helpUrl: "https://docs.microsoft.com",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
        {
          id: "fakeId",
          content: "Reserved Tab Name property should not be specified.",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
        {
          id: GeneralValidationErrorId,
          content: "content",
          code: "Invalid TypeB Plugin document",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
        {
          id: GeneralValidationErrorId,
          content: "content",
          code: "Invalid DC document",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
        {
          id: GeneralValidationErrorId,
          content: "content with code missing",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
      ],
      status: "Rejected",
      warnings: [
        {
          id: "fakeId",
          content: "Valid domains cannot contain a hosting site with a wildcard.",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "domain",
          title: "valid domain",
        },
        {
          id: "fakeId",
          content: "Valid domains cannot contain a hosting site with a wildcard.",
          filePath: "",
          helpUrl: "https://docs.microsoft.com",
          shortCodeNumber: 123,
          validationCategory: "domain",
          title: "valid domain",
        },
      ],
      notes: [
        {
          id: "fakeId",
          content: "Schema URL is present.",
          title: "schema",
        },
      ],
      addInDetails: {
        displayName: "fake name",
        developerName: "fake name",
        version: "1.14.1",
        manifestVersion: "1.14.1",
      },
    });
    sinon.stub(fs, "pathExists").resolves(true);
    // sinon.stub(fs, "readFile").resolves(Buffer.from(""));
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakePath",
      showMessage: true,
    };
    let result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());

    result = (await teamsAppDriver.execute(args, contextWithoutUI)).result;
    chai.assert(result.isErr());

    const msg = (mockedDriverContext.logProvider as MockedLogProvider).msg;
    chai.assert(
      msg.includes("Invalid API Plugin document") &&
        msg.includes("Invalid DC document") &&
        msg.includes("content with code missing")
    );
  });

  it("validate app package - no error", async () => {
    sinon.stub(teamsDevPortalClient, "partnerCenterAppPackageValidation").resolves({
      errors: [],
      status: "Accepted",
      warnings: [],
      notes: [
        {
          id: "fakeId",
          content: "Schema URL is present.",
          title: "schema",
        },
        {
          id: "632652a7-0cf8-43c7-a65d-6a19e5822467",
          title: "Manifest Version is valid",
          code: "The app is using manifest version '1.17'",
        } as any as IAppValidationNote,
      ],
      addInDetails: {
        displayName: "fake name",
        developerName: "fake name",
        version: "1.14.1",
        manifestVersion: "1.14.1",
      },
    });
    sinon.stub(fs, "pathExists").resolves(true);
    // sinon.stub(fs, "readFile").resolves(Buffer.from(""));
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakePath",
      showMessage: true,
    };
    let result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());

    result = (await teamsAppDriver.execute(args, contextWithoutUI)).result;
    chai.assert(result.isOk());
  });

  it("validate app package - stop-on-error", async () => {
    sinon.stub(teamsDevPortalClient, "partnerCenterAppPackageValidation").resolves({
      errors: [
        {
          id: "fakeId",
          content: "Reserved Tab Name property should not be specified.",
          filePath: "",
          helpUrl: "https://docs.microsoft.com",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
      ],
      status: "Rejected",
      warnings: [],
      notes: [],
      addInDetails: {
        displayName: "fake name",
        developerName: "fake name",
        version: "1.14.1",
        manifestVersion: "1.14.1",
      },
    });
    sinon.stub(fs, "pathExists").resolves(true);
    // sinon.stub(fs, "readFile").resolves(Buffer.from(""));
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakePath",
      showMessage: false,
    };
    let result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());

    result = (await teamsAppDriver.execute(args, contextWithoutUI)).result;
    chai.assert(result.isErr());
  });

  it("errors - cli", async () => {
    sinon.stub(teamsDevPortalClient, "partnerCenterAppPackageValidation").resolves({
      errors: [
        {
          id: "fakeId",
          content: "Reserved Tab Name property should not be specified.",
          filePath: "",
          helpUrl: "https://docs.microsoft.com",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
        {
          id: "fakeId",
          content: "Reserved Tab Name property should not be specified.",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
        {
          id: GeneralValidationErrorId,
          content: "content",
          code: "Invalid TypeB Plugin document",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
        {
          id: GeneralValidationErrorId,
          content: "content",
          code: "Invalid DC document",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
        {
          id: GeneralValidationErrorId,
          content: "content with code missing",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
      ],
      status: "Rejected",
      warnings: [
        {
          id: "fakeId",
          content: "Valid domains cannot contain a hosting site with a wildcard.",
          filePath: "",
          helpUrl: "https://docs.microsoft.com",
          shortCodeNumber: 123,
          validationCategory: "domain",
          title: "valid domain",
        },
        {
          id: "fakeId",
          content: "Valid domains cannot contain a hosting site with a wildcard.",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "domain",
          title: "valid domain",
        },
      ],
      notes: [
        {
          id: "fakeId",
          content: "Schema URL is present.",
          title: "schema",
        },
      ],
      addInDetails: {
        displayName: "fake name",
        developerName: "fake name",
        version: "1.14.1",
        manifestVersion: "1.14.1",
      },
    });
    sinon.stub(fs, "pathExists").resolves(true);
    // sinon.stub(fs, "readFile").resolves(Buffer.from(""));
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakePath",
    };

    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isErr());
  });

  it("validation with only errors - cli", async () => {
    sinon.stub(teamsDevPortalClient, "partnerCenterAppPackageValidation").resolves({
      errors: [
        {
          id: "fakeId",
          content: "Reserved Tab Name property should not be specified.",
          filePath: "",
          helpUrl: "https://docs.microsoft.com",
          shortCodeNumber: 123,
          validationCategory: "tab",
          title: "tab name",
        },
      ],
      status: "Rejected",
      warnings: [],
      notes: [
        {
          id: "fakeId",
          content: "Schema URL is present.",
          title: "schema",
        },
      ],
      addInDetails: {
        displayName: "fake name",
        developerName: "fake name",
        version: "1.14.1",
        manifestVersion: "1.14.1",
      },
    });
    sinon.stub(fs, "pathExists").resolves(true);
    // sinon.stub(fs, "readFile").resolves(Buffer.from(""));
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakePath",
    };

    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isErr());
  });

  it("validation with warnings - cli", async () => {
    sinon.stub(teamsDevPortalClient, "partnerCenterAppPackageValidation").resolves({
      errors: [],
      status: "Rejected",
      warnings: [
        {
          id: "fakeId",
          content: "Valid domains cannot contain a hosting site with a wildcard.",
          filePath: "",
          helpUrl: "https://docs.microsoft.com",
          shortCodeNumber: 123,
          validationCategory: "domain",
          title: "valid domain",
        },
        {
          id: "fakeId",
          content: "Valid domains cannot contain a hosting site with a wildcard.",
          filePath: "",
          shortCodeNumber: 123,
          validationCategory: "domain",
          title: "valid domain",
        },
      ],
      notes: [
        {
          id: "fakeId",
          content: "Schema URL is present.",
          title: "schema",
        },
      ],
      addInDetails: {
        displayName: "fake name",
        developerName: "fake name",
        version: "1.14.1",
        manifestVersion: "1.14.1",
      },
    });
    sinon.stub(fs, "pathExists").resolves(true);
    // sinon.stub(fs, "readFile").resolves(Buffer.from(""));
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakePath",
    };

    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("happy path - cli", async () => {
    sinon.stub(teamsDevPortalClient, "partnerCenterAppPackageValidation").resolves({
      errors: [],
      status: "Rejected",
      warnings: [],
      notes: [
        {
          id: "fakeId",
          content: "Schema URL is present.",
          title: "schema",
        },
        {
          id: "632652a7-0cf8-43c7-a65d-6a19e5822467",
          title: "Manifest Version is valid",
          code: "The app is using manifest version '1.17'",
        } as any as IAppValidationNote,
      ],
      addInDetails: {
        displayName: "fake name",
        developerName: "fake name",
        version: "1.14.1",
        manifestVersion: "1.14.1",
      },
    });
    sinon.stub(fs, "pathExists").resolves(true);
    // sinon.stub(fs, "readFile").resolves(Buffer.from(""));
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      zip.addFile("color.png", new Buffer(""));
      zip.addFile("outlie.png", new Buffer(""));

      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    const args: ValidateAppPackageArgs = {
      appPackagePath: "fakePath",
    };

    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isOk());
  });
});

describe("teamsApp/validateWithTestCases", async () => {
  const tools = new MockTools();
  setTools(tools);

  const teamsAppDriver = new ValidateWithTestCasesDriver();

  const mockedDriverContext: any = {
    m365TokenProvider: new MockedM365Provider(),
    logProvider: new MockedLogProvider(),
    ui: new MockedUserInteraction(),
    projectPath: "./",
  };

  beforeEach(() => {
    sinon.stub(commonTools, "waitSeconds").resolves();
  });

  afterEach(() => {
    sinon.restore();
  });

  it("file not found - app package", async () => {
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(AppStudioError.FileNotFoundError.name, result.error.name);
    }
  });

  it("file not found - manifest.json", async () => {
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.equal(AppStudioError.FileNotFoundError.name, result.error.name);
    }
  });

  it("invalid param error", async () => {
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "",
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
    if (result.isErr()) {
      chai.assert.isTrue(result.error instanceof InvalidActionInputError);
    }
  });

  it("Failed to get token", async () => {
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakePath",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");
    sinon
      .stub(mockedDriverContext.m365TokenProvider, "getAccessToken")
      .resolves(err(new SystemError({})));

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isErr());
  });

  it("Invalid validation result response - Null details", async () => {
    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves(undefined);
    const mockSubmitValidationResponse: AsyncAppValidationResponse = {
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    };
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const invalidValidationResultResponseJson: any = {
      appValidationId: "appValidationId123",
      appId: "appId123",
      status: "Completed",
      appVersion: "1.0.0",
      manifestVersion: "1.0.0",
      createdAt: "2024-03-27T12:00:00.000Z",
      updatedAt: "2024-03-27T12:00:00.000Z",
      validationResults: {
        successes: null,
        warnings: null,
        failures: null,
        skipped: null,
      },
    };
    const invalidValidationResultResponse: AsyncAppValidationResultsResponse = <
      AsyncAppValidationResultsResponse
    >invalidValidationResultResponseJson;
    sinon
      .stub(teamsDevPortalClient, "getAppValidationById")
      .resolves(invalidValidationResultResponse);
    await teamsAppDriver.runningBackgroundJob(
      args,
      mockedDriverContext,
      "test_token",
      mockSubmitValidationResponse,
      "test_id"
    );
    chai.assert(
      mockedDriverContext.logProvider.msg.includes("Validation request completed, status:")
    );
  });

  it("Invalid validation result response - Null validation results", async () => {
    const mockSubmitValidationResponse: AsyncAppValidationResponse = {
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    };
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const invalidValidationResultResponseJson: any = {
      appValidationId: "appValidationId123",
      appId: "appId123",
      status: "Completed",
      appVersion: "1.0.0",
      manifestVersion: "1.0.0",
      createdAt: "2024-03-27T12:00:00.000Z",
      updatedAt: "2024-03-27T12:00:00.000Z",
      validationResults: null,
    };
    const invalidValidationResultResponse: AsyncAppValidationResultsResponse = <
      AsyncAppValidationResultsResponse
    >invalidValidationResultResponseJson;
    sinon
      .stub(teamsDevPortalClient, "getAppValidationById")
      .resolves(invalidValidationResultResponse);
    await teamsAppDriver.runningBackgroundJob(
      args,
      mockedDriverContext,
      "test_token",
      mockSubmitValidationResponse,
      "test_id"
    );
    chai.assert(
      mockedDriverContext.logProvider.msg.includes("Validation request completed, status:")
    );
  });

  it("Valid validation result response", async () => {
    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({
      appValidations: [
        {
          id: "fakeId",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Completed,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
        {
          id: "fakeId2",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Aborted,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      ],
    });
    const mockSubmitValidationResponse: AsyncAppValidationResponse = {
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    };
    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };
    sinon.stub(teamsDevPortalClient, "getAppValidationById").resolves({
      status: AsyncAppValidationStatus.Completed,
      appValidationId: "fakeId",
      appId: "fakeAppId",
      appVersion: "1.0.0",
      manifestVersion: "1.17",
      validationResults: {
        successes: [
          {
            title: "Validation_Success_Example",
            message: "Success validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        warnings: [
          {
            title: "Validation_Warning_Example",
            message: "Warning validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        failures: [
          {
            title: "Validation_Failure_Example",
            message: "Failure validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        skipped: [
          {
            title: "Validation_Skipped_Example",
            message: "Skipped validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    });
    await teamsAppDriver.runningBackgroundJob(
      args,
      mockedDriverContext,
      "test_token",
      mockSubmitValidationResponse,
      "test_id"
    );
    chai.assert(
      mockedDriverContext.logProvider.msg.includes("Validation request completed, status:")
    );
    chai.assert(
      mockedDriverContext.logProvider.msg.includes("1 failed, 1 warning, 1 skipped, 1 passed")
    );
  });

  it("Duplicate validations - InProgress", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({
      appValidations: [
        {
          id: "fakeId",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Completed,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
        {
          id: "fakeId2",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.InProgress,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      ],
    });
    sinon.stub(teamsDevPortalClient, "submitAppValidationRequest").throws("should not be called");
    sinon.stub(teamsDevPortalClient, "getAppValidationById").throws("should not be called");

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("Duplicate validations - Created", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({
      appValidations: [
        {
          id: "fakeId",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Completed,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
        {
          id: "fakeId2",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Created,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      ],
    });
    sinon.stub(teamsDevPortalClient, "submitAppValidationRequest").throws("should not be called");
    sinon.stub(teamsDevPortalClient, "getAppValidationById").throws("should not be called");

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("Duplicate validations - CLI", async () => {
    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({
      appValidations: [
        {
          id: "fakeId",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Completed,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
        {
          id: "fakeId2",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.InProgress,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      ],
    });
    sinon.stub(teamsDevPortalClient, "submitAppValidationRequest").throws("should not be called");
    sinon.stub(teamsDevPortalClient, "getAppValidationById").throws("should not be called");

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("Invalid list validation response", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({});
    sinon.stub(teamsDevPortalClient, "submitAppValidationRequest").resolves({
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    });

    sinon.stub(teamsDevPortalClient, "getAppValidationById").resolves({
      status: AsyncAppValidationStatus.Completed,
      appValidationId: "fakeId",
      appId: "fakeAppId",
      appVersion: "1.0.0",
      manifestVersion: "1.17",
      validationResults: {
        successes: [
          {
            title: "Validation_Success_Example",
            message: "Success validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        warnings: [],
        failures: [],
        skipped: [],
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    });

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("Happy path", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({
      appValidations: [
        {
          id: "fakeId",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Completed,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
        {
          id: "fakeId2",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Aborted,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      ],
    });
    sinon.stub(teamsDevPortalClient, "submitAppValidationRequest").resolves({
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    });

    sinon.stub(teamsDevPortalClient, "getAppValidationById").resolves({
      status: AsyncAppValidationStatus.Completed,
      appValidationId: "fakeId",
      appId: "fakeAppId",
      appVersion: "1.0.0",
      manifestVersion: "1.17",
      validationResults: {
        successes: [
          {
            title: "Validation_Success_Example",
            message: "Success validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        warnings: [
          {
            title: "Validation_Warning_Example",
            message: "Warning validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        failures: [
          {
            title: "Validation_Failure_Example",
            message: "Failure validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
        skipped: [
          {
            title: "Validation_Skipped_Example",
            message: "Skipped validation example message.",
            artifacts: {
              filePath: "fakePath",
              docsUrl: "https://docs.microsoft.com",
              policyNumber: "123",
              policyLinkUrl: "https://docs.microsoft.com",
              recommendation: "fakeRecommendation",
            },
          },
        ],
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    });

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("Aborted", async () => {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({
      appValidations: [
        {
          id: "fakeId",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Completed,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
        {
          id: "fakeId2",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Aborted,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      ],
    });
    sinon.stub(teamsDevPortalClient, "submitAppValidationRequest").resolves({
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    });

    sinon.stub(teamsDevPortalClient, "getAppValidationById").resolves({
      status: AsyncAppValidationStatus.Aborted,
      appValidationId: "fakeId",
      appId: "fakeAppId",
      appVersion: "1.0.0",
      manifestVersion: "1.17",
      validationResults: {
        failures: [],
        warnings: [],
        successes: [],
        skipped: [],
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    });

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: false,
    };

    const result = (await teamsAppDriver.execute(args, mockedDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("Happy path - CLI", async () => {
    const mockedCliDriverContext = {
      ...mockedDriverContext,
      platform: Platform.CLI,
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").callsFake(async () => {
      const zip = new AdmZip();
      zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(new TeamsAppManifest())));
      const archivedFile = zip.toBuffer();
      return archivedFile;
    });
    sinon.stub(metadataUtil, "parseManifest");

    sinon.stub(teamsDevPortalClient, "getAppValidationRequestList").resolves({
      appValidations: [
        {
          id: "fakeId",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Completed,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
        {
          id: "fakeId2",
          appId: "fakeAppId",
          appVersion: "1.0.0",
          manifestVersion: "1.17",
          status: AsyncAppValidationStatus.Aborted,
          createdAt: new Date(),
          updatedAt: new Date(),
        },
      ],
    });
    sinon.stub(teamsDevPortalClient, "submitAppValidationRequest").resolves({
      status: AsyncAppValidationStatus.Created,
      appValidationId: "fakeId",
    });

    sinon.stub(teamsDevPortalClient, "getAppValidationById").resolves({
      status: AsyncAppValidationStatus.Completed,
      appValidationId: "fakeId",
      appId: "fakeAppId",
      appVersion: "1.0.0",
      manifestVersion: "1.17",
      validationResults: {
        failures: [],
        warnings: [],
        successes: [],
        skipped: [],
      },
      createdAt: new Date(),
      updatedAt: new Date(),
    });

    const args: ValidateWithTestCasesArgs = {
      appPackagePath: "fakepath",
      showMessage: true,
      showProgressBar: true,
    };

    const result = (await teamsAppDriver.execute(args, mockedCliDriverContext)).result;
    chai.assert(result.isOk());
  });

  it("CLI - succeed", async () => {
    sinon.stub(ValidateWithTestCasesDriver.prototype, "validate").resolves(ok(new Map()));
    const result = await teamsappMgr.validateTeamsApp({
      projectPath: "xxx",
      platform: Platform.CLI,
      "package-file": "xxx",
      "validate-method": "test-cases",
    });
    chai.assert(result.isOk());
  });

  it("CLI - failed", async () => {
    sinon
      .stub(ValidateWithTestCasesDriver.prototype, "validate")
      .resolves(err(new UserCancelError()));
    const result = await teamsappMgr.validateTeamsApp({
      projectPath: "xxx",
      platform: Platform.CLI,
      "package-file": "xxx",
      "validate-method": "test-cases",
    });
    chai.assert(result.isErr());
  });
});
