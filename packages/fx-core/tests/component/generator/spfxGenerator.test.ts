// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Context, err, Inputs, ok, Platform, Stage, SystemError } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import fs from "fs-extra";
import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import os from "os";
import * as path from "path";
import * as sinon from "sinon";
import * as uuid from "uuid";
import { createContext, setTools } from "../../../src/common/globalVars";
import { getLocalizedString } from "../../../src/common/localizeUtils";
import { cpUtils } from "../../../src/component/deps-checker/";
import { ManifestUtils } from "../../../src/component/driver/teamsApp/utils/ManifestUtils";
import { Generator } from "../../../src/component/generator/generator";
import { GeneratorChecker } from "../../../src/component/generator/spfx/depsChecker/generatorChecker";
import { YoChecker } from "../../../src/component/generator/spfx/depsChecker/yoChecker";
import {
  SPFxGenerator,
  SPFxGeneratorImport,
  SPFxGeneratorNew,
} from "../../../src/component/generator/spfx/spfxGenerator";
import { getShellOptionValue, Utils } from "../../../src/component/generator/spfx/utils/utils";
import { envUtil } from "../../../src/component/utils/envUtil";
import { FileNotFoundError, UserCancelError } from "../../../src/error";
import { QuestionNames, SPFxVersionOptionIds } from "../../../src/question";
import { MockTools } from "../../core/utils";
import { TeamsProjectTypeOptions } from "../../../src/question/scaffold/vsc/teamsProjectTypeNode";
import { ProjectTypeOptions } from "../../../src/question/scaffold/vsc/ProjectTypeOptions";
import { TabCapabilityOptions } from "../../../src/question/scaffold/vsc/CapabilityOptions";

describe("SPFxGenerator", function () {
  const testFolder = path.resolve("./tmp");
  let context: Context;
  let mockedEnvRestore: RestoreFn | undefined;

  beforeEach(async () => {
    const gtools = new MockTools();
    setTools(gtools);
    context = createContext();

    await fs.ensureDir(testFolder);
    sinon.stub(Utils, "configure");

    const manifestId = uuid.v4();
    sinon
      .stub(fs, "readFile")
      .resolves(
        new Buffer(
          `{"id": "${manifestId}", "preconfiguredEntries": [{"title": {"default": "helloworld"}}]}`
        )
      );
    sinon.stub(fs, "writeFile").resolves();
    sinon.stub(fs, "rename").resolves();
    sinon.stub(fs, "copyFile").resolves();
    sinon.stub(fs, "remove").resolves();
    sinon.stub(fs, "readJson").callsFake((directory: string) => {
      if (directory.includes("teams")) {
        return {
          $schema:
            "https://developer.microsoft.com/en-us/json-schemas/teams/v1.17/MicrosoftTeams.schema.json",
          manifestVersion: "1.17",
          id: "fakedId",
          name: {
            short: "thisisaverylongappnametotestifitwillbetruncated",
          },
          icons: {
            color: "color.png",
            outline: "outline.png",
          },
          staticTabs: [],
          configurableTabs: [],
        };
      } else if (directory.includes(".yo-rc.json")) {
        return {
          "@microsoft/generator-sharepoint": {
            solutionName: "fakedSolutionName",
            version: "1.17.4",
          },
        };
      } else {
        return { id: "fakedid", preconfiguredEntries: [{ title: { default: "helloworld" } }] };
      }
    });
    sinon.stub(fs, "ensureFile").resolves();
    sinon.stub(fs, "writeJSON").resolves();
    sinon.stub(fs, "ensureDir").resolves();
  });

  afterEach(async () => {
    sinon.restore();
    if (mockedEnvRestore) {
      mockedEnvRestore();
    }
    if (await fs.pathExists(testFolder)) {
      await fs.remove(testFolder);
    }
  });

  it("Both yeoman generator and template generator is called when scaffold SPFx project", async function () {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "spfxTestApp",
      "spfx-solution": "new",
    };
    const doYeomanScaffoldStub = sinon
      .stub(SPFxGenerator, "doYeomanScaffold" as any)
      .resolves(ok(undefined));
    const generateTemplateStub = sinon
      .stub(Generator, "generateTemplate" as any)
      .resolves(ok(undefined));
    sinon.stub(cpUtils, "executeCommand").resolves("succeed");

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
    chai.expect(doYeomanScaffoldStub.calledOnce).to.be.true;
    chai.expect(generateTemplateStub.calledOnce).to.be.true;
  });

  it("scaffold SPFx project without framework", async function () {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(Generator, "generateTemplate" as any).resolves(ok(undefined));
    sinon.stub(cpUtils, "executeCommand").resolves("succeed");
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      [QuestionNames.SPFxFramework]: "none",
      [QuestionNames.SPFxWebpartDesc]: "test",
      [QuestionNames.SPFxWebpartName]: "hello",
      "app-name": "spfxTestApp",
      "spfx-solution": "new",
    };
    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
  });

  it("scaffold SPFx project with react framework", async function () {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(cpUtils, "executeCommand").resolves("succeed");
    sinon.stub(Generator, "generateTemplate" as any).resolves(ok(undefined));
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      [QuestionNames.SPFxFramework]: "react",
      [QuestionNames.SPFxWebpartDesc]: "test",
      [QuestionNames.SPFxWebpartName]: "hello",
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxSolution]: "new",
    };
    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
  });

  it("scaffold SPFx project with minimal framework", async function () {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(cpUtils, "executeCommand").resolves("succeed");
    sinon.stub(Generator, "generateTemplate" as any).resolves(ok(undefined));
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      [QuestionNames.SPFxFramework]: "minimal",
      [QuestionNames.SPFxWebpartDesc]: "test",
      [QuestionNames.SPFxWebpartName]: "hello",
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxSolution]: "new",
    };
    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
  });

  it("scaffold SPFx project with extremely long webpart name", async function () {
    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(cpUtils, "executeCommand").resolves("succeed");
    sinon.stub(Generator, "generateTemplate" as any).resolves(ok(undefined));
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      [QuestionNames.SPFxFramework]: "minimal",
      [QuestionNames.SPFxWebpartDesc]: "test",
      [QuestionNames.SPFxWebpartName]:
        "extremelylongextremelylongextremelylongextremelylongspfxwebpartname",
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxSolution]: "new",
    };
    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
  });

  it("select to install locally but no need to install", async function () {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxInstallPackage]: SPFxVersionOptionIds.installLocally,
      [QuestionNames.SPFxSolution]: "new",
    };
    sinon.stub(YoChecker.prototype, "isLatestInstalled").resolves(true);
    sinon.stub(GeneratorChecker.prototype, "isLatestInstalled").resolves(true);
    sinon.stub(cpUtils, "executeCommand").resolves("succeed");

    const generateTemplateStub = sinon
      .stub(Generator, "generateTemplate" as any)
      .resolves(ok(undefined));

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);

    chai.expect(generateTemplateStub.calledOnce).to.be.true;
  });

  it("select to install locally and install only sp", async function () {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxInstallPackage]: SPFxVersionOptionIds.installLocally,
      [QuestionNames.SPFxSolution]: "new",
    };
    sinon.stub(YoChecker.prototype, "isLatestInstalled").resolves(true);
    sinon.stub(cpUtils, "executeCommand").resolves("succeed");
    sinon.stub(GeneratorChecker.prototype, "isLatestInstalled").resolves(false);
    const yoInstaller = sinon.stub(YoChecker.prototype, "ensureDependency").resolves(ok(true));
    const generatorInstaller = sinon
      .stub(GeneratorChecker.prototype, "ensureDependency")
      .resolves(ok(true));

    const generateTemplateStub = sinon
      .stub(Generator, "generateTemplate" as any)
      .resolves(ok(undefined));

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);

    chai.expect(generateTemplateStub.calledOnce).to.be.true;
    chai.expect(yoInstaller.calledOnce).to.be.false;
    chai.expect(generatorInstaller.calledOnce).to.be.true;
  });

  it("select to install locally and install only yo", async function () {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxInstallPackage]: SPFxVersionOptionIds.installLocally,
      [QuestionNames.SPFxSolution]: "new",
    };
    sinon.stub(YoChecker.prototype, "isLatestInstalled").resolves(false);
    sinon.stub(GeneratorChecker.prototype, "isLatestInstalled").resolves(true);
    sinon.stub(cpUtils, "executeCommand").resolves("succeed");
    const yoInstaller = sinon.stub(YoChecker.prototype, "ensureDependency").resolves(ok(true));
    const generatorInstaller = sinon
      .stub(GeneratorChecker.prototype, "ensureDependency")
      .resolves(ok(true));

    const generateTemplateStub = sinon
      .stub(Generator, "generateTemplate" as any)
      .resolves(ok(undefined));

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);

    chai.expect(generateTemplateStub.calledOnce).to.be.true;
    chai.expect(yoInstaller.calledOnce).to.be.true;
    chai.expect(generatorInstaller.calledOnce).to.be.false;
  });

  it("select to install locally and install sp error", async function () {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxInstallPackage]: SPFxVersionOptionIds.installLocally,
      [QuestionNames.SPFxSolution]: "new",
    };
    sinon.stub(YoChecker.prototype, "isLatestInstalled").resolves(true);
    sinon.stub(GeneratorChecker.prototype, "isLatestInstalled").resolves(false);
    sinon.stub(YoChecker.prototype, "ensureDependency").resolves(ok(true));
    sinon
      .stub(GeneratorChecker.prototype, "ensureDependency")
      .resolves(err(new SystemError("source", "name", "msg", "msg")));

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isErr()).to.eq(true);
    if (result.isErr()) {
      chai.expect(result.error.name).equal("LatestPackageInstallFailed");
    }
  });

  it("select to install locally and install yo error", async function () {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxInstallPackage]: SPFxVersionOptionIds.installLocally,
      [QuestionNames.SPFxSolution]: "new",
    };
    sinon.stub(YoChecker.prototype, "isLatestInstalled").resolves(false);
    sinon.stub(GeneratorChecker.prototype, "isLatestInstalled").resolves(true);
    sinon
      .stub(YoChecker.prototype, "ensureDependency")
      .resolves(err(new SystemError("source", "name", "msg", "msg")));

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isErr()).to.eq(true);
    if (result.isErr()) {
      chai.expect(result.error.name).equal("LatestPackageInstallFailed");
    }
  });

  it("Yeoman Generator scaffolding error", async function () {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxInstallPackage]: SPFxVersionOptionIds.installLocally,
      [QuestionNames.SPFxSolution]: "new",
    };
    sinon.stub(YoChecker.prototype, "isLatestInstalled").resolves(true);
    sinon.stub(GeneratorChecker.prototype, "isLatestInstalled").resolves(true);
    sinon.stub(cpUtils, "executeCommand").throws(new Error("errorMessage"));
    sinon.stub(Generator, "generateTemplate" as any).resolves(ok(undefined));

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isErr()).to.eq(true);
  });

  it("Yeoman Generator scaffolding error with unknown", async function () {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxInstallPackage]: SPFxVersionOptionIds.installLocally,
      [QuestionNames.SPFxSolution]: "new",
    };
    sinon.stub(YoChecker.prototype, "isLatestInstalled").resolves(false);
    sinon.stub(GeneratorChecker.prototype, "isLatestInstalled").resolves(true);
    sinon.stub(cpUtils, "executeCommand").throws(new Error("errorMessage"));
    sinon.stub(Generator, "generateTemplate" as any).resolves(ok(undefined));
    sinon.stub(YoChecker.prototype, "ensureDependency").throws(new Error("unknown"));

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isErr()).to.eq(true);
    if (result.isErr()) {
      chai.expect(result.error.name).equal("SPFxScaffoldError");
    }
  });

  it("install locally and use path", async function () {
    mockedEnvRestore = mockedEnv({
      PATH: undefined,
    });
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxInstallPackage]: SPFxVersionOptionIds.installLocally,
      [QuestionNames.SPFxSolution]: "new",
    };
    sinon.stub(YoChecker.prototype, "isLatestInstalled").resolves(true);
    sinon.stub(GeneratorChecker.prototype, "isLatestInstalled").resolves(true);
    sinon.stub(cpUtils, "executeCommand").resolves("succeed");

    const generateTemplateStub = sinon
      .stub(Generator, "generateTemplate" as any)
      .resolves(ok(undefined));

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);

    chai.expect(generateTemplateStub.calledOnce).to.be.true;
  });

  it("use global packages and use path", async function () {
    mockedEnvRestore = mockedEnv({
      PATH: undefined,
    });
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxInstallPackage]: SPFxVersionOptionIds.globalPackage,
      [QuestionNames.SPFxSolution]: "new",
      globalSpfxPackageVersion: "1.17.0",
    };
    sinon.stub(YoChecker.prototype, "isLatestInstalled").resolves(true);
    sinon.stub(GeneratorChecker.prototype, "isLatestInstalled").resolves(true);
    sinon.stub(cpUtils, "executeCommand").resolves("succeed");

    const generateTemplateStub = sinon
      .stub(Generator, "generateTemplate" as any)
      .resolves(ok(undefined));

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);

    chai.expect(generateTemplateStub.calledOnce).to.be.true;
  });

  it("No web part in imported SPFx solution", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: testFolder,
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxSolution]: "import",
      [QuestionNames.SPFxFolder]: "c:\\test",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readdir").resolves([]);
    sinon.stub(fs, "copy").resolves();

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isErr()).to.eq(true);
    if (result.isErr()) {
      chai.expect(result.error.name).to.eq("RetrieveSPFxInfoFailed");
    }
  });

  it("No valid web part manifest when import SPFx solution", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: testFolder,
      "app-name": "spfxTestApp",
      "spfx-solution": "import",
      "spfx-folder": "c:\\test",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readdir").callsFake((directory: any) => {
      if (directory === path.join("c:\\test", "teams")) {
        return ["1_color.png", "1_outline.png"] as any;
      } else if (directory === path.join("c:\\test", "src", "webparts")) {
        return ["helloworld", "second"] as any;
      } else {
        return [];
      }
    });
    sinon.stub(fs, "statSync").returns({
      isDirectory: () => {
        return true;
      },
    } as any);
    sinon.stub(fs, "copy").resolves();

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isErr()).to.eq(true);
    if (result.isErr()) {
      chai.expect(result.error.name).to.eq("FileNotFoundError");
    }
  });

  it("Copy existing SPFx solution failed when import SPFx solution", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: testFolder,
      "app-name": "spfxTestApp",
      "spfx-solution": "import",
      "spfx-folder": "c:\\test",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readdir").resolves([]);
    sinon.stub(fs, "copy").throwsException("Failed to copy");

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isErr()).to.eq(true);
    if (result.isErr()) {
      chai.expect(result.error.name).to.eq("CopyExistingSPFxSolutioinFailed");
    }
  });

  it("Update SPFx template failed when import SPFx solution", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: testFolder,
      "app-name": "spfxTestApp",
      "spfx-solution": "import",
      "spfx-folder": "c:\\test",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readdir").callsFake((directory: any) => {
      if (directory === path.join("c:\\test", "teams")) {
        return ["1_color.png", "1_outline.png"] as any;
      } else if (directory === path.join("c:\\test", "src", "webparts")) {
        return ["helloworld", "second"] as any;
      } else {
        return ["HelloWorldWebPart.manifest.json"] as any;
      }
    });
    sinon.stub(fs, "statSync").returns({
      isDirectory: () => {
        return true;
      },
    } as any);
    sinon.stub(fs, "copy").resolves();
    sinon.stub(Generator, "generateTemplate" as any).resolves(ok(undefined));
    sinon
      .stub(ManifestUtils.prototype, "_readAppManifest")
      .throwsException("Failed to read manifest");

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isErr()).to.eq(true);
    if (result.isErr()) {
      chai.expect(result.error.name).to.eq("UpdateSPFxTemplateFailed");
    }
  });

  it("Web part with invalid manifeset will not be imported", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: testFolder,
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxSolution]: "import",
      [QuestionNames.SPFxFolder]: "c:\\test",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readdir").callsFake((directory: any) => {
      if (directory === path.join("c:\\test", "teams")) {
        return ["1_color.png", "1_outline.png"] as any;
      } else if (directory === path.join("c:\\test", "src", "webparts")) {
        return ["helloworld", "second"] as any;
      } else if (directory === path.join("c:\\test", "src", "webparts", "helloworld")) {
        return ["HelloWorldWebPart.manifest.json"] as any;
      } else {
        return [] as any;
      }
    });
    sinon.stub(fs, "statSync").returns({
      isDirectory: () => {
        return true;
      },
    } as any);
    const generateTemplateStub = sinon
      .stub(Generator, "generateTemplate" as any)
      .resolves(ok(undefined));
    const fakedManifest = {
      name: { short: "thisisaverylongappnametotestifitwillbetruncated" },
      staticTabs: [{ name: "default" }],
    };
    const readAppManifestStub = sinon
      .stub(ManifestUtils.prototype, "_readAppManifest")
      .resolves(ok(fakedManifest as any));
    const writeAppManifestStub = sinon
      .stub(ManifestUtils.prototype, "_writeAppManifest")
      .resolves();
    const writeEnvStub = sinon.stub(envUtil, "writeEnv");
    sinon.stub(fs, "copy").resolves();

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
    chai.expect(fakedManifest.staticTabs.length).to.eq(1);
    chai.expect(generateTemplateStub.calledOnce).to.eq(true);
    chai.expect(writeEnvStub.calledOnce).to.eq(true);
    chai.expect(readAppManifestStub.calledTwice).to.eq(true);
    chai.expect(writeAppManifestStub.calledTwice).to.eq(true);
  });

  it("Generate template fail when import SPFx solution", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: testFolder,
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxSolution]: "import",
      [QuestionNames.SPFxFolder]: "c:\\test",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readdir").callsFake((directory: any) => {
      if (directory === path.join("c:\\test", "teams")) {
        return ["1_color.png", "1_outline.png"] as any;
      } else if (directory === path.join("c:\\test", "src", "webparts")) {
        return ["helloworld", "second"] as any;
      } else {
        return ["HelloWorldWebPart.manifest.json"] as any;
      }
    });
    sinon.stub(fs, "statSync").returns({
      isDirectory: () => {
        return true;
      },
    } as any);
    const generateTemplateStub = sinon
      .stub(Generator, "generateTemplate" as any)
      .resolves(err(undefined));
    sinon.stub(fs, "copy").resolves();

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isErr()).to.eq(true);
    chai.expect(generateTemplateStub.calledOnce).to.eq(true);
  });

  it("Teams manifest staticTabs is updated if imported SPFx solution has multiple web parts", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: testFolder,
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxSolution]: "import",
      [QuestionNames.SPFxFolder]: "c:\\test",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readdir").callsFake((directory: any) => {
      if (directory === path.join("c:\\test", "teams")) {
        return ["1_color.png", "1_outline.png"] as any;
      } else if (directory === path.join("c:\\test", "src", "webparts")) {
        return ["helloworld", "second"] as any;
      } else {
        return ["HelloWorldWebPart.manifest.json"] as any;
      }
    });
    sinon.stub(fs, "statSync").returns({
      isDirectory: () => {
        return true;
      },
    } as any);
    const generateTemplateStub = sinon
      .stub(Generator, "generateTemplate" as any)
      .resolves(ok(undefined));
    const fakedManifest = {
      name: { short: "thisisaverylongappnametotestifitwillbetruncated" },
      staticTabs: [{ name: "default" }],
    };
    const readAppManifestStub = sinon
      .stub(ManifestUtils.prototype, "_readAppManifest")
      .resolves(ok(fakedManifest as any));
    const writeAppManifestStub = sinon
      .stub(ManifestUtils.prototype, "_writeAppManifest")
      .resolves();
    const writeEnvStub = sinon.stub(envUtil, "writeEnv");
    sinon.stub(fs, "copy").resolves();

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
    chai.expect(fakedManifest.staticTabs.length).to.eq(3);
    chai.expect(generateTemplateStub.calledOnce).to.eq(true);
    chai.expect(writeEnvStub.calledOnce).to.eq(true);
    chai.expect(readAppManifestStub.calledTwice).to.eq(true);
    chai.expect(writeAppManifestStub.calledTwice).to.eq(true);
  });

  it("Teams manifest staticTabs is updated if imported SPFx solution has multiple web parts - SPFx higher than 1.21.0", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      projectPath: testFolder,
      [QuestionNames.AppName]: "spfxTestApp",
      [QuestionNames.SPFxSolution]: "import",
      [QuestionNames.SPFxFolder]: "c:\\test",
    };

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readdir").callsFake((directory: any) => {
      if (directory === path.join("c:\\test", "teams")) {
        return ["1_color.png", "1_outline.png"] as any;
      } else if (directory === path.join("c:\\test", "src", "webparts")) {
        return ["helloworld", "second"] as any;
      } else {
        return ["HelloWorldWebPart.manifest.json"] as any;
      }
    });
    sinon.stub(fs, "statSync").returns({
      isDirectory: () => {
        return true;
      },
    } as any);
    const generateTemplateStub = sinon
      .stub(Generator, "generateTemplate" as any)
      .resolves(ok(undefined));
    const fakedManifest = {
      name: { short: "thisisaverylongappnametotestifitwillbetruncated" },
      staticTabs: [{ name: "default" }],
    };
    const readAppManifestStub = sinon
      .stub(ManifestUtils.prototype, "_readAppManifest")
      .resolves(ok(fakedManifest as any));
    const writeAppManifestStub = sinon
      .stub(ManifestUtils.prototype, "_writeAppManifest")
      .resolves();
    const writeEnvStub = sinon.stub(envUtil, "writeEnv");
    sinon.stub(fs, "copy").resolves();
    sinon.stub(SPFxGenerator, "getSolutionVersion").resolves("1.21.0");
    if (context) {
      context.templateVariables = context.templateVariables || {};
      context.templateVariables["useNewDevUrl"] = "true";
    }

    const result = await SPFxGenerator.generate(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
    chai.expect(fakedManifest.staticTabs.length).to.eq(3);
    chai.expect(generateTemplateStub.calledOnce).to.eq(true);
    chai.expect(writeEnvStub.calledOnce).to.eq(true);
    chai.expect(readAppManifestStub.calledTwice).to.eq(true);
    chai.expect(writeAppManifestStub.calledTwice).to.eq(true);
  });

  describe("get node versions from SPFx package.json", async () => {
    it("found node version", async () => {
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readJSON").callsFake((directory: string) => {
        if (directory.includes("package.json")) {
          return { engines: { node: ">= 10.13.0 < 11.0.0" } };
        } else {
          return "";
        }
      });
      sinon.stub(Generator, "generateTemplate" as any).resolves(ok(undefined));
      sinon.stub(cpUtils, "executeCommand").resolves("succeed");
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.SPFxFramework]: "none",
        [QuestionNames.SPFxWebpartDesc]: "test",
        [QuestionNames.SPFxWebpartName]: "hello",
        "app-name": "spfxTestApp",
        "spfx-solution": "new",
      };
      const result = await SPFxGenerator.generate(context, inputs, testFolder);

      chai.expect(context.templateVariables!.SpfxNodeVersion).eq(">= 10.13.0 < 11.0.0");
      chai.expect(result.isOk()).to.eq(true);
    });

    it("cannot found engine", async () => {
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readJSON").callsFake((directory: string) => {
        if (directory.includes("package.json")) {
          return {};
        } else {
          return "";
        }
      });
      sinon.stub(Generator, "generateTemplate" as any).resolves(ok(undefined));
      sinon.stub(cpUtils, "executeCommand").resolves("succeed");
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.SPFxFramework]: "none",
        [QuestionNames.SPFxWebpartDesc]: "test",
        [QuestionNames.SPFxWebpartName]: "hello",
        "app-name": "spfxTestApp",
        "spfx-solution": "new",
      };
      const result = await SPFxGenerator.generate(context, inputs, testFolder);

      chai.expect(context.templateVariables!.SpfxNodeVersion).eq("18");
      chai.expect(result.isOk()).to.eq(true);
    });

    it("cannot found engines.node", async () => {
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readJSON").callsFake((directory: string) => {
        if (directory.includes("package.json")) {
          return { engines: {} };
        } else {
          return "";
        }
      });
      sinon.stub(Generator, "generateTemplate" as any).resolves(ok(undefined));
      sinon.stub(cpUtils, "executeCommand").resolves("succeed");
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.SPFxFramework]: "none",
        [QuestionNames.SPFxWebpartDesc]: "test",
        [QuestionNames.SPFxWebpartName]: "hello",
        "app-name": "spfxTestApp",
        "spfx-solution": "new",
      };
      const result = await SPFxGenerator.generate(context, inputs, testFolder);

      chai.expect(context.templateVariables!.SpfxNodeVersion).eq("18");
      chai.expect(result.isOk()).to.eq(true);
    });
  });

  describe("doYeomanScaffold: add web part", async () => {
    it("add web part with global package", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.AppName]: "spfxTestApp",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFolder]: "folder",
        [QuestionNames.SPFxWebpartName]: "hello",
        stage: Stage.addWebpart,
      };

      sinon.stub(GeneratorChecker.prototype, "findGloballyInstalledVersion").resolves("1.17.4");
      sinon.stub(GeneratorChecker.prototype, "findLocalInstalledVersion").resolves(undefined);
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(cpUtils, "executeCommand").resolves("succeed");
      const yoInstaller = sinon.stub(YoChecker.prototype, "ensureDependency").resolves(ok(true));
      const generatorInstaller = sinon
        .stub(GeneratorChecker.prototype, "ensureDependency")
        .resolves(ok(true));

      const result = await SPFxGenerator.doYeomanScaffold(context, inputs, testFolder);
      if (result.isErr()) {
        console.log(result.error);
      }

      chai.expect(result.isOk()).to.eq(true);

      chai.expect(yoInstaller.called).to.be.false;
      chai.expect(generatorInstaller.called).to.be.false;
    });

    it("add web part with global package - SPFx higher than 1.21", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFolder]: "folder",
        [QuestionNames.SPFxWebpartName]: "hello",
        stage: Stage.addWebpart,
      };

      sinon.stub(GeneratorChecker.prototype, "findGloballyInstalledVersion").resolves("1.21.0");
      sinon.stub(GeneratorChecker.prototype, "findLocalInstalledVersion").resolves(undefined);
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(cpUtils, "executeCommand").resolves("succeed");
      const yoInstaller = sinon.stub(YoChecker.prototype, "ensureDependency").resolves(ok(true));
      const generatorInstaller = sinon
        .stub(GeneratorChecker.prototype, "ensureDependency")
        .resolves(ok(true));
      sinon.stub(SPFxGenerator, "shouldAddWebPartWithLocalDependencies").resolves(false);

      const result = await SPFxGenerator.doYeomanScaffold(context, inputs, testFolder);
      if (result.isErr()) {
        console.log(result.error);
      }

      chai.expect(result.isOk()).to.eq(true);

      chai.expect(yoInstaller.called).to.be.false;
      chai.expect(generatorInstaller.called).to.be.false;
    });

    it("add web part with local package", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.AppName]: "spfxTestApp",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFolder]: "folder",
        [QuestionNames.SPFxWebpartName]: "hello",
        stage: Stage.addWebpart,
      };

      sinon.stub(GeneratorChecker.prototype, "findGloballyInstalledVersion").resolves(undefined);
      sinon.stub(GeneratorChecker.prototype, "findLocalInstalledVersion").resolves("1.17.4");
      const localYoChecker = sinon
        .stub(YoChecker.prototype, "findLocalInstalledVersion")
        .resolves("4.3.1");
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(cpUtils, "executeCommand").resolves("succeed");
      const yoInstaller = sinon.stub(YoChecker.prototype, "ensureDependency").resolves(ok(true));
      const generatorInstaller = sinon
        .stub(GeneratorChecker.prototype, "ensureDependency")
        .resolves(ok(true));

      const result = await SPFxGenerator.doYeomanScaffold(context, inputs, testFolder);
      if (result.isErr()) {
        console.log(result.error);
      }

      chai.expect(result.isOk()).to.eq(true);

      chai.expect(yoInstaller.called).to.be.false;
      chai.expect(generatorInstaller.called).to.be.false;
      chai.expect(localYoChecker.called).to.be.true;
    });

    it("add web part with installing yo and SPFx locally", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.AppName]: "spfxTestApp",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFolder]: "folder",
        [QuestionNames.SPFxWebpartName]: "hello",
        stage: Stage.addWebpart,
      };

      sinon.stub(GeneratorChecker.prototype, "findGloballyInstalledVersion").resolves("1.17.0");
      sinon.stub(GeneratorChecker.prototype, "findLocalInstalledVersion").resolves(undefined);
      const localYoChecker = sinon
        .stub(YoChecker.prototype, "findLocalInstalledVersion")
        .resolves(undefined);
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(cpUtils, "executeCommand").resolves("succeed");
      const yoInstaller = sinon.stub(YoChecker.prototype, "ensureDependency").resolves(ok(true));
      const generatorInstaller = sinon
        .stub(GeneratorChecker.prototype, "ensureDependency")
        .resolves(ok(true));
      const userConfirm = sinon
        .stub(context.userInteraction, "showMessage")
        .resolves(ok(getLocalizedString("plugins.spfx.addWebPart.install")));

      const result = await SPFxGenerator.doYeomanScaffold(context, inputs, testFolder);
      if (result.isErr()) {
        console.log(result.error);
      }

      chai.expect(result.isOk()).to.eq(true);

      chai.expect(yoInstaller.called).to.be.true;
      chai.expect(generatorInstaller.called).to.be.true;
      chai.expect(localYoChecker.called).to.be.true;
      chai.expect(userConfirm.called).to.be.true;
    });

    it("add web part with upgrading SPFx locally", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.AppName]: "spfxTestApp",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFolder]: "folder",
        [QuestionNames.SPFxWebpartName]: "hello",
        stage: Stage.addWebpart,
      };

      sinon.stub(GeneratorChecker.prototype, "findGloballyInstalledVersion").resolves("1.17.0");
      sinon.stub(GeneratorChecker.prototype, "findLocalInstalledVersion").resolves("1.16.1");
      const localYoChecker = sinon
        .stub(YoChecker.prototype, "findLocalInstalledVersion")
        .resolves("4.3.1");
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(cpUtils, "executeCommand").resolves("succeed");
      const yoInstaller = sinon.stub(YoChecker.prototype, "ensureDependency").resolves(ok(true));
      const generatorInstaller = sinon
        .stub(GeneratorChecker.prototype, "ensureDependency")
        .resolves(ok(true));
      const userConfirm = sinon
        .stub(context.userInteraction, "showMessage")
        .resolves(ok(getLocalizedString("plugins.spfx.addWebPart.upgrade")));

      const result = await SPFxGenerator.doYeomanScaffold(context, inputs, testFolder);

      chai.expect(result.isOk()).to.eq(true);

      chai.expect(yoInstaller.called).to.be.false;
      chai.expect(generatorInstaller.called).to.be.true;
      chai.expect(localYoChecker.called).to.be.true;
      chai.expect(userConfirm.called).to.be.true;
    });

    it("add web part with mismatch SPFx version locally. click 'help' first and then 'continue'", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.AppName]: "spfxTestApp",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFolder]: "folder",
        [QuestionNames.SPFxWebpartName]: "hello",
        stage: Stage.addWebpart,
      };

      sinon.stub(GeneratorChecker.prototype, "findGloballyInstalledVersion").resolves("1.17.0");
      sinon.stub(GeneratorChecker.prototype, "findLocalInstalledVersion").resolves("1.18.2");
      const localYoChecker = sinon
        .stub(YoChecker.prototype, "findLocalInstalledVersion")
        .resolves("4.3.1");
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(cpUtils, "executeCommand").resolves("succeed");
      const yoInstaller = sinon.stub(YoChecker.prototype, "ensureDependency").resolves(ok(true));
      const generatorInstaller = sinon
        .stub(GeneratorChecker.prototype, "ensureDependency")
        .resolves(ok(true));
      const userConfirm = sinon
        .stub(context.userInteraction, "showMessage")
        .onFirstCall()
        .resolves(ok(getLocalizedString("plugins.spfx.addWebPart.versionMismatch.help")))
        .onSecondCall()
        .resolves(ok(getLocalizedString("plugins.spfx.addWebPart.versionMismatch.continue")));
      const openUrl = sinon.stub(context.userInteraction, "openUrl").resolves(ok(true));

      const result = await SPFxGenerator.doYeomanScaffold(context, inputs, testFolder);

      chai.expect(result.isOk()).to.eq(true);

      chai.expect(yoInstaller.called).to.be.false;
      chai.expect(generatorInstaller.called).to.be.false;
      chai.expect(localYoChecker.called).to.be.true;
      chai.expect(userConfirm.called).to.be.true;
      chai.expect(openUrl.called).to.be.true;
    });

    it("add web part with installing SPFx cancel", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.AppName]: "spfxTestApp",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFolder]: "folder",
        [QuestionNames.SPFxWebpartName]: "hello",
        stage: Stage.addWebpart,
      };

      sinon.stub(GeneratorChecker.prototype, "findGloballyInstalledVersion").resolves("1.17.0");
      sinon.stub(GeneratorChecker.prototype, "findLocalInstalledVersion").resolves("1.16.1");
      const localYoChecker = sinon
        .stub(YoChecker.prototype, "findLocalInstalledVersion")
        .resolves("4.3.1");
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(cpUtils, "executeCommand").resolves("succeed");
      const yoInstaller = sinon.stub(YoChecker.prototype, "ensureDependency").resolves(ok(true));
      const generatorInstaller = sinon
        .stub(GeneratorChecker.prototype, "ensureDependency")
        .resolves(ok(true));
      const userConfirm = sinon
        .stub(context.userInteraction, "showMessage")
        .resolves(ok(undefined));

      const result = await SPFxGenerator.doYeomanScaffold(context, inputs, testFolder);

      chai.expect(result.isErr()).to.eq(true);

      chai.expect(yoInstaller.called).to.be.false;
      chai.expect(generatorInstaller.called).to.be.false;
      chai.expect(localYoChecker.called).to.be.false;
      chai.expect(userConfirm.called).to.be.true;
    });

    it("add web part with upgrading SPFx cancel", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.AppName]: "spfxTestApp",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFolder]: "folder",
        [QuestionNames.SPFxWebpartName]: "hello",
        stage: Stage.addWebpart,
      };

      sinon.stub(GeneratorChecker.prototype, "findGloballyInstalledVersion").resolves("1.17.0");
      sinon.stub(GeneratorChecker.prototype, "findLocalInstalledVersion").resolves("1.16.1");
      const localYoChecker = sinon
        .stub(YoChecker.prototype, "findLocalInstalledVersion")
        .resolves("4.3.1");
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(cpUtils, "executeCommand").resolves("succeed");
      const yoInstaller = sinon.stub(YoChecker.prototype, "ensureDependency").resolves(ok(true));
      const generatorInstaller = sinon
        .stub(GeneratorChecker.prototype, "ensureDependency")
        .resolves(ok(true));
      const userConfirm = sinon
        .stub(context.userInteraction, "showMessage")
        .resolves(ok(undefined));

      const result = await SPFxGenerator.doYeomanScaffold(context, inputs, testFolder);

      chai.expect(result.isErr()).to.eq(true);

      chai.expect(yoInstaller.called).to.be.false;
      chai.expect(generatorInstaller.called).to.be.false;
      chai.expect(localYoChecker.called).to.be.false;
      chai.expect(userConfirm.called).to.be.true;
    });

    it("Cancel adding web part due to mismatch SPFx version locally. ", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.AppName]: "spfxTestApp",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFolder]: "folder",
        [QuestionNames.SPFxWebpartName]: "hello",
        stage: Stage.addWebpart,
      };

      sinon.stub(GeneratorChecker.prototype, "findGloballyInstalledVersion").resolves("1.17.0");
      sinon.stub(GeneratorChecker.prototype, "findLocalInstalledVersion").resolves("1.18.2");
      const localYoChecker = sinon
        .stub(YoChecker.prototype, "findLocalInstalledVersion")
        .resolves("4.3.1");
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(cpUtils, "executeCommand").resolves("succeed");
      const yoInstaller = sinon.stub(YoChecker.prototype, "ensureDependency").resolves(ok(true));
      const generatorInstaller = sinon
        .stub(GeneratorChecker.prototype, "ensureDependency")
        .resolves(ok(true));
      const userConfirm = sinon
        .stub(context.userInteraction, "showMessage")
        .resolves(ok(undefined));

      const result = await SPFxGenerator.doYeomanScaffold(context, inputs, testFolder);

      chai.expect(result.isErr()).to.eq(true);

      if (result.isErr()) {
        chai.expect(result.error.name).equal("UserCancel");
      }

      chai.expect(yoInstaller.called).to.be.false;
      chai.expect(generatorInstaller.called).to.be.false;
      chai.expect(localYoChecker.called).to.be.false;
      chai.expect(userConfirm.called).to.be.true;
    });

    it("failed to install yo", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.AppName]: "spfxTestApp",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFolder]: "folder",
        [QuestionNames.SPFxWebpartName]: "hello",
        stage: Stage.addWebpart,
      };

      sinon.stub(GeneratorChecker.prototype, "findGloballyInstalledVersion").resolves("1.17.0");
      sinon.stub(GeneratorChecker.prototype, "findLocalInstalledVersion").resolves(undefined);
      const localYoChecker = sinon
        .stub(YoChecker.prototype, "findLocalInstalledVersion")
        .resolves(undefined);
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(cpUtils, "executeCommand").resolves("succeed");
      const yoInstaller = sinon
        .stub(YoChecker.prototype, "ensureDependency")
        .resolves(err(new SystemError("error", "error", "", "")));
      const userConfirm = sinon
        .stub(context.userInteraction, "showMessage")
        .resolves(ok(getLocalizedString("plugins.spfx.addWebPart.install")));

      const result = await SPFxGenerator.doYeomanScaffold(context, inputs, testFolder);

      chai.expect(result.isOk()).to.eq(false);

      chai.expect(yoInstaller.called).to.be.true;
      chai.expect(localYoChecker.called).to.be.true;
      chai.expect(userConfirm.called).to.be.true;
    });

    it("failed to install SPFx generator", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.AppName]: "spfxTestApp",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFolder]: "folder",
        [QuestionNames.SPFxWebpartName]: "hello",
        stage: Stage.addWebpart,
      };

      sinon.stub(GeneratorChecker.prototype, "findGloballyInstalledVersion").resolves("1.17.0");
      sinon.stub(GeneratorChecker.prototype, "findLocalInstalledVersion").resolves("1.16.1");
      const localYoChecker = sinon
        .stub(YoChecker.prototype, "findLocalInstalledVersion")
        .resolves("4.3.1");
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(cpUtils, "executeCommand").resolves("succeed");
      const generatorInstaller = sinon
        .stub(GeneratorChecker.prototype, "ensureDependency")
        .resolves(err(new SystemError("error", "error", "", "")));
      const userConfirm = sinon
        .stub(context.userInteraction, "showMessage")
        .resolves(ok(getLocalizedString("plugins.spfx.addWebPart.upgrade")));

      const result = await SPFxGenerator.doYeomanScaffold(context, inputs, testFolder);
      chai.expect(result.isOk()).to.eq(false);
      chai.expect(generatorInstaller.called).to.be.true;
      chai.expect(localYoChecker.called).to.be.true;
      chai.expect(userConfirm.called).to.be.true;
    });

    it("Cannot find version in .yo-rc.json file", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.AppName]: "spfxTestApp",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFolder]: "folder",
        [QuestionNames.SPFxWebpartName]: "hello",
        stage: Stage.addWebpart,
      };
      sinon.restore();
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readJson").resolves({
        "@microsoft/generator-sharepoint": {
          solutionName: "fakedSolutionName",
        },
      });

      const result = await SPFxGenerator.doYeomanScaffold(context, inputs, testFolder);
      chai.expect(result.isErr()).to.eq(true);
      if (result.isErr()) {
        chai.expect(result.error.name).equals("SolutionVersionMissing");
      }
    });

    it("Empty content in .yo-rc.json file", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.AppName]: "spfxTestApp",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFolder]: "folder",
        [QuestionNames.SPFxWebpartName]: "hello",
        stage: Stage.addWebpart,
      };
      sinon.restore();
      sinon.stub(fs, "pathExists").resolves(true);
      sinon.stub(fs, "readJson").resolves({});

      const result = await SPFxGenerator.doYeomanScaffold(context, inputs, testFolder);
      chai.expect(result.isErr()).to.eq(true);
      if (result.isErr()) {
        chai.expect(result.error.name).equals("SolutionVersionMissing");
      }
    });

    it("Cannot find .yo-rc.json file", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: testFolder,
        [QuestionNames.AppName]: "spfxTestApp",
        [QuestionNames.SPFxSolution]: "new",
        [QuestionNames.SPFxFolder]: "folder",
        [QuestionNames.SPFxWebpartName]: "hello",
        stage: Stage.addWebpart,
      };
      sinon.restore();
      sinon.stub(fs, "pathExists").resolves(false);

      const result = await SPFxGenerator.doYeomanScaffold(context, inputs, testFolder);
      chai.expect(result.isErr()).to.eq(true);
      if (result.isErr()) {
        chai.expect(result.error instanceof FileNotFoundError).to.eq(true);
      }
    });
  });
});

describe("Utils", () => {
  it("truncate name with app name suffix", () => {
    const appName = "thisisasuperlongappNameWithSuffix${{APP_NAME_SUFFIX}}";
    const res = Utils.truncateAppShortName(appName);
    chai.expect(res).equals("thisisasuperlongappNameWi${{APP_NAME_SUFFIX}}");
  });
  it("no need to truncate name with app name with suffix", () => {
    const appName = "appNameWithSuffix${{APP_NAME_SUFFIX}}";
    const res = Utils.truncateAppShortName(appName);
    chai.expect(res).equals("appNameWithSuffix${{APP_NAME_SUFFIX}}");
  });

  it("truncate name with app name without suffix", () => {
    const appName = "thisisasuperlongappNameWithoutSuffix";
    const res = Utils.truncateAppShortName(appName);
    chai.expect(res).equals("thisisasuperlongappNameWithout");
  });

  it("no need to truncate name with app name without suffix", () => {
    const appName = "appNameWithoutSuffix";
    const res = Utils.truncateAppShortName(appName);
    chai.expect(res).equals("appNameWithoutSuffix");
  });

  describe("getShellOptionValue", () => {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });

    it("windows", () => {
      sandbox.stub(os, "type").returns("Windows_NT");
      const res = getShellOptionValue();

      chai.expect(res).equal("cmd.exe");
    });

    it("non windowns", () => {
      sandbox.stub(os, "type").returns("Linux");
      const res = getShellOptionValue();

      chai.expect(res).true;
    });
  });
});

describe("SPFxGeneratorNew", () => {
  const gtools = new MockTools();
  setTools(gtools);
  const generator = new SPFxGeneratorNew();
  const context = createContext();
  describe("activate", () => {
    it("happy path", () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: TabCapabilityOptions.SPFxTab().id,
        [QuestionNames.ProjectType]: ProjectTypeOptions.teamsAppOptionId,
        [QuestionNames.TeamsAppType]: TeamsProjectTypeOptions.tabOptionId,
        [QuestionNames.SPFxSolution]: "new",
      };
      const isActive = generator.activate(context, inputs);
      chai.expect(isActive).to.be.true;
    });
  });
  describe("getTemplateInfos", () => {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });
    it("happy path", async () => {
      sandbox.stub(SPFxGenerator, "doYeomanScaffold").resolves(ok(""));
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: TabCapabilityOptions.SPFxTab().id,
        [QuestionNames.ProjectType]: ProjectTypeOptions.teamsAppOptionId,
        [QuestionNames.TeamsAppType]: TeamsProjectTypeOptions.tabOptionId,
        [QuestionNames.SPFxSolution]: "new",
      };
      const res = await generator.getTemplateInfos(context, inputs, "");
      chai.expect(res.isOk()).to.be.true;
    });
    it("doYeomanScaffold error", async () => {
      sandbox.stub(SPFxGenerator, "doYeomanScaffold").resolves(err(new UserCancelError()));
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: TabCapabilityOptions.SPFxTab().id,
        [QuestionNames.ProjectType]: ProjectTypeOptions.teamsAppOptionId,
        [QuestionNames.TeamsAppType]: TeamsProjectTypeOptions.tabOptionId,
        [QuestionNames.SPFxSolution]: "new",
      };
      const res = await generator.getTemplateInfos(context, inputs, "");
      chai.expect(res.isErr()).to.be.true;
    });
  });
});

describe("SPFxGeneratorImport", () => {
  const gtools = new MockTools();
  setTools(gtools);
  const generator = new SPFxGeneratorImport();
  const context = createContext();
  describe("activate", () => {
    it("happy path", () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.Capabilities]: TabCapabilityOptions.SPFxTab().id,
        [QuestionNames.ProjectType]: ProjectTypeOptions.teamsAppOptionId,
        [QuestionNames.TeamsAppType]: TeamsProjectTypeOptions.tabOptionId,
        [QuestionNames.SPFxSolution]: "import",
      };
      const isActive = generator.activate(context, inputs);
      chai.expect(isActive).to.be.true;
    });
  });
  describe("getTemplateInfos", () => {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });
    it("happy path", async () => {
      sandbox.stub(SPFxGenerator, "copySPFxSolution").resolves();
      sandbox.stub(SPFxGenerator, "getWebpartManifest").resolves({
        id: "test-id",
        preconfiguredEntries: [{ title: { default: "defaultTitle" } }],
      });
      sandbox.stub(SPFxGenerator, "getNodeVersion").resolves("18.0");
      sandbox.stub(SPFxGenerator, "getSolutionVersion").resolves("1.21.0");
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.AppName]: "testspfx",
        [QuestionNames.Capabilities]: TabCapabilityOptions.SPFxTab().id,
        [QuestionNames.ProjectType]: ProjectTypeOptions.teamsAppOptionId,
        [QuestionNames.TeamsAppType]: TeamsProjectTypeOptions.tabOptionId,
        [QuestionNames.SPFxSolution]: "import",
        [QuestionNames.SPFxFolder]: "c:\\test",
      };
      const res = await generator.getTemplateInfos(context, inputs, "");
      chai.expect(res.isOk()).to.be.true;
    });

    it("happy path - SPFx lower than 1.21", async () => {
      sandbox.stub(SPFxGenerator, "copySPFxSolution").resolves();
      sandbox.stub(SPFxGenerator, "getWebpartManifest").resolves({
        id: "test-id",
        preconfiguredEntries: [{ title: { default: "defaultTitle" } }],
      });
      sandbox.stub(SPFxGenerator, "getNodeVersion").resolves("18.0");
      sandbox.stub(SPFxGenerator, "getSolutionVersion").resolves("1.17.0");
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.AppName]: "testspfx",
        [QuestionNames.Capabilities]: TabCapabilityOptions.SPFxTab().id,
        [QuestionNames.ProjectType]: ProjectTypeOptions.teamsAppOptionId,
        [QuestionNames.TeamsAppType]: TeamsProjectTypeOptions.tabOptionId,
        [QuestionNames.SPFxSolution]: "import",
        [QuestionNames.SPFxFolder]: "c:\\test",
      };
      const res = await generator.getTemplateInfos(context, inputs, "");
      chai.expect(res.isOk()).to.be.true;
    });

    it("throw error", async () => {
      sandbox.stub(SPFxGenerator, "copySPFxSolution").rejects(new Error());
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.AppName]: "testspfx",
        [QuestionNames.Capabilities]: TabCapabilityOptions.SPFxTab().id,
        [QuestionNames.ProjectType]: ProjectTypeOptions.teamsAppOptionId,
        [QuestionNames.TeamsAppType]: TeamsProjectTypeOptions.tabOptionId,
        [QuestionNames.SPFxSolution]: "import",
      };
      const res = await generator.getTemplateInfos(context, inputs, "");
      chai.expect(res.isErr()).to.be.true;
    });

    it("throw FxError", async () => {
      sandbox.stub(SPFxGenerator, "copySPFxSolution").rejects(new UserCancelError());
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.AppName]: "testspfx",
        [QuestionNames.Capabilities]: TabCapabilityOptions.SPFxTab().id,
        [QuestionNames.ProjectType]: ProjectTypeOptions.teamsAppOptionId,
        [QuestionNames.TeamsAppType]: TeamsProjectTypeOptions.tabOptionId,
        [QuestionNames.SPFxSolution]: "import",
      };
      const res = await generator.getTemplateInfos(context, inputs, "");
      chai.expect(res.isErr()).to.be.true;
    });

    it("RetrieveSPFxInfoError", async () => {
      sandbox.stub(SPFxGenerator, "copySPFxSolution").resolves();
      sandbox.stub(SPFxGenerator, "getWebpartManifest").resolves({});
      sandbox.stub(SPFxGenerator, "getNodeVersion").resolves("18.0");
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.AppName]: "testspfx",
        [QuestionNames.Capabilities]: TabCapabilityOptions.SPFxTab().id,
        [QuestionNames.ProjectType]: ProjectTypeOptions.teamsAppOptionId,
        [QuestionNames.TeamsAppType]: TeamsProjectTypeOptions.tabOptionId,
        [QuestionNames.SPFxSolution]: "import",
      };
      const res = await generator.getTemplateInfos(context, inputs, "");
      chai.expect(res.isErr()).to.be.true;
    });
  });

  describe("post", () => {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });
    it("happy path", async () => {
      sandbox.stub(SPFxGenerator, "updateSPFxTemplate").resolves();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.AppName]: "testspfx",
        [QuestionNames.Capabilities]: TabCapabilityOptions.SPFxTab().id,
        [QuestionNames.ProjectType]: ProjectTypeOptions.teamsAppOptionId,
        [QuestionNames.TeamsAppType]: TeamsProjectTypeOptions.tabOptionId,
        [QuestionNames.SPFxSolution]: "import",
      };
      const res = await generator.post(context, inputs, "");
      chai.expect(res.isOk()).to.be.true;
    });

    it("happy path with template variables not exist", async () => {
      sandbox.stub(SPFxGenerator, "updateSPFxTemplate").resolves();
      context.templateVariables = undefined;

      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.AppName]: "testspfx",
        [QuestionNames.Capabilities]: TabCapabilityOptions.SPFxTab().id,
        [QuestionNames.ProjectType]: ProjectTypeOptions.teamsAppOptionId,
        [QuestionNames.TeamsAppType]: TeamsProjectTypeOptions.tabOptionId,
        [QuestionNames.SPFxSolution]: "import",
      };
      const res = await generator.post(context, inputs, "");
      chai.expect(res.isOk()).to.be.true;
    });

    it("throw error", async () => {
      sandbox.stub(SPFxGenerator, "updateSPFxTemplate").rejects(new Error());
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.AppName]: "testspfx",
        [QuestionNames.Capabilities]: TabCapabilityOptions.SPFxTab().id,
        [QuestionNames.ProjectType]: ProjectTypeOptions.teamsAppOptionId,
        [QuestionNames.TeamsAppType]: TeamsProjectTypeOptions.tabOptionId,
        [QuestionNames.SPFxSolution]: "import",
      };
      const res = await generator.post(context, inputs, "");
      chai.expect(res.isErr()).to.be.true;
    });

    it("throw FxError", async () => {
      sandbox.stub(SPFxGenerator, "updateSPFxTemplate").rejects(new UserCancelError());
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.AppName]: "testspfx",
        [QuestionNames.Capabilities]: TabCapabilityOptions.SPFxTab().id,
        [QuestionNames.ProjectType]: ProjectTypeOptions.teamsAppOptionId,
        [QuestionNames.TeamsAppType]: TeamsProjectTypeOptions.tabOptionId,
        [QuestionNames.SPFxSolution]: "import",
      };
      const res = await generator.post(context, inputs, "");
      chai.expect(res.isErr()).to.be.true;
    });
  });
});
