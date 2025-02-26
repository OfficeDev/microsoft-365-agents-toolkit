// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yefuwang@microsoft.com
 */

import {
  Context,
  devPreview,
  err,
  Inputs,
  ManifestUtil,
  ok,
  Platform,
  SystemError,
} from "@microsoft/teamsfx-api";
import * as chai from "chai";
import fs from "fs";
import fse from "fs-extra";
import "mocha";
import mockfs from "mock-fs";
import mockedEnv, { RestoreFn } from "mocked-env";
import { OfficeAddinManifest } from "office-addin-manifest";
import * as path from "path";
import proxyquire from "proxyquire";
import * as sinon from "sinon";
import * as uuid from "uuid";
import { createContext, setTools } from "../../../src/common/globalVars";
import { cpUtils } from "../../../src/component/deps-checker/";
import { manifestUtils } from "../../../src/component/driver/teamsApp/utils/ManifestUtils";
import {
  getHost,
  OfficeAddinGenerator,
  OfficeAddinGeneratorNew,
} from "../../../src/component/generator/officeAddin/generator";
import { HelperMethods } from "../../../src/component/generator/officeAddin/helperMethods";
import { TemplateNames } from "../../../src/component/generator/templates/templateNames";
import { envUtil } from "../../../src/component/utils/envUtil";
import { UserCancelError } from "../../../src/error";
import {
  CapabilityOptions,
  ProgrammingLanguage,
  ProjectTypeOptions,
  QuestionNames,
} from "../../../src/question";
import { MockTools } from "../../core/utils";

describe("OfficeAddinGenerator for Outlook Addin", function () {
  const testFolder = path.resolve("./tmp");
  let context: Context;
  let mockedEnvRestore: RestoreFn;
  const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");

  beforeEach(async () => {
    mockedEnvRestore = mockedEnv({ TEAMSFX_V3: "true" }, { clear: true });
    const gtools = new MockTools();
    setTools(gtools);
    context = createContext();

    await fse.ensureDir(testFolder);
    sinon.stub(fs, "stat").resolves();
    sinon.stub(cpUtils, "executeCommand").resolves("succeed");
    const manifestId = uuid.v4();
    sinon.stub(fs, "readFile").resolves(new Buffer(`{"id": "${manifestId}"}`));
    sinon.stub(fs, "writeFile").resolves();
    sinon.stub(fs, "rename").resolves();
    sinon.stub(fs, "copyFile").resolves();
    sinon.stub(fse, "remove").resolves();
    sinon.stub(fse, "readJson").resolves({});
    sinon.stub(fse, "ensureFile").resolves();
    sinon.stub(fse, "writeJSON").resolves();
  });

  it("should scaffold taskpane successfully on happy path if project-type is outlookAddin", async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "outlook-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.outlookAddin().id;
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.OfficeAddinFolder] = undefined;
    inputs[QuestionNames.ProgrammingLanguage] = "typescript";
    sinon.stub(HelperMethods, "fetchAndUnzip").resolves(ok(undefined));
    sinon.stub(OfficeAddinManifest, "modifyManifestFile").resolves({});
    const result = await OfficeAddinGenerator.doScaffolding(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
  });

  it("should scaffold taskpane failed, throw error", async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "outlook-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.outlookAddin().id;
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.OfficeAddinFolder] = undefined;
    inputs[QuestionNames.ProgrammingLanguage] = "typescript";
    sinon.stub(HelperMethods, "fetchAndUnzip").rejects(new UserCancelError());
    sinon.stub(OfficeAddinManifest, "modifyManifestFile").resolves({});
    const result = await OfficeAddinGenerator.doScaffolding(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
  });

  it("should copy addin files and updateManifest if addin folder is specified with json manifest", async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "outlook-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.outlookAddin().id;
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.OfficeAddinFolder] = "somepath";
    inputs[QuestionNames.ProgrammingLanguage] = "typescript";
    inputs[QuestionNames.OfficeAddinManifest] = "manifest.json";

    const copyAddinFilesStub = sinon
      .stub(HelperMethods, "copyAddinFiles")
      .callsFake((from: string, to: string) => {
        return;
      });
    const updateManifestStub = sinon
      .stub(HelperMethods, "updateManifest")
      .callsFake(async (destination: string, manifestPath: string) => {
        return;
      });

    sinon.stub<any, any>(ManifestUtil, "loadFromPath").resolves({
      extensions: [
        {
          requirements: {
            scopes: ["mail"],
          },
        },
      ],
    });

    const result = await OfficeAddinGenerator.doScaffolding(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
    chai.expect(copyAddinFilesStub.calledOnce).to.be.true;
    chai.expect(updateManifestStub.calledOnce).to.be.true;
    chai.expect(inputs[QuestionNames.OfficeAddinHost]).to.eq("Outlook");

    const hostResult = await getHost(inputs[QuestionNames.OfficeAddinFolder]);
    chai.expect(hostResult).to.equal("Outlook");
  });

  it("should copy addin files and convert manifest if addin folder is specified with xml manifest", async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "outlook-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.outlookAddin().id;
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.OfficeAddinFolder] = "somepath";
    inputs[QuestionNames.ProgrammingLanguage] = "typescript";
    inputs[QuestionNames.OfficeAddinManifest] = "manifest.xml";

    let progressBarStartCalled = 0;
    let progressBarNextCalled = 0;
    let progessBarEndCalled = 0;
    const createProgressBarStub = sinon.stub(context.userInteraction, "createProgressBar").returns({
      start: async () => {
        progressBarStartCalled++;
      },
      next: async () => {
        progressBarNextCalled++;
      },
      end: async () => {
        progessBarEndCalled++;
      },
    });

    const copyAddinFilesStub = sinon
      .stub(HelperMethods, "copyAddinFiles")
      .callsFake((from: string, to: string) => {
        return;
      });
    const updateManifestStub = sinon
      .stub(HelperMethods, "updateManifest")
      .callsFake(async (destination: string, manifestPath: string) => {
        return;
      });
    const convertProjectStub = sinon
      .stub()
      .callsFake(async (manifestPath?: string, backupPath?: string) => {
        return;
      });

    const generator = proxyquire("../../../src/component/generator/officeAddin/generator", {
      "office-addin-project": {
        convertProject: convertProjectStub,
      },
    });

    sinon.stub<any, any>(ManifestUtil, "loadFromPath").resolves({
      extensions: [
        {
          requirements: {
            scopes: ["mail"],
          },
        },
      ],
    });

    const result = await generator.OfficeAddinGenerator.doScaffolding(context, inputs, testFolder);

    chai.expect(result.isOk()).to.eq(true);
    chai.expect(copyAddinFilesStub.calledOnce).to.be.true;
    chai.expect(updateManifestStub.calledOnce).to.be.true;
    chai.expect(convertProjectStub.calledOnce).to.be.true;
    chai.expect(inputs[QuestionNames.OfficeAddinHost]).to.eq("Outlook");
    chai.expect(progressBarStartCalled).to.eq(1);
    chai.expect(progressBarNextCalled).to.eq(3);
    chai.expect(progessBarEndCalled).to.eq(1);

    const hostResult = await getHost(inputs[QuestionNames.OfficeAddinFolder]);
    chai.expect(hostResult).to.equal("Outlook");
  });

  afterEach(async () => {
    sinon.restore();
    mockedEnvRestore();
    if (await fse.pathExists(testFolder)) {
      await fse.remove(testFolder);
    }
  });
});

describe("HelperMethods", async () => {
  describe("updateManifest", () => {
    const sandbox = sinon.createSandbox();
    const manifestPath = "manifestPath";
    const manifestTemplatePath = "manifestTemplatePath";
    let writePathResult: devPreview.DevPreviewSchema | undefined = undefined;

    beforeEach(() => {
      sandbox.stub(ManifestUtil, "loadFromPath").callsFake(async (path) => {
        if (path === manifestPath) {
          return {
            extensions: [],
            authorization: {
              permissions: {
                resourceSpecific: [],
              },
            },
          } as unknown as devPreview.DevPreviewSchema;
        } else if (path === manifestTemplatePath) {
          return {
            extensions: undefined,
            authorization: undefined,
          } as unknown as devPreview.DevPreviewSchema;
        }

        throw new Error("Invalid path");
      });

      sandbox.stub(ManifestUtil, "writeToPath").callsFake(async (path, manifest) => {
        writePathResult = manifest as devPreview.DevPreviewSchema;
        return;
      });

      sandbox.stub(manifestUtils, "getTeamsAppManifestPath").returns(manifestTemplatePath);
    });

    afterEach(() => {
      sandbox.restore();
      writePathResult = undefined;
    });

    it("should update manifest's extenstions and authorization", async () => {
      sandbox.stub(fse, "pathExists").resolves(true);
      await HelperMethods.updateManifest("", manifestPath);

      chai.assert.isDefined(writePathResult);
      chai.assert.equal(writePathResult?.extensions?.length, 0);
      chai.assert.equal(writePathResult?.authorization?.permissions?.resourceSpecific?.length, 0);
    });

    it("should early return if there's no appPackage folder", async () => {
      sandbox.stub(fse, "pathExists").resolves(false);
      await HelperMethods.updateManifest("", manifestPath);

      chai.assert.isUndefined(writePathResult, "writeToPath should not be called");
    });
  });

  describe("copyAddinFiles", () => {
    const projectRoot = "/home/user/teamsapp";

    beforeEach(() => {
      mockfs({
        "/home/user/teamsapp/.gitignore": "xxx",
        "/home/user/teamsapp/project": {
          file1: "xxx",
          file2: "yyy",
        },
        "/home/user/teamsapp/node_modules": {
          file3: "xxx",
        },
      });
    });

    afterEach(() => {
      mockfs.restore();
    });

    it("should copy project files and .gitignore but ignore node_modules", async () => {
      try {
        const destination = "/home/user/destination";
        HelperMethods.copyAddinFiles(projectRoot, destination);
        chai.assert.equal(fs.existsSync(path.join(destination, "project", "file1")), true);
        chai.assert.equal(fs.existsSync(path.join(destination, "project", "file2")), true);
        chai.assert.equal(fs.existsSync(path.join(destination, ".gitignore")), true);
        chai.assert.equal(fs.existsSync(path.join(destination, "node_modules")), false);
      } catch (err) {
        chai.assert.fail(err);
      }
    });
  });

  describe("moveManifestLocation", () => {
    const projectRoot = "/home/user/addin";

    beforeEach(() => {
      mockfs({
        "/home/user/addin/manifest.json": "{}",
        "/home/user/addin/assets": {
          file1: "xxx",
        },
        "/home/user/addin/webpack.config.js": JSON.stringify([
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.json",
            to: "[name]" + "[ext]",
          },
        ]),
        "/home/user/addin/package.json": JSON.stringify({
          scripts: {
            start: "office-addin-debugging start manifest.json",
            stop: "office-addin-debugging stop manifest.json",
            validate: "office-addin-manifest validate manifest.json",
          },
        }),
        "/home/user/addin/src/taskpane/taskpane.html": `<img width="90" height="90" src="../../assets/logo-filled.png" alt="Contoso" title="Contoso" />`,
      });
    });

    afterEach(() => {
      mockfs.restore();
    });

    it("should move manifest.json into appPackage folder", async () => {
      await HelperMethods.moveManifestLocation(projectRoot, "manifest.json");
      chai.assert.isFalse(await fse.pathExists(path.join(projectRoot, "manifest.json")));
      chai.assert.isFalse(await fse.pathExists(path.join(projectRoot, "assets")));

      chai.assert.isTrue(
        await fse.pathExists(path.join(projectRoot, "appPackage", "manifest.json"))
      );
      chai.assert.isTrue(
        await fse.pathExists(path.join(projectRoot, "appPackage", "assets", "file1"))
      );

      const webpackConfigPath = path.join(projectRoot, "webpack.config.js");
      const webpackConfigJson = JSON.parse(await fse.readFile(webpackConfigPath, "utf8"));
      chai.assert.equal(webpackConfigJson[0].from, "appPackage/assets/*");
      chai.assert.equal(webpackConfigJson[1].from, "appPackage/manifest*.json");

      const packageJsonPath = path.join(projectRoot, "package.json");
      const packageJson = JSON.parse(await fse.readFile(packageJsonPath, "utf8"));
      chai.assert.equal(
        packageJson.scripts.start,
        "office-addin-debugging start appPackage/manifest.json"
      );

      chai.assert.equal(
        packageJson.scripts.stop,
        "office-addin-debugging stop appPackage/manifest.json"
      );
      chai.assert.equal(
        packageJson.scripts.validate,
        "office-addin-manifest validate appPackage/manifest.json"
      );

      const htmlPath = path.join(projectRoot, "src", "taskpane", "taskpane.html");
      const html = await fse.readFile(htmlPath, "utf8");
      chai.assert.equal(
        html,
        `<img width="90" height="90" src="../../appPackage/assets/logo-filled.png" alt="Contoso" title="Contoso" />`
      );
    });
  });
});

describe("OfficeAddinGenerator for Office Addin", function () {
  const testFolder = path.resolve("./tmp");
  let context: Context;
  let mockedEnvRestore: RestoreFn = () => {};
  const mockedError = new SystemError("mockedSource", "mockedError", "mockedMessage");

  beforeEach(async () => {
    mockedEnvRestore = mockedEnv({ clear: true });
    const gtools = new MockTools();
    setTools(gtools);
    context = createContext();

    await fse.ensureDir(testFolder);
    sinon.stub(fs, "stat").resolves();
    sinon.stub(cpUtils, "executeCommand").resolves("succeed");
    const manifestId = uuid.v4();
    sinon.stub(fs, "readFile").resolves(new Buffer(`{"id": "${manifestId}"}`));
    sinon.stub(fs, "writeFile").resolves();
    sinon.stub(fs, "rename").resolves();
    sinon.stub(fs, "copyFile").resolves();
    sinon.stub(fse, "remove").resolves();
    sinon.stub(fse, "readJson").resolves({});
    sinon.stub(fse, "ensureFile").resolves();
    sinon.stub(fse, "writeJSON").resolves();
  });

  it("should scaffold taskpane successfully on happy path if project-type is officeAddin and capability is json-taskpane", async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: testFolder,
      "app-name": "office-addin-test",
    };
    inputs[QuestionNames.ProjectType] = ProjectTypeOptions.outlookAddin().id;
    inputs[QuestionNames.Capabilities] = CapabilityOptions.officeAddinTaskpane().id;
    inputs[QuestionNames.OfficeAddinFolder] = undefined;
    inputs[QuestionNames.ProgrammingLanguage] = "typescript";
    const result = await OfficeAddinGenerator.doScaffolding(context, inputs, testFolder);
    chai.expect(result.isOk()).to.eq(true);
  });
  afterEach(async () => {
    sinon.restore();
    mockedEnvRestore();
    if (await fse.pathExists(testFolder)) {
      await fse.remove(testFolder);
    }
  });
});

describe("OfficeAddinGeneratorNew", () => {
  const gtools = new MockTools();
  setTools(gtools);
  const generator = new OfficeAddinGeneratorNew();
  const context = createContext();
  const sandbox = sinon.createSandbox();
  describe("active()", () => {
    it(`should return true`, async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
      };
      inputs[QuestionNames.ProjectType] = ProjectTypeOptions.outlookAddin().id;
      inputs[QuestionNames.ProgrammingLanguage] = ProgrammingLanguage.JS;
      inputs[QuestionNames.TemplateName] = TemplateNames.OutlookTaskpane;
      const res = generator.activate(context, inputs);
      chai.assert.isTrue(res);
    });

    it(`should return false`, async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
      };
      inputs[QuestionNames.ProjectType] = ProjectTypeOptions.bot().id;
      inputs[QuestionNames.ProgrammingLanguage] = ProgrammingLanguage.JS;
      const res = generator.activate(context, inputs);
      chai.assert.isFalse(res);
    });
  });

  describe("getTemplateInfos()", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it(`should return office-addin-config template officeMetaOS`, async () => {
      sandbox.stub(OfficeAddinGenerator, "doScaffolding").resolves(ok(undefined));
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
      };
      inputs[QuestionNames.ProjectType] = ProjectTypeOptions.officeMetaOS().id;
      inputs[QuestionNames.Capabilities] = CapabilityOptions.officeAddinImport().id;
      inputs[QuestionNames.TemplateName] = TemplateNames.OfficeAddinCommon;
      const res = await generator.getTemplateInfos(context, inputs, "./");
      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        const templates = res.value;
        chai.assert.isTrue(templates.length === 1);
        const template = templates[0];
        chai.assert.equal(template.templateName, "office-addin-config");
        chai.assert.isTrue(template.language === ProgrammingLanguage.TS);
      }
    });

    it(`should return office-addin-config template outlookAddin`, async () => {
      sandbox.stub(OfficeAddinGenerator, "doScaffolding").resolves(ok(undefined));
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
      };
      inputs[QuestionNames.ProjectType] = ProjectTypeOptions.outlookAddin().id;
      inputs[QuestionNames.Capabilities] = CapabilityOptions.outlookAddinImport().id;
      inputs[QuestionNames.TemplateName] = TemplateNames.OfficeAddinCommon;
      const res = await generator.getTemplateInfos(context, inputs, "./");
      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        const templates = res.value;
        chai.assert.isTrue(templates.length === 1);
        const template = templates[0];
        chai.assert.equal(template.templateName, "office-addin-config");
        chai.assert.isTrue(template.language === ProgrammingLanguage.TS);
      }
    });

    it(`should return office-addin-outlook-taskpane template`, async () => {
      sandbox.stub(OfficeAddinGenerator, "doScaffolding").resolves(ok(undefined));
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
      };
      inputs[QuestionNames.ProjectType] = ProjectTypeOptions.outlookAddin().id;
      inputs[QuestionNames.Capabilities] = CapabilityOptions.officeAddinTaskpane().id;
      inputs[QuestionNames.TemplateName] = TemplateNames.OutlookTaskpane;
      const res = await generator.getTemplateInfos(context, inputs, "./");
      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        const templates = res.value;
        chai.assert.isTrue(templates.length === 1);
        const template = templates[0];
        chai.assert.isTrue(template.templateName === TemplateNames.OutlookTaskpane);
        chai.assert.isTrue(template.language === ProgrammingLanguage.TS);
      }
    });
    it(`should return office-addin-outlook-taskpane template`, async () => {
      sandbox.stub(OfficeAddinGenerator, "doScaffolding").resolves(ok(undefined));
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
      };
      inputs[QuestionNames.ProjectType] = ProjectTypeOptions.officeMetaOS().id;
      inputs[QuestionNames.Capabilities] = CapabilityOptions.officeAddinTaskpane().id;
      inputs[QuestionNames.TemplateName] = TemplateNames.WXPTaskpane;
      const res = await generator.getTemplateInfos(context, inputs, "./");
      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        const templates = res.value;
        chai.assert.isTrue(templates.length === 1);
        const template = templates[0];
        chai.assert.isTrue(template.templateName === TemplateNames.WXPTaskpane);
        chai.assert.isTrue(template.language === ProgrammingLanguage.TS);
      }
    });
    it("should fail", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
      };
      sandbox.stub(OfficeAddinGenerator, "doScaffolding").resolves(err(new UserCancelError()));
      const res = await generator.getTemplateInfos(context, inputs, "./");
      chai.assert.isTrue(res.isErr());
    });
  });

  describe("post()", () => {
    afterEach(() => {
      sandbox.restore();
    });
    it(`happy`, async () => {
      sandbox.stub(envUtil, "listEnv").resolves(ok(["dev", "dev2"]));
      const reset = sandbox.stub(envUtil, "resetEnv").resolves();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
      };
      inputs[QuestionNames.OfficeAddinFolder] = "testfolder";
      const res = await generator.post(context, inputs, "./");
      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(reset.calledTwice);
    });
    it(`not import`, async () => {
      const reset = sandbox.stub(envUtil, "resetEnv").resolves();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
      };
      const res = await generator.post(context, inputs, "./");
      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(reset.notCalled);
    });
    it(`list env error`, async () => {
      sandbox.stub(envUtil, "listEnv").resolves(err(new UserCancelError()));
      const reset = sandbox.stub(envUtil, "resetEnv").resolves();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
      };
      const res = await generator.post(context, inputs, "./");
      chai.assert.isTrue(res.isOk());
      chai.assert.isTrue(reset.notCalled);
    });
  });
});

describe("doScaffolding()", () => {
  it("doScaffolding: should failed because of invalid addin-host", async () => {
    const inputs: Inputs = {
      platform: Platform.CLI,
      projectPath: ".",
      "app-name": "outlook-addin-test",
      [QuestionNames.OfficeAddinHost]: "invalid",
    };
    inputs[QuestionNames.Capabilities] = "json-taskpane";
    inputs[QuestionNames.OfficeAddinFolder] = undefined;
    inputs[QuestionNames.ProgrammingLanguage] = "typescript";
    const context = createContext();
    const res = await OfficeAddinGenerator.doScaffolding(context, inputs, ".");
    chai.assert.isTrue(res.isOk());
  });
});
