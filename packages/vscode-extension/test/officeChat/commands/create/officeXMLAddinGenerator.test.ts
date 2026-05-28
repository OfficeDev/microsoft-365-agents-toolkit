import { Context, Inputs, Platform, ok } from "@microsoft/teamsfx-api";
import { QuestionNames, HelperMethods, ProgrammingLanguage } from "@microsoft/teamsfx-core";
import * as chai from "chai";
import * as sinon from "sinon";
import * as childProcess from "child_process";
import "mocha";
import { OfficeXMLAddinGenerator } from "../../../../src/officeChat/commands/create/officeXMLAddinGenerator/generator";
import { OfficeAddinManifest } from "office-addin-manifest";

describe("OfficeXMLAddinGenerator", () => {
  const generator = new OfficeXMLAddinGenerator();
  const context: Context = {
    userInteraction: undefined as any,
    logProvider: undefined as any,
    telemetryReporter: undefined as any,
    tokenProvider: undefined as any,
  };
  describe("activate", () => {
    it(`should return true`, async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "./",
        [QuestionNames.ProjectType]: "office-xml-addin-type",
        [QuestionNames.OfficeAddinHost]: "word",
        agent: "office",
      };
      const res = generator.activate(context, inputs);
      chai.assert.isTrue(res);
    });

    it(`should return false`, async () => {
      const inputs: Inputs = {
        platform: Platform.VSCode,
        projectPath: "./",
        [QuestionNames.ProjectType]: "office-xml-addin-type",
        [QuestionNames.OfficeAddinHost]: "outlook",
      };
      const res = generator.activate(context, inputs);
      chai.assert.isFalse(res);
    });
  });

  describe("getTemplateInfos", () => {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });
    it("happy path for word-taskpane", async () => {
      sandbox.stub(HelperMethods, "fetchAndUnzip").resolves(ok(undefined));
      sandbox.stub(OfficeXMLAddinGenerator, "childProcessExec").resolves();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.ProjectType]: "office-xml-addin-type",
        [QuestionNames.OfficeAddinHost]: "word",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.Capabilities]: "word-taskpane",
        agent: "office",
      };
      const res = await generator.getTemplateInfos(context, inputs, "./");
      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.equal(res.value.length, 2);
      }
    });
    it("happy path for word-manifest", async () => {
      sandbox.stub(HelperMethods, "fetchAndUnzip").resolves(ok(undefined));
      sandbox.stub(OfficeXMLAddinGenerator, "childProcessExec").resolves();
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
        [QuestionNames.ProjectType]: "office-xml-addin-type",
        [QuestionNames.OfficeAddinHost]: "word",
        [QuestionNames.ProgrammingLanguage]: ProgrammingLanguage.TS,
        [QuestionNames.Capabilities]: "word-manifest",
        agent: "office",
      };
      const res = await generator.getTemplateInfos(context, inputs, "./");
      chai.assert.isTrue(res.isOk());
      if (res.isOk()) {
        chai.assert.equal(res.value.length, 3);
      }
    });
  });

  describe("post()", () => {
    const sandbox = sinon.createSandbox();
    afterEach(() => {
      sandbox.restore();
    });
    it("happy", async () => {
      const inputs: Inputs = {
        platform: Platform.CLI,
        projectPath: "./",
      };
      sandbox.stub(OfficeAddinManifest, "modifyManifestFile").resolves();
      const res = await generator.post(context, inputs, "./");
      chai.assert.isTrue(res.isOk());
    });
  });

  describe("childProcessExec()", () => {
    it("should run childProcessExec command success", async function () {
      sinon.stub(childProcess, "exec").yields(null, "test", null);
      chai.assert(await OfficeXMLAddinGenerator.childProcessExec(`echo 'test'`), "test");
    });
  });
});
