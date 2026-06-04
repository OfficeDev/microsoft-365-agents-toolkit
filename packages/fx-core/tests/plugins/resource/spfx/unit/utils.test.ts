/* eslint-disable @typescript-eslint/no-non-null-asserted-optional-chain */
import { FuncValidation, Inputs, Platform, Stage, TextInputQuestion } from "@microsoft/teamsfx-api";
import * as chai from "chai";
import fs from "fs-extra";
import mockedEnv, { RestoreFn } from "mocked-env";
import * as path from "path";
import * as sinon from "sinon";
import { getLocalizedString } from "../../../../../src/common/localizeUtils";
import { cpUtils } from "../../../../../src/component/deps-checker/util/cpUtils";
import { Utils } from "../../../../../src/component/generator/spfx/utils/utils";
import {
  QuestionNames,
  SPFxWebpartNameQuestion,
  appNameQuestion,
} from "../../../../../src/question";

describe("utils", () => {
  afterEach(async () => {
    sinon.restore();
  });

  describe("webpart name", () => {
    let mockedEnvRestore: RestoreFn;
    const previousInputs: Inputs = { platform: Platform.VSCode };
    beforeEach(() => {
      mockedEnvRestore = mockedEnv({ TEAMSFX_V3: "false" });
      previousInputs["projectPath"] = "c:\\testPath";
    });
    afterEach(() => {
      mockedEnvRestore();
    });

    it("Returns undefined when web part name not duplicated in create stage", async () => {
      previousInputs.stage = Stage.create;

      const res = await (
        (SPFxWebpartNameQuestion() as TextInputQuestion).validation! as FuncValidation<string>
      ).validFunc("helloworld", previousInputs);

      chai.expect(res).equal(undefined);
    });

    it("Returns not match pattern when web part name pattern mismatch in create stage", async () => {
      previousInputs.stage = Stage.create;
      const input = "1";

      const res = await (
        (SPFxWebpartNameQuestion() as TextInputQuestion).validation! as FuncValidation<string>
      ).validFunc(input, previousInputs);

      chai
        .expect(res)
        .equal(
          getLocalizedString(
            "plugins.spfx.questions.webpartName.error.notMatch",
            input,
            "^[a-zA-Z_][a-zA-Z0-9_]*$"
          )
        );
    });

    it("Returns undefined when web part name pattern duplicated in create stage", async () => {
      previousInputs.stage = Stage.create;
      const input = "helloworld";
      sinon.stub(fs, "pathExists").callsFake(async (directory) => {
        if (
          directory === path.join(previousInputs?.projectPath!, "SPFx", "src", "webparts", input)
        ) {
          return true;
        }
      });

      const res = await (
        (SPFxWebpartNameQuestion() as TextInputQuestion).validation! as FuncValidation<string>
      ).validFunc(input, previousInputs);

      chai.expect(res).equal(undefined);
      sinon.restore();
    });

    it("Returns undefined when web part name not duplicated in addFeature stage", async () => {
      previousInputs.stage = Stage.addFeature;
      const input = "helloworld";
      sinon.stub(fs, "pathExists").callsFake(async (directory) => {
        if (
          directory === path.join(previousInputs?.projectPath!, "SPFx", "src", "webparts", input)
        ) {
          return false;
        }
      });
      previousInputs["spfx-folder"] = path.join(previousInputs?.projectPath!, "SPFx");
      const res = await (
        (SPFxWebpartNameQuestion() as TextInputQuestion).validation! as FuncValidation<string>
      ).validFunc(input, previousInputs);

      chai.expect(res).equal(undefined);
      sinon.restore();
    });

    it("Returns not match pattern when web part name pattern mismatch in addFeature stage", async () => {
      previousInputs.stage = Stage.addFeature;
      const input = "1";

      const res = await (
        (SPFxWebpartNameQuestion() as TextInputQuestion).validation! as FuncValidation<string>
      ).validFunc(input, previousInputs);

      chai
        .expect(res)
        .equal(
          getLocalizedString(
            "plugins.spfx.questions.webpartName.error.notMatch",
            input,
            "^[a-zA-Z_][a-zA-Z0-9_]*$"
          )
        );
    });

    it("Returns duplicated when web part name pattern duplicated in addFeature stage", async () => {
      previousInputs.stage = Stage.addFeature;
      const input = "helloworld";
      sinon.stub(fs, "pathExists").callsFake(async (directory) => {
        if (
          directory === path.join(previousInputs?.projectPath!, "SPFx", "src", "webparts", input)
        ) {
          return true;
        }
      });
      previousInputs["spfx-folder"] = path.join(previousInputs?.projectPath!, "SPFx");
      const res = await (
        (SPFxWebpartNameQuestion() as TextInputQuestion).validation! as FuncValidation<string>
      ).validFunc(input, previousInputs);

      chai
        .expect(res)
        .equal(
          getLocalizedString(
            "plugins.spfx.questions.webpartName.error.duplicate",
            path.join(previousInputs?.projectPath!, "SPFx", "src", "webparts", input)
          )
        );
      sinon.restore();
    });
  });

  it("findLatestVersion: exeute commmand error with undefined logger", async () => {
    sinon.stub(cpUtils, "executeCommand").throws("run command error");

    const res = await Utils.findLatestVersion(undefined, "name", 0);
    chai.expect(res).to.be.undefined;
  });

  it("findGloballyInstalledVersion: exeute commmand error with undefined logger", async () => {
    sinon.stub(cpUtils, "executeCommand").throws("run command error");
    let error = undefined;

    try {
      await Utils.findGloballyInstalledVersion(undefined, "name", 0);
    } catch (e) {
      error = e;
    }
    chai.expect(error).not.undefined;
  });

  it("findGloballyInstalledVersion: exeute commmand error but not throw error", async () => {
    sinon.stub(cpUtils, "executeCommand").throws("run command error");

    const res = await Utils.findGloballyInstalledVersion(undefined, "name", 0, false);

    chai.expect(res).to.be.undefined;
  });

  it("Returns solution name as default app name", async () => {
    const inputs: Inputs = {
      platform: Platform.VSCode,
      folder: "c:\\testApp",
      [QuestionNames.SPFxFolder]: "c:\\test",
      [QuestionNames.SPFxSolution]: "import",
    };
    sinon
      .stub(fs, "readJson")
      .resolves({ "@microsoft/generator-sharepoint": { solutionName: "fakedSolutionName" } });
    sinon.stub(fs, "pathExists").resolves(true);

    const defaultName = await (appNameQuestion() as any).default(inputs);
    chai.expect(defaultName).equal("fakedSolutionName");
  });

  describe("configure", () => {
    it("replaces content in a single file when path is a file", async () => {
      const filePath = "c:\\test\\config.json";
      sinon.stub(fs, "lstatSync").returns({ isFile: () => true } as any);
      sinon.stub(fs, "readFile").resolves(Buffer.from("hello OLD world"));
      const writeStub = sinon.stub(fs, "writeFile").resolves();

      await Utils.configure(filePath, new Map([["OLD", "NEW"]]));

      chai.expect(writeStub.calledOnce).to.be.true;
      chai.expect(writeStub.firstCall.args[1]).to.equal("hello NEW world");
    });

    it("replaces content in files found by glob when path is a directory", async () => {
      const dirPath = path.join(__dirname, "tmp-configure-dir");
      const fakeFile = path.join(dirPath, "src", "file.json");
      await fs.ensureDir(path.dirname(fakeFile));
      await fs.writeFile(fakeFile, "value: OLD");

      await Utils.configure(dirPath, new Map([["OLD", "NEW"]]));

      const content = (await fs.readFile(fakeFile)).toString();
      chai.expect(content).to.equal("value: NEW");
      await fs.remove(dirPath);
    });
  });
});
