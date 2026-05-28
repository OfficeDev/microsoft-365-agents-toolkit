import * as chai from "chai";
import sinon from "ts-sinon";
import { Spec } from "../../../../src/officeChat/common/skills/spec";
import { CancellationToken, LanguageModelChatMessage, LanguageModelChatMessageRole } from "vscode";
import { ExecutionResultEnum } from "../../../../src/officeChat/common/skills/executionResultEnum";
import { projectCreator } from "../../../../src/officeChat/common/skills/projectCreator";
import path = require("path");
import * as helper from "../../../../src/chat/commands/create/helper";
import fs from "fs-extra";
import * as vscode from "vscode";
import { SampleData } from "../../../../src/officeChat/common/samples/sampleData";
import { CreateProjectResult, ok } from "@microsoft/teamsfx-api";
import { core } from "../../../../src/globalVariables";

describe("projectCreator", () => {
  let invokeParametersInit: () => any;
  const sandbox = sinon.createSandbox();

  beforeEach(() => {
    invokeParametersInit = function () {
      const spec = new Spec("some user input");
      spec.taskSummary = "some task summary";
      spec.sections = ["section1", "section2"];
      spec.inspires = ["inspire1", "inspire2"];
      spec.resources = ["resource1", "resource2"];
      spec.appendix = {
        host: "some host",
        codeSnippet: "some code",
        codeExplanation: "some explanation",
        codeTaskBreakdown: ["task1", "task2"],
        codeSample: "",
        apiDeclarationsReference: new Map<string, SampleData>(),
        isCustomFunction: false,
        telemetryData: {
          requestId: "Id",
          isHarmful: false,
          relatedSampleName: ["sample1", "sample2"],
          chatMessages: [
            new LanguageModelChatMessage(LanguageModelChatMessageRole.User, "requestMessage1"),
            new LanguageModelChatMessage(LanguageModelChatMessageRole.User, "requestMessage2"),
          ],
          responseChatMessages: [
            new LanguageModelChatMessage(LanguageModelChatMessageRole.User, "responseMessage1"),
            new LanguageModelChatMessage(LanguageModelChatMessageRole.User, "responseMessage2"),
          ],
          properties: { property1: "value1", property2: "value2" },
          measurements: { measurement1: 1, measurement2: 2 },
        },
        complexity: 0,
        shouldContinue: false,
      };

      const model: LanguageModelChatMessage = {
        role: LanguageModelChatMessageRole.User,
        content: "",
        name: undefined,
      };

      const fakeResponse = {
        markdown: sandbox.stub(),
        anchor: sandbox.stub(),
        button: sandbox.stub(),
        filetree: sandbox.stub(),
        progress: sandbox.stub(),
        reference: sandbox.stub(),
        push: sandbox.stub(),
      } as unknown as vscode.ChatResponseStream;

      const fakeToken: CancellationToken = {
        isCancellationRequested: false,
        onCancellationRequested: sandbox.stub(),
      };

      return { spec, model, fakeResponse, fakeToken };
    };
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("constructor", () => {
    const project_creator = new projectCreator();

    chai.assert.isNotNull(project_creator);
    chai.assert.equal(project_creator.name, "Project Creator");
    chai.assert.equal(project_creator.capability, "Create a new project template");
  });

  it("canInvoke returns true", () => {
    const project_creator = new projectCreator();
    const spec = new Spec("Some user input");
    spec.taskSummary = "Some task summary";
    spec.sections = ["section1", "section2"];
    spec.inspires = ["inspire1", "inspire2"];
    spec.resources = ["resource1", "resource2"];
    spec.appendix = {
      host: "Some host",
      codeSnippet: "Some code snippet",
      codeExplanation: "Some code explanation",
      codeTaskBreakdown: ["task1", "task2"],
      codeSample: "",
      apiDeclarationsReference: new Map<string, SampleData>(),
      isCustomFunction: true,
      telemetryData: {
        requestId: "Id",
        isHarmful: false,
        relatedSampleName: ["sample1", "sample2"],
        chatMessages: [
          new LanguageModelChatMessage(LanguageModelChatMessageRole.User, "requestMessage1"),
          new LanguageModelChatMessage(LanguageModelChatMessageRole.User, "requestMessage2"),
        ],
        responseChatMessages: [
          new LanguageModelChatMessage(LanguageModelChatMessageRole.User, "responseMessage1"),
          new LanguageModelChatMessage(LanguageModelChatMessageRole.User, "responseMessage2"),
        ],
        properties: {
          property1: "value1",
          property2: "value2",
        },
        measurements: {
          measurement1: 1,
          measurement2: 2,
        },
      },
      complexity: 3,
      shouldContinue: false,
    };

    const result = project_creator.canInvoke(spec);
    chai.assert.isTrue(result);
  });

  it("Invoke - mergeCFCode - no write file error", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const project_creator = new projectCreator();
    spec.appendix.isCustomFunction = true;

    //buildProjectFromSpec
    sandbox.stub(vscode.commands, "executeCommand");

    /* mergeCFCode */
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "readFile").resolves(Buffer.from(""));
    sandbox.stub(fs, "writeFile").resolves();
    /* traverseFiles */
    sandbox.stub(path, "relative").returns("relative path");
    sandbox.stub(helper, "fileTreeAdd");
    const res: CreateProjectResult = { projectPath: path.join("testFolder", "test") };
    sandbox.stub(core, "createProjectByCustomizedGenerator").resolves(ok(res));

    const lstatSyncStub = sandbox.stub(fs, "lstatSync");

    const fakeDirentfiles0 = ["dir1"];
    ///store original readdirSync
    const originalReaddirSync = Reflect.get(fs, "readdirSync");
    //fake readdirSync
    Reflect.set(fs, "readdirSync", (dir: string) => {
      return fakeDirentfiles0;
    });

    const fakeStats0 = {
      isDirectory: () => true,
    } as fs.Stats;
    lstatSyncStub.onCall(0).returns(fakeStats0);

    const fakeStats1 = {
      isDirectory: () => false,
    } as fs.Stats;
    lstatSyncStub.onCall(1).returns(fakeStats1);

    const result = await project_creator.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.Success);
    chai.expect(spec).to.equal(spec);

    //restore reflect functions
    Reflect.set(fs, "readdirSync", originalReaddirSync);
  });

  it("Invoke - mergeCFCode - write file error", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const project_creator = new projectCreator();
    spec.appendix.isCustomFunction = true;

    //buildProjectFromSpec
    sandbox.stub(vscode.commands, "executeCommand");

    /* mergeCFCode */
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "readFile").resolves(Buffer.from(""));
    sandbox.stub(fs, "writeFile").rejects(Error("write file error"));

    /* traverseFiles */
    sandbox.stub(path, "relative").returns("relative path");
    sandbox.stub(helper, "fileTreeAdd");
    const res: CreateProjectResult = { projectPath: path.join("testFolder", "test") };
    sandbox.stub(core, "createProjectByCustomizedGenerator").resolves(ok(res));

    const lstatSyncStub = sandbox.stub(fs, "lstatSync");

    const fakeDirentfiles0 = ["dir1"];
    ///store original readdirSync
    const originalReaddirSync = Reflect.get(fs, "readdirSync");
    //fake readdirSync
    Reflect.set(fs, "readdirSync", (dir: string) => {
      return fakeDirentfiles0;
    });

    const fakeStats0 = {
      isDirectory: () => true,
    } as fs.Stats;
    lstatSyncStub.onCall(0).returns(fakeStats0);

    const fakeStats1 = {
      isDirectory: () => false,
    } as fs.Stats;
    lstatSyncStub.onCall(1).returns(fakeStats1);

    try {
      await project_creator.invoke(model, fakeResponse, fakeToken, spec);
      chai.assert.fail("should not reach here");
    } catch (error) {
      chai.assert.strictEqual((error as Error).message, "Failed to merge the CF project.");
    }

    //restore reflect functions
    Reflect.set(fs, "readdirSync", originalReaddirSync);
  });

  it("Invoke - mergeTaskpaneCode - no write file error", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const project_creator = new projectCreator();
    spec.appendix.isCustomFunction = false;

    //buildProjectFromSpec
    sandbox.stub(vscode.commands, "executeCommand");

    /* mergeTaskpaneCode */
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "readFile").resolves(Buffer.from(""));
    sandbox.stub(fs, "writeFile").resolves();

    /* traverseFiles */
    sandbox.stub(path, "relative").returns("relative path");
    sandbox.stub(helper, "fileTreeAdd");
    const res: CreateProjectResult = { projectPath: path.join("testFolder", "test") };
    sandbox.stub(core, "createProjectByCustomizedGenerator").resolves(ok(res));

    const lstatSyncStub = sandbox.stub(fs, "lstatSync");

    const fakeDirentfiles0 = ["dir1"];
    ///store original readdirSync
    const originalReaddirSync = Reflect.get(fs, "readdirSync");
    //fake readdirSync
    Reflect.set(fs, "readdirSync", (dir: string) => {
      return fakeDirentfiles0;
    });

    const fakeStats0 = {
      isDirectory: () => true,
    } as fs.Stats;
    lstatSyncStub.onCall(0).returns(fakeStats0);

    const fakeStats1 = {
      isDirectory: () => false,
    } as fs.Stats;
    lstatSyncStub.onCall(1).returns(fakeStats1);

    const result = await project_creator.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.Success);
    chai.expect(spec).to.equal(spec);

    //restore reflect functions
    Reflect.set(fs, "readdirSync", originalReaddirSync);
  });

  it("Invoke - mergeTaskpaneCode - write file error", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const project_creator = new projectCreator();
    spec.appendix.isCustomFunction = false;

    //buildProjectFromSpec
    sandbox.stub(vscode.commands, "executeCommand");

    /* mergeTaskpaneCode */
    ///store original fs
    sandbox.stub(fs, "ensureDir").resolves();
    sandbox.stub(fs, "readFile").resolves(Buffer.from(""));
    sandbox.stub(fs, "writeFile").rejects(Error("write file error"));

    /* traverseFiles */
    sandbox.stub(path, "relative").returns("relative path");
    sandbox.stub(helper, "fileTreeAdd");
    const res: CreateProjectResult = { projectPath: path.join("testFolder", "test") };
    sandbox.stub(core, "createProjectByCustomizedGenerator").resolves(ok(res));

    const lstatSyncStub = sandbox.stub(fs, "lstatSync");

    const fakeDirentfiles0 = ["dir1"];
    ///store original readdirSync
    const originalReaddirSync = Reflect.get(fs, "readdirSync");
    //fake readdirSync
    Reflect.set(fs, "readdirSync", (dir: string) => {
      return fakeDirentfiles0;
    });

    const fakeStats0 = {
      isDirectory: () => true,
    } as fs.Stats;
    lstatSyncStub.onCall(0).returns(fakeStats0);

    const fakeStats1 = {
      isDirectory: () => false,
    } as fs.Stats;
    lstatSyncStub.onCall(1).returns(fakeStats1);

    try {
      await project_creator.invoke(model, fakeResponse, fakeToken, spec);
      chai.assert.fail("should not reach here");
    } catch (error) {
      chai.assert.strictEqual((error as Error).message, "Failed to merge the taskpane project.");
    }

    //restore reflect functions
    Reflect.set(fs, "readdirSync", originalReaddirSync);
  });
});
