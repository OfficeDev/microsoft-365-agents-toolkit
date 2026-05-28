import * as chai from "chai";
import sinon from "ts-sinon";
import { Spec } from "../../../../src/officeChat/common/skills/spec";
import {
  CancellationToken,
  ChatResponseStream,
  LanguageModelChatMessage,
  LanguageModelChatMessageRole,
} from "vscode";
import * as utils from "../../../../src/officeChat/utils";
import { ExecutionResultEnum } from "../../../../src/officeChat/common/skills/executionResultEnum";
import { Printer } from "../../../../src/officeChat/common/skills/printer";
import { SampleData } from "../../../../src/officeChat/common/samples/sampleData";

describe("printer", () => {
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
      } as unknown as ChatResponseStream;

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
    const printer = new Printer();

    chai.assert.isNotNull(printer);
    chai.assert.equal(printer.name, "printer");
    chai.assert.equal(printer.capability, "Print the output in a readable format to user");
  });

  it("canInvoke returns true", () => {
    const printer = new Printer();
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

    const result = printer.canInvoke(spec);
    chai.assert.isTrue(result);
  });

  it("Invoke failure", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const printer = new Printer();

    sandbox.stub(utils, "isOutputHarmful").resolves(false);

    const result = await printer.invoke(model, fakeResponse, fakeToken, spec);
    chai.expect(result.result).to.equal(ExecutionResultEnum.Success);
    chai.expect(spec).to.equal(spec);
  });

  it("Invoke Success", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const printer = new Printer();

    sandbox.stub(utils, "isOutputHarmful").resolves(true);

    const result = await printer.invoke(model, fakeResponse, fakeToken, spec);
    chai.expect(result.result).to.equal(ExecutionResultEnum.Failure);
    chai.expect(spec).to.equal(spec);
  });
});
