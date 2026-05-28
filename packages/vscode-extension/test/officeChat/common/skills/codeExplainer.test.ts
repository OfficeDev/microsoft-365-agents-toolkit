import * as chai from "chai";
import sinon from "ts-sinon";
import { Explainer } from "../../../../src/officeChat/common/skills/codeExplainer";
import { Spec } from "../../../../src/officeChat/common/skills/spec";
import {
  CancellationToken,
  ChatResponseStream,
  LanguageModelChatMessage,
  LanguageModelChatMessageRole,
} from "vscode";
import * as utils from "../../../../src/chat/utils";
import { ExecutionResultEnum } from "../../../../src/officeChat/common/skills/executionResultEnum";
import { SampleData } from "../../../../src/officeChat/common/samples/sampleData";

describe("CodeExplainer", () => {
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
    const codeExplainer = new Explainer();

    chai.assert.isNotNull(codeExplainer);
    chai.assert.equal(codeExplainer.name, "Explainer");
    chai.assert.equal(codeExplainer.capability, "Explain code snippet");
  });

  it("canInvoke returns true", () => {
    const codeExplainer = new Explainer();
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

    const result = codeExplainer.canInvoke(spec);
    chai.assert.isTrue(result);
  });

  it("Invoke failure", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeExplainer = new Explainer();

    sandbox.stub(utils, "getCopilotResponseAsString").resolves(undefined);

    const result = await codeExplainer.invoke(model, fakeResponse, fakeToken, spec);
    chai.expect(result.result).to.equal(ExecutionResultEnum.Failure);
    chai.expect(spec).to.equal(spec);
  });

  it("Invoke Success", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeExplainer = new Explainer();

    sandbox.stub(utils, "getCopilotResponseAsString").resolves("Some response");

    const result = await codeExplainer.invoke(model, fakeResponse, fakeToken, spec);
    chai.expect(result.result).to.equal(ExecutionResultEnum.Success);
    chai.expect(spec).to.equal(spec);
  });
});
