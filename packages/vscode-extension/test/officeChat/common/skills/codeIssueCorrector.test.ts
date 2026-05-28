import * as chai from "chai";
import { CodeIssueCorrector } from "../../../../src/officeChat/common/skills/codeIssueCorrector";
import * as sinon from "sinon";
import { Spec } from "../../../../src/officeChat/common/skills/spec";
import * as utils from "../../../../src/chat/utils";
import {
  CodeIssueDetector,
  DetectionResult,
} from "../../../../src/officeChat/common/skills/codeIssueDetector";
import {
  CancellationToken,
  ChatResponseStream,
  LanguageModelChatMessage,
  LanguageModelChatMessageRole,
} from "vscode";
import { ExecutionResultEnum } from "../../../../src/officeChat/common/skills/executionResultEnum";
import { SampleProvider } from "../../../../src/officeChat/common/samples/sampleProvider";
import { SampleData } from "../../../../src/officeChat/common/samples/sampleData";

describe("CodeIssueCorrector", () => {
  const sandbox = sinon.createSandbox();
  let invokeParametersInit: () => any;

  beforeEach(() => {
    invokeParametersInit = function () {
      const spec = new Spec("some user input");
      spec.taskSummary = "some task summary";
      spec.sections = ["section1", "section2"];
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
    const codeIssueCorrector = new CodeIssueCorrector();

    chai.assert.isNotNull(codeIssueCorrector);
    chai.assert.equal(codeIssueCorrector.name, "codeIssueCorrector");
    chai.assert.equal(codeIssueCorrector.capability, "Fix code issues");
  });

  it("canInvoke returns true", async () => {
    const corrector = new CodeIssueCorrector();
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

    const result = await corrector.canInvoke(spec);
    chai.assert.isTrue(result);
  });

  it("fixIssueAsync no error return codeSnippet", async () => {
    const sampleCodeLong =
      `Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.
    To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries.
    Themes and styles also help keep your document coordinated. When you click Design and choose a new Theme, the pictures, charts, and SmartArt graphics change to match your new theme. When you apply styles, your headings change to match the new theme.
    Save time in Word with new buttons that show up where you need them. To change the way a picture fits in your document, click it and a button for layout options appears next to it. When you work on a table, click where you want to add a row or a column, and then click the plus sign.
    Reading is easier, too, in the new Reading view. You can collapse parts of the document and focus on the text you want. If you need to stop reading before you reach the end, Word remembers where you left off - even on another device.
    `.repeat(20);
    const corrector = new CodeIssueCorrector();
    const fakeLanguageModelChatSystemMessage: LanguageModelChatMessage = {
      role: LanguageModelChatMessageRole.User,
      content: "some sample message",
      name: undefined,
    };
    const fakeSampleCodeLanguageModelChatSystemMessage: LanguageModelChatMessage = {
      role: LanguageModelChatMessageRole.User,
      content: sampleCodeLong,
      name: undefined,
    };
    const spec = new Spec("some user input");
    sandbox
      .stub(utils, "countMessagesTokens")
      .onFirstCall()
      .returns(10000)
      .onSecondCall()
      .returns(100);
    const result = await corrector.fixIssueAsync(
      {
        isCancellationRequested: false,
        onCancellationRequested: undefined as any, // Assign undefined
      }, // CancellationToken
      "Excel", // host
      false, // isCustomFunctions
      "original code snippet", // codeSnippet
      ["step1", "step2"], // substeps
      [], // errorMessages
      ["warning1", "warning2"], // warningMessage
      [], // historical errors
      "additional info", // additionalInfo
      "copilot-gpt-3.5-turbo", // model
      fakeLanguageModelChatSystemMessage,
      fakeSampleCodeLanguageModelChatSystemMessage,
      spec
    );

    chai.assert.equal(result, "original code snippet");
  });

  it("fixIssueAsync error with the LLM output and Excel host, isCustomFunctions false", async () => {
    const sampleCodeLong =
      `Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.
    To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries.
    Themes and styles also help keep your document coordinated. When you click Design and choose a new Theme, the pictures, charts, and SmartArt graphics change to match your new theme. When you apply styles, your headings change to match the new theme.
    Save time in Word with new buttons that show up where you need them. To change the way a picture fits in your document, click it and a button for layout options appears next to it. When you work on a table, click where you want to add a row or a column, and then click the plus sign.
    Reading is easier, too, in the new Reading view. You can collapse parts of the document and focus on the text you want. If you need to stop reading before you reach the end, Word remembers where you left off - even on another device.
    `.repeat(20);
    const corrector = new CodeIssueCorrector();
    const fakeLanguageModelChatSystemMessage: LanguageModelChatMessage = {
      role: LanguageModelChatMessageRole.User,
      content: "some sample message",
      name: undefined,
    };
    const fakeSampleCodeLanguageModelChatSystemMessage: LanguageModelChatMessage = {
      role: LanguageModelChatMessageRole.User,
      content: sampleCodeLong,
      name: undefined,
    };

    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    const spec = new Spec("some user input");
    getCopilotResponseAsStringStub.returns(
      Promise.resolve("```typescript\nfixed code snippet\n```")
    );
    sandbox.stub(console, "log");
    sandbox.stub(console, "error");
    sandbox
      .stub(utils, "countMessagesTokens")
      .onFirstCall()
      .returns(10000)
      .onSecondCall()
      .returns(100);
    // sandbox.stub(utils, "countMessageTokens").returns(100);
    sandbox.stub(RegExp.prototype, "exec").returns(null);

    const result = await corrector.fixIssueAsync(
      {
        isCancellationRequested: false,
        onCancellationRequested: undefined as any, // Assign undefined
      }, // CancellationToken
      "Excel", // host
      false, // isCustomFunctions
      "original code snippet", // codeSnippet
      ["step1", "step2"], // substeps
      ["error1", "error2"], // errorMessages
      ["warning1", "warning2"], // warningMessage
      [], // historical errors
      "additional info", // additionalInfo
      "copilot-gpt-3.5-turbo", // model
      fakeLanguageModelChatSystemMessage,
      fakeSampleCodeLanguageModelChatSystemMessage,
      spec
    );

    chai.assert.equal(result, null);
  });

  it("fixIssueAsync error with the LLM output and Excel host, isCustomFunctions true", async () => {
    const sampleCodeLong =
      `Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.
    To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries.
    Themes and styles also help keep your document coordinated. When you click Design and choose a new Theme, the pictures, charts, and SmartArt graphics change to match your new theme. When you apply styles, your headings change to match the new theme.
    Save time in Word with new buttons that show up where you need them. To change the way a picture fits in your document, click it and a button for layout options appears next to it. When you work on a table, click where you want to add a row or a column, and then click the plus sign.
    Reading is easier, too, in the new Reading view. You can collapse parts of the document and focus on the text you want. If you need to stop reading before you reach the end, Word remembers where you left off - even on another device.
    `.repeat(20);
    const corrector = new CodeIssueCorrector();
    const fakeLanguageModelChatSystemMessage: LanguageModelChatMessage = {
      role: LanguageModelChatMessageRole.User,
      content: "some sample message",
      name: undefined,
    };
    const fakeSampleCodeLanguageModelChatSystemMessage: LanguageModelChatMessage = {
      role: LanguageModelChatMessageRole.User,
      content: sampleCodeLong,
      name: undefined,
    };

    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    const spec = new Spec("some user input");
    getCopilotResponseAsStringStub.returns(
      Promise.resolve("```typescript\nfixed code snippet\n```")
    );
    sandbox.stub(console, "log");
    sandbox.stub(console, "error");
    sandbox
      .stub(utils, "countMessagesTokens")
      .onFirstCall()
      .returns(10000)
      .onSecondCall()
      .returns(100);
    // sandbox.stub(utils, "countMessageTokens").returns(100);
    sandbox.stub(RegExp.prototype, "exec").returns(null);

    const result = await corrector.fixIssueAsync(
      {
        isCancellationRequested: false,
        onCancellationRequested: undefined as any, // Assign undefined
      }, // CancellationToken
      "Excel", // host
      true, // isCustomFunctions
      "original code snippet", // codeSnippet
      ["step1", "step2"], // substeps
      ["error1", "error2"], // errorMessages
      ["warning1", "warning2"], // warningMessage
      [], // historical errors
      "additional info", // additionalInfo
      "copilot-gpt-3.5-turbo", // model
      fakeLanguageModelChatSystemMessage, // sampleMessage
      fakeSampleCodeLanguageModelChatSystemMessage,
      spec
    );

    chai.assert.equal(result, null);
  });

  it("fixIssueAsync error with the LLM output and other host", async () => {
    const sampleCodeLong =
      `Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.
    To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries.
    Themes and styles also help keep your document coordinated. When you click Design and choose a new Theme, the pictures, charts, and SmartArt graphics change to match your new theme. When you apply styles, your headings change to match the new theme.
    Save time in Word with new buttons that show up where you need them. To change the way a picture fits in your document, click it and a button for layout options appears next to it. When you work on a table, click where you want to add a row or a column, and then click the plus sign.
    Reading is easier, too, in the new Reading view. You can collapse parts of the document and focus on the text you want. If you need to stop reading before you reach the end, Word remembers where you left off - even on another device.
    `.repeat(20);
    const corrector = new CodeIssueCorrector();
    const fakeLanguageModelChatSystemMessage: LanguageModelChatMessage = {
      role: LanguageModelChatMessageRole.User,
      content: "some sample message",
      name: undefined,
    };
    const fakeSampleCodeLanguageModelChatSystemMessage: LanguageModelChatMessage = {
      role: LanguageModelChatMessageRole.User,
      content: sampleCodeLong,
      name: undefined,
    };

    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    const spec = new Spec("some user input");
    getCopilotResponseAsStringStub.returns(
      Promise.resolve("```typescript\nfixed code snippet\n```")
    );
    sandbox.stub(console, "log");
    sandbox.stub(console, "error");
    sandbox
      .stub(utils, "countMessagesTokens")
      .onFirstCall()
      .returns(10000)
      .onSecondCall()
      .returns(100);
    // sandbox.stub(utils, "countMessageTokens").returns(100);
    sandbox.stub(RegExp.prototype, "exec").returns(null);

    const result = await corrector.fixIssueAsync(
      {
        isCancellationRequested: false,
        onCancellationRequested: undefined as any, // Assign undefined
      }, // CancellationToken
      "Word", // host
      false, // isCustomFunctions
      "original code snippet", // codeSnippet
      ["step1", "step2"], // substeps
      ["error1", "error2"], // errorMessages
      ["warning1", "warning2"], // warningMessage
      [], // historical errors
      "additional info", // additionalInfo
      "copilot-gpt-3.5-turbo", // model
      fakeLanguageModelChatSystemMessage,
      fakeSampleCodeLanguageModelChatSystemMessage,
      spec
    );

    chai.assert.equal(result, null);
  });

  it("fixIssueAsync error with code length reduced too much", async () => {
    const corrector = new CodeIssueCorrector();
    const fakeLanguageModelChatSystemMessage: LanguageModelChatMessage = {
      role: LanguageModelChatMessageRole.User,
      content: "some sample message",
      name: undefined,
    };

    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    const spec = new Spec("some user input");
    getCopilotResponseAsStringStub.returns(
      Promise.resolve("```typescript\nfixed code snippet\n```")
    );
    sandbox.stub(console, "log");
    sandbox.stub(console, "error");
    sandbox.stub(console, "debug");
    sandbox.stub(utils, "countMessagesTokens").returns(100);
    sandbox.stub(utils, "countMessageTokens").returns(100);
    sandbox.stub(RegExp.prototype, "exec").returns(["++"] as RegExpExecArray);

    const result = await corrector.fixIssueAsync(
      {
        isCancellationRequested: false,
        onCancellationRequested: undefined as any, // Assign undefined
      }, // CancellationToken
      "Word", // host
      false, // isCustomFunctions
      "++++++++++", // codeSnippet
      ["step1", "step2"], // substeps
      ["error1", "error2"], // errorMessages
      ["warning1", "warning2"], // warningMessage
      [], // historical errors
      "additional info", // additionalInfo
      "copilot-gpt-3.5-turbo", // model
      fakeLanguageModelChatSystemMessage,
      fakeLanguageModelChatSystemMessage,
      spec
    );

    chai.assert.equal(result, null);
  });

  it("fixIssueAsync return newCodeStr", async () => {
    const corrector = new CodeIssueCorrector();
    const fakeLanguageModelChatSystemMessage: LanguageModelChatMessage = {
      role: LanguageModelChatMessageRole.User,
      content: "some sample message",
      name: undefined,
    };

    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    const spec = new Spec("some user input");
    getCopilotResponseAsStringStub.returns(
      Promise.resolve("```typescript\nfixed code snippet\n```")
    );
    sandbox.stub(console, "log");
    sandbox.stub(console, "error");
    sandbox.stub(utils, "countMessagesTokens").returns(100);
    sandbox.stub(utils, "countMessageTokens").returns(100);
    sandbox.stub(RegExp.prototype, "exec").returns(["++++++++"] as RegExpExecArray);

    const result = await corrector.fixIssueAsync(
      {
        isCancellationRequested: false,
        onCancellationRequested: undefined as any, // Assign undefined
      }, // CancellationToken
      "Word", // host
      false, // isCustomFunctions
      "++++++++++", // codeSnippet
      ["step1", "step2"], // substeps
      ["error1", "error2"], // errorMessages
      ["warning1", "warning2"], // warningMessage
      [], // historical errors
      "additional info", // additionalInfo
      "copilot-gpt-3.5-turbo", // model
      fakeLanguageModelChatSystemMessage,
      fakeLanguageModelChatSystemMessage,
      spec
    );

    chai.assert.equal(result, "++++++++");
  });

  it("invoke return success when no issues are found in baseline with complexity < 25", async () => {
    const corrector = new CodeIssueCorrector();
    const detector = CodeIssueDetector.getInstance();
    const detectionResult = new DetectionResult();

    sandbox.stub(detector, "detectIssuesAsync").returns(Promise.resolve(detectionResult));

    sandbox.stub(console, "debug");

    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    spec.appendix.complexity = 10;
    spec.appendix.codeSample = "some code sample";
    spec.appendix.apiDeclarationsReference.set(
      "definition",
      new SampleData("key1", "docLink", "sample", "description", "definition", "usage")
    );

    const result = await corrector.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.Success);
    chai.expect(result.spec).to.equal(spec);
  });

  it("invoke return success when no issues are found in baseline with complexity < 50", async () => {
    const corrector = new CodeIssueCorrector();
    const detector = CodeIssueDetector.getInstance();
    const detectionResult = new DetectionResult();

    sandbox.stub(detector, "detectIssuesAsync").returns(Promise.resolve(detectionResult));

    sandbox.stub(console, "debug");

    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    spec.appendix.complexity = 30;

    const result = await corrector.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.Success);
    chai.expect(result.spec).to.equal(spec);
  });

  it("invoke return success when no issues are found in baseline with complexity < 75", async () => {
    const corrector = new CodeIssueCorrector();
    const detector = CodeIssueDetector.getInstance();
    const detectionResult = new DetectionResult();

    sandbox.stub(detector, "detectIssuesAsync").returns(Promise.resolve(detectionResult));

    sandbox.stub(console, "debug");

    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    spec.appendix.complexity = 60;

    const result = await corrector.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.Success);
    chai.expect(result.spec).to.equal(spec);
  });

  it("invoke return success when no issues are found in baseline with complexity >= 75", async () => {
    const corrector = new CodeIssueCorrector();
    const detector = CodeIssueDetector.getInstance();
    const detectionResult = new DetectionResult();

    sandbox.stub(detector, "detectIssuesAsync").returns(Promise.resolve(detectionResult));

    sandbox.stub(console, "debug");

    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    spec.appendix.complexity = 80;

    const result = await corrector.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.Success);
    chai.expect(result.spec).to.equal(spec);
  });

  it("invoke return failure low quality code", async () => {
    const corrector = new CodeIssueCorrector();
    const detector = CodeIssueDetector.getInstance();
    const detectionResult = new DetectionResult();
    detectionResult.compileErrors = ["error1", "error2", "error3", "error4", "error5"];

    sandbox.stub(detector, "detectIssuesAsync").returns(Promise.resolve(detectionResult));

    sandbox.stub(console, "debug");

    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    spec.appendix.complexity = 10;

    const result = await corrector.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.FailedAndGoNext);
    chai.expect(result.spec).to.equal(spec);
  });

  it("invoke found issue and self reflection fail fast", async () => {
    const corrector = new CodeIssueCorrector();
    const detector = CodeIssueDetector.getInstance();
    const detectionResult = new DetectionResult();
    detectionResult.compileErrors = ["error1", "error2"];
    detectionResult.runtimeErrors = ["error1"];

    sandbox.stub(detector, "detectIssuesAsync").returns(Promise.resolve(detectionResult));

    sandbox.stub(console, "debug");

    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    spec.appendix.complexity = 80;
    spec.appendix.codeSample = "some code sample";
    spec.appendix.apiDeclarationsReference.set(
      "definition",
      new SampleData("key1", "docLink", "sample", "description", "definition", "usage")
    );
    sandbox.stub(corrector, "fixIssueAsync").returns(Promise.resolve(null));

    const result = await corrector.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.FailedAndGoNext);
    chai.expect(result.spec).to.equal(spec);
  });

  it("invoke found issue and self reflection fail fast, terminateFixIteration codeLengthDelta < 0", async () => {
    const corrector = new CodeIssueCorrector();
    const detector = CodeIssueDetector.getInstance();
    const detectionResult = new DetectionResult();
    detectionResult.compileErrors = ["error1", "error2"];
    detectionResult.runtimeErrors = ["error1"];

    sandbox.stub(console, "debug");

    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    spec.appendix.complexity = 80;

    sandbox.stub(corrector, "fixIssueAsync").returns(Promise.resolve("less"));
    const detectorInstance = CodeIssueDetector.getInstance();
    sandbox.stub(detectorInstance, "detectIssuesAsync").returns(Promise.resolve(detectionResult));

    const result = await corrector.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.FailedAndGoNext);
    chai.expect(result.spec).to.equal(spec);
  });

  it("invoke success", async () => {
    const corrector = new CodeIssueCorrector();
    const detector = CodeIssueDetector.getInstance();
    const detectionResult = new DetectionResult();
    detectionResult.compileErrors = ["error1"];
    detectionResult.runtimeErrors = ["error1"];
    const detectionResultAfterFix = new DetectionResult();
    detectionResultAfterFix.compileErrors = ["error1"];
    detectionResultAfterFix.runtimeErrors = ["error1"];
    const detetionResultIncreaseError = new DetectionResult();
    detetionResultIncreaseError.compileErrors = ["error1", "error2"];
    detetionResultIncreaseError.runtimeErrors = [];
    const detectionResultFinal = new DetectionResult();
    detectionResultFinal.compileErrors = [];
    detectionResultFinal.runtimeErrors = [];

    sandbox.stub(console, "debug");

    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    spec.appendix.complexity = 80;

    const fixIssueStub = sandbox
      .stub(corrector, "fixIssueAsync")
      .returns(Promise.resolve("some more code; await main(); "));
    fixIssueStub.onCall(0).returns(Promise.resolve("less"));
    const detectorInstance = CodeIssueDetector.getInstance();
    const detectIssuesStub = sandbox.stub(detectorInstance, "detectIssuesAsync");

    detectIssuesStub.returns(Promise.resolve(detectionResultFinal));
    detectIssuesStub.onCall(0).returns(Promise.resolve(detectionResult));
    // detectIssuesStub.onCall(1).returns(Promise.resolve(detectionResultAfterFix));
    detectIssuesStub.onCall(1).returns(Promise.resolve(detetionResultIncreaseError));
    // detectIssuesStub.onCall(3).returns(Promise.resolve(detectionResultFinal));
    detectIssuesStub.onCall(2).returns(Promise.resolve(detectionResultFinal));

    const result = await corrector.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.Success);
    chai.expect(result.spec).to.equal(spec);
  });

  it("invoke success with 3 errors", async () => {
    const corrector = new CodeIssueCorrector();
    const detector = CodeIssueDetector.getInstance();
    const detectionResult = new DetectionResult();
    detectionResult.compileErrors = ["error1", "error2", "error3"];
    detectionResult.runtimeErrors = ["error1"];
    const detectionResultAfterFix = new DetectionResult();
    detectionResultAfterFix.compileErrors = ["error1"];
    detectionResultAfterFix.runtimeErrors = ["error1"];
    const detetionResultIncreaseError = new DetectionResult();
    detetionResultIncreaseError.compileErrors = ["error1", "error2"];
    detetionResultIncreaseError.runtimeErrors = [];
    const detectionResultFinal = new DetectionResult();
    detectionResultFinal.compileErrors = [];
    detectionResultFinal.runtimeErrors = [];

    sandbox.stub(console, "debug");

    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    spec.appendix.complexity = 80;

    const fixIssueStub = sandbox
      .stub(corrector, "fixIssueAsync")
      .returns(Promise.resolve("some more code"));
    fixIssueStub.onCall(0).returns(Promise.resolve("less"));
    const detectorInstance = CodeIssueDetector.getInstance();
    const detectIssuesStub = sandbox.stub(detectorInstance, "detectIssuesAsync");
    detectIssuesStub.returns(Promise.resolve(detectionResultFinal));
    detectIssuesStub.onCall(0).returns(Promise.resolve(detectionResult));
    detectIssuesStub.onCall(1).returns(Promise.resolve(detectionResultAfterFix));
    // detectIssuesStub.onCall(2).returns(Promise.resolve(detetionResultIncreaseError));
    // detectIssuesStub.onCall(3).returns(Promise.resolve(detectionResultFinal));
    detectIssuesStub.onCall(2).returns(Promise.resolve(detectionResultFinal));

    const result = await corrector.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.FailedAndGoNext);
    chai.expect(result.spec).to.equal(spec);
  });

  it("invoke success with error increase once", async () => {
    const corrector = new CodeIssueCorrector();
    const detector = CodeIssueDetector.getInstance();
    const detectionResult = new DetectionResult();
    detectionResult.compileErrors = ["error1"];
    detectionResult.runtimeErrors = ["error1"];
    const detectionResultAfterFix = new DetectionResult();
    detectionResultAfterFix.compileErrors = ["error1", "error2", "error3"];
    detectionResultAfterFix.runtimeErrors = ["error1"];
    const detetionResultIncreaseError = new DetectionResult();
    detetionResultIncreaseError.compileErrors = ["error1", "error2"];
    detetionResultIncreaseError.runtimeErrors = [];
    const detectionResultFinal = new DetectionResult();
    detectionResultFinal.compileErrors = [];
    detectionResultFinal.runtimeErrors = [];

    sandbox.stub(console, "debug");

    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    spec.appendix.complexity = 80;

    const fixIssueStub = sandbox
      .stub(corrector, "fixIssueAsync")
      .returns(Promise.resolve("some more code"));
    fixIssueStub.onCall(0).returns(Promise.resolve("less"));
    const detectorInstance = CodeIssueDetector.getInstance();
    const detectIssuesStub = sandbox.stub(detectorInstance, "detectIssuesAsync");
    detectIssuesStub.returns(Promise.resolve(detectionResultFinal));
    detectIssuesStub.onCall(0).returns(Promise.resolve(detectionResult));
    // detectIssuesStub.onCall(1).returns(Promise.resolve(detectionResultAfterFix));
    // detectIssuesStub.onCall(2).returns(Promise.resolve(detetionResultIncreaseError));
    // detectIssuesStub.onCall(3).returns(Promise.resolve(detectionResultFinal));
    detectIssuesStub.onCall(1).returns(Promise.resolve(detectionResultFinal));

    const result = await corrector.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.Success);
    chai.expect(result.spec).to.equal(spec);
  });

  it("invoke partial failed: the first time fix return null and the second time it returns with error", async () => {
    const corrector = new CodeIssueCorrector();
    const detector = CodeIssueDetector.getInstance();
    const detectionResult = new DetectionResult();
    detectionResult.compileErrors = ["error1"];
    detectionResult.runtimeErrors = ["error1"];

    sandbox.stub(console, "debug");

    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    spec.appendix.complexity = 80;

    const fixIssueStub = sandbox
      .stub(corrector, "fixIssueAsync")
      .returns(Promise.resolve("some more code"));
    fixIssueStub.onCall(0).returns(Promise.resolve(null));
    fixIssueStub.onCall(1).returns(Promise.resolve("some more code"));
    const detectorInstance = CodeIssueDetector.getInstance();
    const detectIssuesStub = sandbox.stub(detectorInstance, "detectIssuesAsync");
    detectIssuesStub.returns(Promise.resolve(detectionResult));

    const result = await corrector.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.FailedAndGoNext);
    // chai.expect(result.spec).to.equal(spec);
  });

  it("invoke partial failed: the code fix always returns error", async () => {
    const detectionResult = new DetectionResult();
    detectionResult.compileErrors = ["error1"];
    detectionResult.runtimeErrors = ["error1"];

    sandbox.stub(console, "debug");

    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();

    spec.appendix.complexity = 80;

    const corrector = new CodeIssueCorrector();
    sandbox.stub(corrector, "fixIssueAsync").returns(Promise.resolve("some more code"));
    const detectorInstance = CodeIssueDetector.getInstance();
    sandbox.stub(detectorInstance, "detectIssuesAsync").returns(Promise.resolve(detectionResult));

    const result = await corrector.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.FailedAndGoNext);
  });

  it("invoke success after error increased", async () => {
    const corrector = new CodeIssueCorrector();
    const detectionResult = new DetectionResult();
    detectionResult.compileErrors = ["error1"];
    detectionResult.runtimeErrors = ["error1"];
    const detetionResultIncreaseError = new DetectionResult();
    detetionResultIncreaseError.compileErrors = ["error1", "error2"];
    detetionResultIncreaseError.runtimeErrors = [];
    const detectionResultFinal = new DetectionResult();
    detectionResultFinal.compileErrors = [];
    detectionResultFinal.runtimeErrors = [];

    sandbox.stub(console, "debug");

    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    spec.appendix.complexity = 80;

    const fixIssueStub = sandbox
      .stub(corrector, "fixIssueAsync")
      .returns(Promise.resolve("some more code; await main(); "));
    fixIssueStub.onCall(0).returns(Promise.resolve("some more code; await main();"));
    const detectorInstance = CodeIssueDetector.getInstance();
    const detectIssuesStub = sandbox.stub(detectorInstance, "detectIssuesAsync");

    detectIssuesStub.returns(Promise.resolve(detectionResultFinal));
    detectIssuesStub.onCall(0).returns(Promise.resolve(detectionResult));
    detectIssuesStub.onCall(1).returns(Promise.resolve(detetionResultIncreaseError));
    detectIssuesStub.onCall(2).returns(Promise.resolve(detectionResultFinal));

    const result = await corrector.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect(result.result).to.equal(ExecutionResultEnum.Success);
    chai.expect(result.spec).to.equal(spec);
  });
});
