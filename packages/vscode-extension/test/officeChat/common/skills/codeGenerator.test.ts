import * as chai from "chai";
import sinon from "ts-sinon";
import { Spec } from "../../../../src/officeChat/common/skills/spec";
import {
  CancellationToken,
  ChatResponseStream,
  LanguageModelChatMessage,
  LanguageModelChatMessageRole,
} from "vscode";
import * as utils from "../../../../src/chat/utils";
import { ExecutionResultEnum } from "../../../../src/officeChat/common/skills/executionResultEnum";
import { Printer } from "../../../../src/officeChat/common/skills/printer";
import { CodeGenerator } from "../../../../src/officeChat/common/skills/codeGenerator";
import { SampleProvider } from "../../../../src/officeChat/common/samples/sampleProvider";
import { SampleData } from "../../../../src/officeChat/common/samples/sampleData";

describe("codeGenerator", () => {
  let invokeParametersInit: () => any;
  const sandbox = sinon.createSandbox();

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
    const codeGenerator = new CodeGenerator();

    chai.assert.isNotNull(codeGenerator);
    chai.assert.equal(codeGenerator.name, "Code Generator");
    chai.assert.equal(codeGenerator.capability, "Generate code");
  });

  it("canInvoke returns true", () => {
    const codeGenerator = new CodeGenerator();
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

    const result = codeGenerator.canInvoke(spec);
    chai.assert.isTrue(result);
  });

  it("userAskPreScanningAsync provided empty string, null returned", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();

    sandbox.stub(console, "log");
    sandbox.stub(console, "error");

    const getCopilotResponseStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseStub.resolves("");

    const result = await codeGenerator.userAskPreScanningAsync(spec, fakeToken);

    chai.expect(result).to.equal(null);
  });

  it("userAskPreScanningAsync provided json object, json detected", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();

    sandbox.stub(console, "log");
    sandbox.stub(console, "error");

    const getCopilotResponseStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseStub.resolves(
      JSON.stringify({
        host: "fakeHost",
        shouldContinue: false,
        customFunctions: true,
        complexity: 1,
      })
    );
    const jsonParseStub = sandbox.stub(JSON, "parse");
    const parserResult = {
      host: "fakeHost",
      shouldContinue: false,
      customFunctions: true,
      complexity: 1,
    };
    jsonParseStub.returns(parserResult);

    const result = await codeGenerator.userAskPreScanningAsync(spec, fakeToken);

    chai.expect(result).to.equal(parserResult);
  });

  it("userAskPreScanningAsync provided json object, json detected: not custom function", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();

    sandbox.stub(console, "log");
    sandbox.stub(console, "error");

    const getCopilotResponseStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseStub.resolves(
      JSON.stringify({
        host: "fakeHost",
        shouldContinue: false,
        customFunctions: false,
        complexity: 1,
      })
    );
    const jsonParseStub = sandbox.stub(JSON, "parse");
    const parserResult = {
      host: "fakeHost",
      shouldContinue: false,
      customFunctions: false,
      complexity: 1,
    };
    jsonParseStub.returns(parserResult);

    const result = await codeGenerator.userAskPreScanningAsync(spec, fakeToken);

    chai.expect(result).to.equal(parserResult);
  });

  it("userAskPreScanningAsync provided json with markdown syntax, json detected", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();

    sandbox.stub(console, "log");
    sandbox.stub(console, "error");

    const getCopilotResponseStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseStub.resolves(
      '```json\n{"host": "fakeHost", "shouldContinue": false, "customFunctions": true, "complexity": 1}\n```'
    );
    const jsonParseStub = sandbox.stub(JSON, "parse");
    const parserResult = {
      host: "fakeHost",
      shouldContinue: false,
      customFunctions: true,
      complexity: 1,
    };
    jsonParseStub.returns(parserResult);

    const result = await codeGenerator.userAskPreScanningAsync(spec, fakeToken);

    chai.expect(result).to.equal(parserResult);
  });

  it("userAskPreScanningAsync with invalid json string should not continue, json not detected", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();

    sandbox.stub(console, "log");
    sandbox.stub(console, "error");

    const getCopilotResponseStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseStub.resolves("some random string that is not a JSON object");

    const result = await codeGenerator.userAskPreScanningAsync(spec, fakeToken);

    chai.expect(result).to.equal(null);
  });

  it("userAskBreakdownAsync returns null", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();

    sandbox.stub(utils, "getCopilotResponseAsString").resolves(undefined);

    const result = await codeGenerator.userAskBreakdownAsync(
      fakeToken,
      spec.appendix.complexity,
      true, //isCustomFunction
      spec.appendix.host,
      spec.userInput,
      "Some code sample",
      spec
    );

    chai.expect(result).to.equal(null);
  });

  it("userAskBreakdownAsync returns null with error", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();

    const getCopilotResponseAsStringStub = sandbox
      .stub(utils, "getCopilotResponseAsString")
      .resolves("not valid JSON");
    sandbox.stub(console, "error");

    const result = await codeGenerator.userAskBreakdownAsync(
      fakeToken,
      spec.appendix.complexity,
      spec.appendix.isCustomFunction,
      spec.appendix.host,
      spec.userInput,
      "",
      spec
    );

    chai.expect(result).to.equal(null);
  });

  it("userAskBreakdownAsync with LLM provided json should continue, is customFunctions", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();

    sandbox.stub(console, "error");

    const getCopilotResponseStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseStub.resolves('```json\n{"spec": "fakeSpec", "funcs": ["fakeData1"]}\n```');
    const jsonParseStub = sandbox.stub(JSON, "parse");
    const jsonParseResult = {
      spec: "fakeSpec",
      funcs: ["fakeData1"],
    };
    jsonParseStub.returns(jsonParseResult);

    const result = await codeGenerator.userAskBreakdownAsync(
      fakeToken,
      spec.appendix.complexity,
      spec.appendix.isCustomFunction,
      spec.appendix.host,
      spec.userInput,
      "",
      spec
    );

    jsonParseResult.funcs.filter((task: string) => {
      return !task.includes("'main'");
    });

    chai.expect(result).to.equal(jsonParseResult);
  });

  it("userAskBreakdownAsync with LLM provided json should continue - complex task, is customFunctions", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();

    sandbox.stub(console, "error");

    const getCopilotResponseStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseStub.resolves('```json\n{"spec": "fakeSpec", "funcs": ["fakeData1"]}\n```');
    const jsonParseStub = sandbox.stub(JSON, "parse");
    const jsonParseResult = {
      spec: "fakeSpec",
      funcs: ["fakeData1"],
    };
    jsonParseStub.returns(jsonParseResult);

    const result = await codeGenerator.userAskBreakdownAsync(
      fakeToken,
      100,
      spec.appendix.isCustomFunction,
      spec.appendix.host,
      spec.userInput,
      "",
      spec
    );

    jsonParseResult.funcs.filter((task: string) => {
      return !task.includes("'main'");
    });

    chai.expect(result).to.equal(jsonParseResult);
  });

  it("userAskBreakdownAsync with LLM provided json should continue, not a  customFunctions", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();

    sandbox.stub(console, "error");

    const getCopilotResponseStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseStub.resolves('```json\n{"spec": "fakeSpec", "funcs": ["fakeData1"]}\n```');
    const jsonParseStub = sandbox.stub(JSON, "parse");
    const jsonParseResult = {
      spec: "fakeSpec",
      funcs: ["fakeData1"],
    };
    jsonParseStub.returns(jsonParseResult);

    const result = await codeGenerator.userAskBreakdownAsync(
      fakeToken,
      spec.appendix.complexity,
      false,
      spec.appendix.host,
      spec.userInput,
      "",
      spec
    );

    const mainFunc =
      jsonParseResult.funcs.find((task: string) => {
        return task.includes("'main'");
      }) || "";
    chai.expect(mainFunc).not.empty;
  });

  it("generateCode - Excel - isCustomFunctions", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const host = "Excel";
    const codeSpec = "codeSpec";
    const isCustomFunctions = true;
    const suggestedFunction = ["function1", "function2"];
    const fakeSampleCode = "fakeSampleCode";
    const codeGenerator = new CodeGenerator();
    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");
    sandbox.stub(console, "error");
    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseAsStringStub.returns(Promise.resolve("```typescript\n// Some code\n```"));
    const getMostRelevantDeclarationsUsingLLMStub = sandbox.stub(
      SampleProvider.prototype,
      "getMostRelevantDeclarationsUsingLLM"
    );

    const scenarioSamples = new Map<string, SampleData>();
    scenarioSamples.set(
      "sample1",
      new SampleData(
        "Sample Name",
        "https://docs.example.com",
        "sample code",
        "description",
        "definition",
        "usage"
      )
    );
    getMostRelevantDeclarationsUsingLLMStub.returns(Promise.resolve(scenarioSamples));

    // Act
    const result = await codeGenerator.generateCode(
      fakeToken,
      host,
      spec,
      codeSpec,
      isCustomFunctions,
      suggestedFunction,
      fakeSampleCode
    );

    // Assert
    chai.expect(result).to.exist; // Replace with more specific assertions
  });

  it("generateCode - Excel - not a CustomFunctions", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const host = "Excel";
    const codeSpec = "codeSpec";
    const isCustomFunctions = false;
    const suggestedFunction = ["function1", "function2"];
    const fakeSampleCode = "fakeSampleCode";
    const codeGenerator = new CodeGenerator();
    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");
    sandbox.stub(console, "error");
    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseAsStringStub.returns(Promise.resolve("```typescript\n// Some code\n```"));
    const getMostRelevantDeclarationsUsingLLMStub = sandbox.stub(
      SampleProvider.prototype,
      "getMostRelevantDeclarationsUsingLLM"
    );

    const scenarioSamples = new Map<string, SampleData>();
    scenarioSamples.set(
      "sample1",
      new SampleData(
        "Sample Name",
        "https://docs.example.com",
        "sample code",
        "description",
        "definition",
        "usage"
      )
    );
    getMostRelevantDeclarationsUsingLLMStub.returns(Promise.resolve(scenarioSamples));

    // Act
    const result = await codeGenerator.generateCode(
      fakeToken,
      host,
      spec,
      codeSpec,
      isCustomFunctions,
      suggestedFunction,
      fakeSampleCode
    );

    // Assert
    chai.expect(result).to.exist; // Replace with more specific assertions
  });

  it("generateCode - Word", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const host = "Word";
    const codeSpec = "codeSpec";
    const isCustomFunctions = false;
    const suggestedFunction = ["function1", "function2"];
    const fakeSampleCode = "fakeSampleCode";
    const codeGenerator = new CodeGenerator();
    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");
    sandbox.stub(console, "error");
    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseAsStringStub.returns(Promise.resolve("```typescript\n// Some code\n```"));
    const getMostRelevantDeclarationsUsingLLMStub = sandbox.stub(
      SampleProvider.prototype,
      "getMostRelevantDeclarationsUsingLLM"
    );

    const scenarioSamples = new Map<string, SampleData>();
    scenarioSamples.set(
      "sample1",
      new SampleData(
        "Sample Name",
        "https://docs.example.com",
        "sample code",
        "description",
        "definition",
        "usage"
      )
    );
    getMostRelevantDeclarationsUsingLLMStub.returns(Promise.resolve(scenarioSamples));

    sandbox
      .stub(utils, "countMessagesTokens")
      .onFirstCall()
      .returns(4000)
      .onSecondCall()
      .returns(100);

    // Act
    const result = await codeGenerator.generateCode(
      fakeToken,
      host,
      spec,
      codeSpec,
      isCustomFunctions,
      suggestedFunction,
      fakeSampleCode
    );

    // Assert
    chai.expect(result).to.exist; // Replace with more specific assertions
  });

  it("generateCode - Excel - invalid return", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const host = "Excel";
    const codeSpec = "codeSpec";
    const isCustomFunctions = false;
    const suggestedFunction = ["function1", "function2"];
    const fakeSampleCode = "fakeSampleCode";
    const codeGenerator = new CodeGenerator();
    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");
    sandbox.stub(console, "error");
    const getCopilotResponseAsStringStub = sandbox.stub(utils, "getCopilotResponseAsString");
    getCopilotResponseAsStringStub.returns(Promise.resolve("some text"));
    const getMostRelevantDeclarationsUsingLLMStub = sandbox.stub(
      SampleProvider.prototype,
      "getMostRelevantDeclarationsUsingLLM"
    );

    const scenarioSamples = new Map<string, SampleData>();
    scenarioSamples.set(
      "sample1",
      new SampleData(
        "Sample Name",
        "https://docs.example.com",
        "sample code",
        "description",
        "definition",
        "usage"
      )
    );
    getMostRelevantDeclarationsUsingLLMStub.returns(Promise.resolve(scenarioSamples));

    // Act
    const result = await codeGenerator.generateCode(
      fakeToken,
      host,
      spec,
      codeSpec,
      isCustomFunctions,
      suggestedFunction,
      fakeSampleCode
    );

    // Assert
    chai.expect(result).to.equal(null); // Replace with more specific assertions
  });

  it("Invoke Failure because no breakdownResult", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();
    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");

    spec.appendix.host = "";
    spec.appendix.complexity = 0;
    sandbox.stub(codeGenerator, "userAskPreScanningAsync").resolves(null);
    const result = codeGenerator.invoke(model, fakeResponse, fakeToken, spec);
    chai.expect((await result).result).to.equal(ExecutionResultEnum.Failure);
  });

  it("Invoke Rejected", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();
    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");

    spec.appendix.host = "";
    spec.appendix.complexity = 0;
    sandbox.stub(codeGenerator, "userAskPreScanningAsync").resolves({
      host: "some host",
      shouldContinue: false,
      customFunctions: false,
      complexity: 5,
    });

    const result = codeGenerator.invoke(model, fakeResponse, fakeToken, spec);
    chai.expect((await result).result).to.equal(ExecutionResultEnum.Rejected);
  });

  it("Invoke Failure", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();
    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");

    spec.appendix.host = "";
    spec.appendix.complexity = 0;
    spec.appendix.codeSample = "";
    spec.appendix.codeTaskBreakdown = [];
    spec.appendix.codeExplanation = "";

    sandbox.stub(codeGenerator, "userAskPreScanningAsync").resolves({
      host: "some host",
      shouldContinue: true,
      customFunctions: false,
      complexity: 60,
    });

    sandbox.stub(codeGenerator, "userAskBreakdownAsync").resolves({
      spec: "some host",
      funcs: ["some data"],
    });
    sandbox.stub(codeGenerator, "generateCode").resolves(null);

    const getMostRelevantDeclarationsUsingLLMStub = sandbox.stub(
      SampleProvider.prototype,
      "getTopKMostRelevantScenarioSampleCodesBM25"
    );

    const scenarioSamples = new Map<string, SampleData>();
    scenarioSamples.set(
      "sample1",
      new SampleData(
        "Sample Name",
        "https://docs.example.com",
        "sample code",
        "description",
        "definition",
        "usage"
      )
    );
    getMostRelevantDeclarationsUsingLLMStub.returns(Promise.resolve(scenarioSamples));

    const result = codeGenerator.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect((await result).result).to.equal(ExecutionResultEnum.Failure);
  });

  it("Invoke Failure: Condition 2", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();
    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");

    spec.appendix.host = "";
    spec.appendix.complexity = 0;
    spec.appendix.codeSample = "";
    spec.appendix.codeTaskBreakdown = [];
    spec.appendix.codeExplanation = "";

    sandbox.stub(codeGenerator, "userAskPreScanningAsync").resolves({
      host: "some host",
      shouldContinue: true,
      customFunctions: false,
      complexity: 60,
    });

    sandbox.stub(codeGenerator, "userAskBreakdownAsync").resolves(null);
    sandbox.stub(codeGenerator, "generateCode").resolves(null);

    const getMostRelevantDeclarationsUsingLLMStub = sandbox.stub(
      SampleProvider.prototype,
      "getTopKMostRelevantScenarioSampleCodesBM25"
    );

    const scenarioSamples = new Map<string, SampleData>();
    scenarioSamples.set(
      "sample1",
      new SampleData(
        "Sample Name",
        "https://docs.example.com",
        "sample code",
        "description",
        "definition",
        "usage"
      )
    );
    getMostRelevantDeclarationsUsingLLMStub.returns(Promise.resolve(scenarioSamples));

    const result = codeGenerator.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect((await result).result).to.equal(ExecutionResultEnum.Failure);
  });

  it("Invoke Failure: Condition 3", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();
    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");

    spec.appendix.host = "";
    spec.appendix.complexity = 0;
    spec.appendix.codeSample = "";
    spec.appendix.codeTaskBreakdown = [];
    spec.appendix.codeExplanation = "";

    sandbox.stub(codeGenerator, "userAskPreScanningAsync").resolves({
      host: "some host",
      shouldContinue: true,
      customFunctions: false,
      complexity: 60,
    });

    sandbox.stub(codeGenerator, "userAskBreakdownAsync").resolves({
      spec: "",
      funcs: ["some data"],
    });
    sandbox.stub(codeGenerator, "generateCode").resolves(null);

    const getMostRelevantDeclarationsUsingLLMStub = sandbox.stub(
      SampleProvider.prototype,
      "getTopKMostRelevantScenarioSampleCodesBM25"
    );

    const scenarioSamples = new Map<string, SampleData>();
    scenarioSamples.set(
      "sample1",
      new SampleData(
        "Sample Name",
        "https://docs.example.com",
        "sample code",
        "description",
        "definition",
        "usage"
      )
    );
    getMostRelevantDeclarationsUsingLLMStub.returns(Promise.resolve(scenarioSamples));

    const result = codeGenerator.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect((await result).result).to.equal(ExecutionResultEnum.Failure);
  });

  it("Invoke Failure: Condition 4", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();
    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");

    spec.appendix.host = "";
    spec.appendix.complexity = 0;
    spec.appendix.codeSample = "";
    spec.appendix.codeTaskBreakdown = [];
    spec.appendix.codeExplanation = "";

    sandbox.stub(codeGenerator, "userAskPreScanningAsync").resolves({
      host: "some host",
      shouldContinue: true,
      customFunctions: false,
      complexity: 60,
    });

    sandbox.stub(codeGenerator, "userAskBreakdownAsync").resolves({
      spec: "some spec",
      funcs: [],
    });
    sandbox.stub(codeGenerator, "generateCode").resolves(null);

    const getMostRelevantDeclarationsUsingLLMStub = sandbox.stub(
      SampleProvider.prototype,
      "getTopKMostRelevantScenarioSampleCodesBM25"
    );

    const scenarioSamples = new Map<string, SampleData>();
    scenarioSamples.set(
      "sample1",
      new SampleData(
        "Sample Name",
        "https://docs.example.com",
        "sample code",
        "description",
        "definition",
        "usage"
      )
    );
    getMostRelevantDeclarationsUsingLLMStub.returns(Promise.resolve(scenarioSamples));

    const result = codeGenerator.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect((await result).result).to.equal(ExecutionResultEnum.Failure);
  });

  it("Invoke Success", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();
    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");

    spec.appendix.host = "";
    spec.appendix.complexity = 0;
    spec.appendix.codeSample = "";
    spec.appendix.codeTaskBreakdown = [];
    spec.appendix.codeExplanation = "";

    sandbox.stub(codeGenerator, "userAskPreScanningAsync").resolves({
      host: "some host",
      shouldContinue: true,
      customFunctions: false,
      complexity: 60,
    });

    sandbox.stub(codeGenerator, "userAskBreakdownAsync").resolves({
      spec: "some host 1. point 1. 2. point 2.",
      funcs: ["some data"],
    });
    sandbox.stub(codeGenerator, "generateCode").resolves("code sample");

    const getMostRelevantDeclarationsUsingLLMStub = sandbox.stub(
      SampleProvider.prototype,
      "getTopKMostRelevantScenarioSampleCodesBM25"
    );

    const scenarioSamples = new Map<string, SampleData>();
    scenarioSamples.set(
      "sample1",
      new SampleData(
        "Sample Name",
        "https://docs.example.com",
        "sample code",
        "description",
        "definition",
        "usage"
      )
    );
    getMostRelevantDeclarationsUsingLLMStub.returns(Promise.resolve(scenarioSamples));

    const result = codeGenerator.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect((await result).result).to.equal(ExecutionResultEnum.Success);
  });

  it("Invoke Success with MeasurementCodeGenExecutionTimeInTotalSec", async () => {
    const { spec, model, fakeResponse, fakeToken } = invokeParametersInit();
    const codeGenerator = new CodeGenerator();
    sandbox.stub(console, "log");
    sandbox.stub(console, "debug");
    spec.appendix.telemetryData.measurements["CodeGenExecutionTimeInTotalSec"] = 1;

    spec.appendix.host = "Excel";
    spec.appendix.complexity = 10;
    spec.appendix.shouldContinue = true;
    spec.appendix.codeSample = "sample code";
    spec.appendix.codeTaskBreakdown = ["task1", "task2"];
    spec.appendix.codeExplanation = "some explanation";
    sandbox
      .stub(codeGenerator, "userAskBreakdownAsync")
      .resolves({ spec: "fakeSpec", funcs: ["fakeData1"] });
    sandbox.stub(codeGenerator, "generateCode").resolves("Some code");

    const result = codeGenerator.invoke(model, fakeResponse, fakeToken, spec);

    chai.expect((await result).result).to.equal(ExecutionResultEnum.Success);
  });
});
