import * as chai from "chai";
import * as sinon from "sinon";
import { ChatResponseStream, comments } from "vscode";
import ts = require("typescript");
import {
  CodeIssueDetector,
  DetectionResult,
} from "../../../../src/officeChat/common/skills/codeIssueDetector";
import * as utils from "../../../../src/officeChat/common/utils";
import {
  MeasurementCompilieErrorArgumentCountMismatchCount,
  MeasurementCompilieErrorArgumentTypeMismatchCount,
  MeasurementCompilieErrorCannotAssignToReadOnlyPropertyCount,
  MeasurementCompilieErrorCannotFindModuleCount,
  MeasurementCompilieErrorCannotFindNameCount,
  MeasurementCompilieErrorConvertTypeToTypeMistakeCount,
  MeasurementCompilieErrorExpressionExpectedCount,
  MeasurementCompilieErrorOperatorAddOnTypeMismatchCount,
  MeasurementCompilieErrorOthersCount,
  MeasurementCompilieErrorOverloadMismatchCount,
  MeasurementCompilieErrorPropertyDoesNotExistOnTypeCount,
  MeasurementCompilieErrorPropertyDoesNotExistOnTypeWithSuggestionCount,
  MeasurementCompilieErrorTopLevelExpressionForbidenCount,
  MeasurementCompilieErrorTypeIsNotAssignableToTypeCount,
} from "../../../../src/officeChat/common/telemetryConsts";
import stringSimilarity = require("string-similarity");

describe("File: codeIssueDetector", () => {
  const sandbox = sinon.createSandbox();

  describe("Class: DetectionResult", () => {
    afterEach(async () => {
      sandbox.restore();
    });

    it("merge should success", () => {
      const result1 = new DetectionResult();
      const result2 = new DetectionResult();
      result1.compileErrors.push("error 1");
      result2.runtimeErrors.push("error 2");

      result1.merge(result2);
      chai.assert.deepEqual(result1.compileErrors, ["error 1"]);
      chai.assert.deepEqual(result1.runtimeErrors, ["error 2"]);
      chai.assert.deepEqual(result1.references, []);
    });

    it("areSame should works", () => {
      const result1 = new DetectionResult();
      const result2 = new DetectionResult();
      result1.compileErrors.push("error 1");
      result2.compileErrors.push("error 1");
      result1.runtimeErrors.push("error 2");
      result2.runtimeErrors.push("error 2");
      result1.references.push("ref 3");
      result2.references.push("ref 3");

      chai.assert.isTrue(result1.areSame(result2));
    });
  });

  describe("Class: CodeIssueDetector", () => {
    afterEach(async () => {
      sandbox.restore();
    });

    it("getInstance should works for singleton", () => {
      const detector1 = CodeIssueDetector.getInstance();
      chai.assert.isDefined(detector1);

      const detector2 = CodeIssueDetector.getInstance();
      chai.assert.deepEqual(detector1, detector2);
    });

    describe("Method: detectIssuesAsync", () => {
      let chatResponseStreamMock: {
        progress: sinon.SinonStub;
      };
      let telemetryData: {
        properties: { [key: string]: string };
        measurements: { [key: string]: number };
      };
      let callDetectIssueAsync: (detector: CodeIssueDetector) => Promise<DetectionResult>;

      beforeEach(() => {
        chatResponseStreamMock = {
          progress: sandbox.stub(),
        };
        telemetryData = { properties: {}, measurements: {} };
        callDetectIssueAsync = async (detector: CodeIssueDetector) => {
          return await detector.detectIssuesAsync(
            chatResponseStreamMock as unknown as ChatResponseStream,
            "Word",
            false,
            "test",
            telemetryData
          );
        };
      });

      it("normal input should works", async () => {
        const detector = CodeIssueDetector.getInstance();

        const result = await callDetectIssueAsync(detector);
        chai.assert.isDefined(result);
      }).timeout(5000);

      it("condition of `this.program` is undefined", async () => {
        const detector = CodeIssueDetector.getInstance();

        sandbox.stub(ts, "createProgram").returns(undefined as any);
        const result = await callDetectIssueAsync(detector);
        chai.assert.isDefined(result);
      }).timeout(3500);

      it("buildTypeDefAst: other conditions", async () => {
        let err = undefined;
        const detector = CodeIssueDetector.getInstance();
        const backupCompleteMemberNames = Reflect.get(detector, "completeMemberNames");
        const backupDefinionFile = Reflect.get(detector, "definionFile");
        const backupProcessNamespace = Reflect.get(detector, "processNamespace");

        Reflect.set(detector, "completeMemberNames", [{}]);
        Reflect.set(detector, "definionFile", undefined);
        Reflect.set(detector, "processNamespace", () => ["a", "b", "c"]);
        sandbox.stub(utils, "fetchRawFileContent").resolves("test");
        sandbox.stub(ts, "createSourceFile").returns([] as any);
        sandbox.stub(ts, "forEachChild").callsFake((node, fn) => {
          (node as unknown as []).forEach((n) => {
            fn(n);
          });
        });

        try {
          // Hack to direct call private methond
          detector["buildTypeDefAst"]("Word");
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "completeMemberNames", backupCompleteMemberNames);
        Reflect.set(detector, "definionFile", backupDefinionFile);
        Reflect.set(detector, "processNamespace", backupProcessNamespace);
      });
    });

    describe("Method: getCompilationErrorsAsync", () => {
      let chatResponseStreamMock: {
        progress: sinon.SinonStub;
      };
      let telemetryData: {
        properties: { [key: string]: string };
        measurements: { [key: string]: number };
      };
      let mockTSNodeForErrorTreatment: () => void;
      const backupProgram = Reflect.get(CodeIssueDetector.getInstance(), "program");

      beforeEach(() => {
        chatResponseStreamMock = {
          progress: sandbox.stub(),
        };
        telemetryData = { properties: {}, measurements: {} };
        mockTSNodeForErrorTreatment = () => {
          sandbox.stub(ts, "getPreEmitDiagnostics").returns([
            {
              file: {
                parent: {
                  arguments: [],
                  expression: "",
                },
                text: "text test",
                getStart: () => 0,
                getEnd: () => 1,
                getLineStarts: () => 1,
                getLineEndOfPosition: (x: number) => x,
                getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
              },
              start: false,
            } as any,
          ]);
          sandbox.stub(ts, "getPositionOfLineAndCharacter").returns(0);
        };
        Reflect.set(CodeIssueDetector.getInstance(), "program", "test");
      });

      afterEach(async () => {
        sandbox.restore();
        Reflect.set(CodeIssueDetector.getInstance(), "program", backupProgram);
      });

      it("condition of diagnostic.file is empty", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        Reflect.set(detector, "program", "test");
        sandbox.stub(ts, "getPreEmitDiagnostics").returns([{} as any]);

        const result = detector.getCompilationErrorsAsync("Word", false, telemetryData);

        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      }).timeout(3500);

      it("condition of node is empty", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");
        const backupFindNodeAtPosition = Reflect.get(detector, "findNodeAtPosition");

        Reflect.set(detector, "program", "test");
        Reflect.set(detector, "findNodeAtPosition", () => undefined);
        sandbox
          .stub(ts, "getPreEmitDiagnostics")
          .returns([
            { file: { getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }) } } as any,
          ]);

        const result = detector.getCompilationErrorsAsync("Word", false, telemetryData);

        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "findNodeAtPosition", backupFindNodeAtPosition);
      }).timeout(3500);

      it("other conditions in diagnostics.forEach block", async () => {
        const detector = CodeIssueDetector.getInstance();

        sandbox.stub(ts, "getPreEmitDiagnostics").returns([
          {
            file: {
              getStart: () => 0,
              getEnd: () => 0,
              getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            },
            start: false,
          } as any,
        ]);
        sandbox.stub(ts, "getPositionOfLineAndCharacter").returns(10);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Property Does Not Exist On Type With Suggestions", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("Property 'a' does not exist on type 'b'. Did you mean 'c'?");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Property Does Not Exist On Type", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("Property 'a' does not exist on type 'b'.");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Property Does Not Exist On Type - Condition 1", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        telemetryData.measurements[MeasurementCompilieErrorPropertyDoesNotExistOnTypeCount] = 0;
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("Property 'a' does not exist on type 'string | number'.");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Property Does Not Exist On Type - Condition 2", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        telemetryData.measurements[MeasurementCompilieErrorPropertyDoesNotExistOnTypeCount] = 1;
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("Property 'a' does not exist on type 'string | number'.");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Property Does Not Exist On Type - Condition 3", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("operty 'a' does not exist on type 'string | number'.");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Property Does Not Exist On Type - Condition 4", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupDefinionFile = Reflect.get(detector, "definionFile");

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("Property 'a' does not exist on type 'string'.");
        Reflect.set(detector, "definionFile", undefined);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "definionFile", backupDefinionFile);
      });

      it("error treatment: Property Does Not Exist On Type - Condition 5", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProcessNamespace = Reflect.get(detector, "processNamespace");

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("Property 'a' does not exist on type 'string'.");
        Reflect.set(detector, "processNamespace", () => ["a", "b", "c"]);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "processNamespace", backupProcessNamespace);
      });

      it("error treatment: Property Does Not Exist On Type - Condition 6", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupGetMethodsAndProperties = Reflect.get(detector, "getMethodsAndProperties");

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("Property 'a' does not exist on type 'string'.");
        sandbox.stub(ts, "isModuleDeclaration").returns(true);
        Reflect.set(detector, "getMethodsAndProperties", () => ["a", undefined, "c"]);

        const result = detector.getCompilationErrorsAsync("Office", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "getMethodsAndProperties", backupGetMethodsAndProperties);
      });

      it("error treatment: Property Does Not Exist On Type - Condition 7", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupGetMethodsAndProperties = Reflect.get(detector, "getMethodsAndProperties");

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("Property 'a' does not exist on type 'string'.");
        sandbox.stub(ts, "isModuleDeclaration").returns(true);
        Reflect.set(detector, "getMethodsAndProperties", () => undefined);

        const result = detector.getCompilationErrorsAsync("Office", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "getMethodsAndProperties", backupGetMethodsAndProperties);
      });

      it("error treatment: Property Does Not Exist On Type - Condition 8", async () => {
        const detector = CodeIssueDetector.getInstance();
        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("Property 'a' does not exist on type 'CritiqueAnnotation'.");
        sandbox.stub(ts, "isModuleDeclaration").returns(true);
        sandbox.stub(ts, "isClassDeclaration").returns(true);
        const result = detector.getCompilationErrorsAsync("Word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Property Does Not Exist On Type - Condition 9", async () => {
        const detector = CodeIssueDetector.getInstance();
        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("Property 'a' does not exist on type 'CritiqueAnnotation'.");
        sandbox.stub(ts, "isModuleDeclaration").returns(true);
        sandbox.stub(ts, "isClassDeclaration").returns(true);
        sandbox.stub(ts, "isMethodDeclaration").returns(false);
        sandbox.stub(ts, "isPropertyDeclaration").returns(false);
        const result = detector.getCompilationErrorsAsync("Word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Property Does Not Exist On Type - Condition 10", async () => {
        let err = undefined;
        const detector = CodeIssueDetector.getInstance();
        const backupProcessNamespace = Reflect.get(detector, "processNamespace");
        const backupCompleteMemberNames = Reflect.get(detector, "completeMemberNames");
        const backupgetDeclarationWithComments = Reflect.get(
          detector,
          "getDeclarationWithComments"
        );

        Reflect.set(detector, "getDeclarationWithComments", () => ({
          class: "a",
          comments: "b",
          declaration: "c",
        }));
        Reflect.set(detector, "completeMemberNames", [
          "property/method:123",
          "property/method:dasf",
        ]);
        sandbox.stub(stringSimilarity, "findBestMatch").returns({
          bestMatch: { target: "abc" },
          ratings: [
            { rating: 0.5, target: "" },
            { rating: 0.6, target: "" },
          ],
        } as any);
        sandbox.stub(ts, "forEachChild").callsFake((node, fn) => {
          Reflect.set(detector, "processNamespace", () => [
            "property/method:abc",
            "property/method:",
            "c",
          ]);
          fn(node);
          Reflect.set(detector, "processNamespace", () => undefined);
          fn(node);
        });
        try {
          // Hack to direct call private methond
          detector["getErrorTreatment"](
            "Word",
            {} as any,
            "Property 'a' does not exist on type 'CritiqueAnnotation'.",
            {
              properties: {},
              measurements: {},
            }
          );
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "processNamespace", backupProcessNamespace);
        Reflect.set(detector, "completeMemberNames", backupCompleteMemberNames);
        Reflect.set(detector, "getDeclarationWithComments", backupgetDeclarationWithComments);
      });

      it("error treatment: Property Does Not Exist On Type - Condition 11", async () => {
        let err = undefined;
        const detector = CodeIssueDetector.getInstance();
        const backupProcessNamespace = Reflect.get(detector, "processNamespace");
        const backupCompleteMemberNames = Reflect.get(detector, "completeMemberNames");
        const backupgetDeclarationWithComments = Reflect.get(
          detector,
          "getDeclarationWithComments"
        );

        Reflect.set(detector, "getDeclarationWithComments", () => ({
          class: "a",
          comments: "b",
          declaration: "c",
        }));
        Reflect.set(detector, "completeMemberNames", [
          "property/method:abc",
          "property/method:dasf",
        ]);
        sandbox.stub(stringSimilarity, "findBestMatch").returns({
          bestMatch: { target: "abc" },
          ratings: [
            { rating: 0.5, target: "" },
            { rating: 0.6, target: "" },
          ],
        } as any);
        sandbox.stub(ts, "forEachChild").callsFake((node, fn) => {
          Reflect.set(detector, "processNamespace", () => [
            "property/method:abc",
            "property/method:",
            "c",
          ]);
          fn(node);
          Reflect.set(detector, "processNamespace", () => undefined);
          fn(node);
        });
        try {
          // Hack to direct call private methond
          detector["getErrorTreatment"](
            "Word",
            {} as any,
            "Property 'a' does not exist on type 'CritiqueAnnotation'.",
            {
              properties: {},
              measurements: {},
            }
          );
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "processNamespace", backupProcessNamespace);
        Reflect.set(detector, "completeMemberNames", backupCompleteMemberNames);
        Reflect.set(detector, "getDeclarationWithComments", backupgetDeclarationWithComments);
      });

      it("error treatment: No Function Return Or No Implementation", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns(
            "A function whose declared type is neither 'undefined', 'void', nor 'any' must return a value."
          );

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: No Function Return Or No Implementation - Condition 1", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        telemetryData.measurements[
          MeasurementCompilieErrorPropertyDoesNotExistOnTypeWithSuggestionCount
        ] = 1;
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns(
            "A function whose declared type is neither 'undefined', 'void', nor 'any' must return a value."
          );

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Cannot Find Module", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("Cannot find module");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Cannot Find Module - Condition 1", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("Cannot find module");
        telemetryData.measurements[MeasurementCompilieErrorCannotFindModuleCount] = 1;

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Argument Count Mismatch", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("arguments, but got 1");
        Reflect.set(detector, "program", {
          getTypeChecker: () => ({
            getSymbolAtLocation: () => ({ getDeclarations: () => [1, 2] }),
            getSignatureFromDeclaration: () => ({
              parameters: [1, 2],
              getDeclaration: () => ({ getText: () => "text" }),
            }),
          }),
        });
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(true);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Argument Count Mismatch - Condition 1", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("arguments, but got 1");
        telemetryData.measurements[MeasurementCompilieErrorArgumentCountMismatchCount] = 1;
        Object.defineProperty(
          telemetryData.measurements,
          MeasurementCompilieErrorArgumentCountMismatchCount,
          {
            get() {
              Reflect.set(detector, "program", undefined);
              return 1;
            },
            set() {},
          }
        );
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(false);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Argument Count Mismatch - Condition 2", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("arguments, but got 1");
        Reflect.set(detector, "program", {
          getTypeChecker: () => undefined,
        });
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(true);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Argument Count Mismatch - Condition 3", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("arguments, but got 1");
        Reflect.set(detector, "program", {
          getTypeChecker: () => ({
            getSymbolAtLocation: () => ({
              getDeclarations: () => [{ getText: () => "text" }, { getText: () => "text" }],
            }),
            getSignatureFromDeclaration: () => undefined,
          }),
        });
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(true);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Argument Count Mismatch - Condition 4", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("arguments, but got 1");
        Reflect.set(detector, "program", {
          getTypeChecker: () => ({
            getSymbolAtLocation: () => ({
              getDeclarations: () => [],
            }),
            getSignatureFromDeclaration: () => undefined,
          }),
        });
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(true);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Argument Count Mismatch 4", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        sandbox.stub(ts, "getPreEmitDiagnostics").returns([
          {
            file: {
              parent: null,
              text: "text test",
              getStart: () => 0,
              getEnd: () => 1,
              getLineStarts: () => 1,
              getLineEndOfPosition: (x: number) => x,
              getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            },
            start: false,
          } as any,
        ]);
        sandbox.stub(ts, "getPositionOfLineAndCharacter").returns(0);
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("arguments, but got 1");
        Reflect.set(detector, "program", {
          getTypeChecker: () => ({
            getSymbolAtLocation: () => ({ getDeclarations: () => [1, 2] }),
            getSignatureFromDeclaration: () => ({
              parameters: [1, 2],
              getDeclaration: () => ({ getText: () => "text" }),
            }),
          }),
        });
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Argument Type Mismatch", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("Argument of type");
        Reflect.set(detector, "program", {
          getTypeChecker: () => ({
            getSymbolAtLocation: () => ({
              getDeclarations: () => [{ getFullText: () => "text" }],
            }),
            getSignatureFromDeclaration: () => undefined,
          }),
        });
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(true);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Argument Type Mismatch - Condition 1", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("Argument of type 'aa' is not assignable to parameter of type 'bb'.");
        Reflect.set(detector, "program", {
          getTypeChecker: () => ({
            getSymbolAtLocation: () => ({
              getDeclarations: () => [{ getFullText: () => "text" }],
            }),
            getSignatureFromDeclaration: () => undefined,
          }),
        });
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(false);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Argument Type Mismatch - Condition 2", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        telemetryData.measurements[MeasurementCompilieErrorArgumentTypeMismatchCount] = 1;
        Object.defineProperty(
          telemetryData.measurements,
          MeasurementCompilieErrorArgumentTypeMismatchCount,
          {
            get() {
              Reflect.set(detector, "program", undefined);
              return 1;
            },
            set() {},
          }
        );
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("Argument of type");
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(true);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Argument Type Mismatch - Condition 3", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("Argument of type");
        Reflect.set(detector, "program", {
          getTypeChecker: () => ({
            getSymbolAtLocation: () => ({
              getDeclarations: () => [],
            }),
            getSignatureFromDeclaration: () => undefined,
          }),
        });
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(true);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Argument Type Mismatch - Condition 4", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("Argument of type");
        Reflect.set(detector, "program", {
          getTypeChecker: () => ({
            getSymbolAtLocation: () => ({
              getDeclarations: () => [{ getFullText: () => "text" }],
            }),
            getSignatureFromDeclaration: () => undefined,
          }),
        });
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(false);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Operator Add On Type Mismatch", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("Operator '+' cannot be applied to types");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Operator Add On Type Mismatch - Condition 1", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        telemetryData.measurements[MeasurementCompilieErrorOperatorAddOnTypeMismatchCount] = 1;
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("Operator '+' cannot be applied to types");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Type Is Not Assignable To Type", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("is not assignable to type");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Type Is Not Assignable To Type - Condition 1", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        telemetryData.measurements[MeasurementCompilieErrorTypeIsNotAssignableToTypeCount] = 1;
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("is not assignable to type");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Convert Type To Type Mistake", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("may be a mistake because neither type sufficiently overlaps with the other");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Convert Type To Type Mistake - Condition 1", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        telemetryData.measurements[MeasurementCompilieErrorConvertTypeToTypeMistakeCount] = 1;
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("may be a mistake because neither type sufficiently overlaps with the other");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Overload Mismatch", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("No overload matches this call. Overload 1 of 22");
        Reflect.set(detector, "program", {
          getTypeChecker: () => ({
            getSymbolAtLocation: () => ({
              getDeclarations: () => [{ getFullText: () => "text" }],
            }),
            getSignatureFromDeclaration: () => undefined,
          }),
        });
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(true);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Overload Mismatch - Condition 1", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("No overload matches this call. Overload 1 of 22");
        telemetryData.measurements[MeasurementCompilieErrorOverloadMismatchCount] = 1;
        Object.defineProperty(
          telemetryData.measurements,
          MeasurementCompilieErrorOverloadMismatchCount,
          {
            get() {
              Reflect.set(detector, "program", undefined);
              return 1;
            },
            set() {},
          }
        );
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(true);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Overload Mismatch - Condition 2", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("No overload matches this call. Overload 1 of 22");
        Reflect.set(detector, "program", {
          getTypeChecker: () => undefined,
        });
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(true);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Overload Mismatch - Condition 3", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns("No overload matches this call. Overload 1 of 22");
        Reflect.set(detector, "program", {
          getTypeChecker: () => ({
            getSymbolAtLocation: () => ({
              getDeclarations: () => [],
            }),
            getSignatureFromDeclaration: () => undefined,
          }),
        });
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(true);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Overload Mismatch - Condition 3", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns(
            "No overload matches this call. Overload 1 of 3, 'test', gave the following error."
          );
        Reflect.set(detector, "program", {
          getTypeChecker: () => ({}),
        });
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(false);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Overload Mismatch - Condition 4", async () => {
        const detector = CodeIssueDetector.getInstance();
        const backupProgram = Reflect.get(detector, "program");

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns(
            "No overload matches this call. Overload 1 of 33 of, 'dsd', gave the following error."
          );
        Reflect.set(detector, "program", {
          getTypeChecker: () => ({}),
        });
        const isCallExpressionStub = sandbox.stub(ts, "isCallExpression");
        isCallExpressionStub.onCall(0).returns(false);
        isCallExpressionStub.onCall(1).returns(true);
        isCallExpressionStub.onCall(2).returns(false);

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
        Reflect.set(detector, "program", backupProgram);
      });

      it("error treatment: Cannot Find Name", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("Cannot find name");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Cannot Find Name - Condition 1", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        telemetryData.measurements[MeasurementCompilieErrorCannotFindNameCount] = 1;
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("Cannot find name");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Cannot Assign To Read Only Property", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("Cannot assign to");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Cannot Assign To Read Only Property - Condition 1", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        telemetryData.measurements[MeasurementCompilieErrorCannotAssignToReadOnlyPropertyCount] = 1;
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("Cannot assign to");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Top Level Expression Forbiden", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns(
            "expressions are only allowed at the top level of a file when that file is a module"
          );

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Top Level Expression Forbiden - Condition 1", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        telemetryData.measurements[MeasurementCompilieErrorTopLevelExpressionForbidenCount] = 1;
        sandbox
          .stub(ts, "flattenDiagnosticMessageText")
          .returns(
            "expressions are only allowed at the top level of a file when that file is a module"
          );

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Expression Expected Handlder", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("Expression expected");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Expression Expected Handlder - Condition 1", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        telemetryData.measurements[MeasurementCompilieErrorExpressionExpectedCount] = 1;
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("Expression expected");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("error treatment: Others", async () => {
        const detector = CodeIssueDetector.getInstance();

        mockTSNodeForErrorTreatment();
        telemetryData.measurements[MeasurementCompilieErrorOthersCount] = 1;
        sandbox.stub(ts, "flattenDiagnosticMessageText").returns("Others Others");

        const result = detector.getCompilationErrorsAsync("word", false, telemetryData);
        chai.assert.isDefined(result);
      });

      it("getMethodsAndProperties - condition 1", () => {
        let err = undefined;
        const detector = CodeIssueDetector.getInstance();

        sandbox.stub(ts, "isClassDeclaration").returns(true);

        try {
          // Hack to direct call private methond
          detector["getMethodsAndProperties"]("Yes", {
            name: {
              getText: () => false,
            },
          } as any);
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
      });
    });

    it("getMethodsAndProperties - condition 1", () => {
      let err = undefined;
      const detector = CodeIssueDetector.getInstance();

      sandbox.stub(ts, "isClassDeclaration").returns(true);
      sandbox.stub(ts, "isMethodDeclaration").throws(new Error("error"));
      sandbox.stub(console, "error").callsFake(() => {});
      try {
        // Hack to direct call private methond
        detector["getMethodsAndProperties"](null, {
          name: {
            getText: () => false,
          },
          members: [{}],
        } as any);
      } catch (e) {
        err = e;
      }

      chai.assert.isUndefined(err);
    });

    it("getMethodsAndProperties - condition 2", () => {
      let err = undefined;
      const detector = CodeIssueDetector.getInstance();

      sandbox.stub(ts, "isClassDeclaration").returns(true);
      sandbox.stub(ts, "isMethodDeclaration").throws(new Error("error"));
      sandbox.stub(console, "error").callsFake(() => {});
      try {
        // Hack to direct call private methond
        detector["getMethodsAndProperties"](null, {
          name: undefined,
          members: [{}],
        } as any);
      } catch (e) {
        err = e;
      }

      chai.assert.isUndefined(err);
    });

    it("getDeclarationWithComments - condition 1", () => {
      let err = undefined;
      const detector = CodeIssueDetector.getInstance();
      const backupDefinionFile = Reflect.get(detector, "definionFile");

      Reflect.set(detector, "definionFile", {
        text: "text",
        getFullText: () => "text",
        children: [
          {
            name: { getText: () => "a" },
            getFullText: () => "text",
            children: [
              {
                name: { getText: () => "b" },
                getFullText: () => "text",
                children: [
                  {
                    name: { getText: () => "d" },
                    getFullText: () => "text",
                    children: [
                      {
                        name: { getText: () => "c" },
                        getFullText: () => "text",
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      });
      sandbox.stub(ts, "isModuleDeclaration").returns(true);
      sandbox.stub(ts, "isClassDeclaration").returns(true);
      sandbox.stub(ts, "isPropertyDeclaration").returns(false);
      sandbox.stub(ts, "isMethodDeclaration").returns(true);
      sandbox.stub(ts, "getLeadingCommentRanges").returns([{ pos: 0, end: 1 }] as any);
      sandbox.stub(ts, "forEachChild").callsFake((node, fn) => {
        (node as unknown as any).children?.forEach((n: any) => {
          fn(n);
        });
      });

      try {
        // Hack to direct call private methond
        detector["getDeclarationWithComments"]("a", "b", "c");
      } catch (e) {
        err = e;
      }

      chai.assert.isUndefined(err);
      Reflect.set(detector, "definionFile", backupDefinionFile);
    });

    it("getDeclarationWithComments - condition 2", () => {
      let err = undefined;
      const detector = CodeIssueDetector.getInstance();
      const backupDefinionFile = Reflect.get(detector, "definionFile");

      Reflect.set(detector, "definionFile", {
        text: "text",
        getFullText: () => "text",
        children: [
          {
            name: undefined,
            getFullText: () => "text",
            children: [
              {
                name: { getText: () => "c" },
                getFullText: () => "text",
              },
            ],
          },
        ],
      });
      sandbox.stub(ts, "isModuleDeclaration").returns(false);
      sandbox.stub(ts, "isClassDeclaration").returns(true);
      sandbox.stub(ts, "isPropertyDeclaration").returns(false);
      sandbox
        .stub(ts, "isMethodDeclaration")
        .onFirstCall()
        .returns(false)
        .onSecondCall()
        .returns(true)
        .onThirdCall()
        .returns(true);
      sandbox.stub(ts, "getLeadingCommentRanges").returns(false as any);
      sandbox.stub(ts, "forEachChild").callsFake((node, fn) => {
        (node as unknown as any).children?.forEach((n: any) => {
          fn(n);
        });
      });

      try {
        // Hack to direct call private methond
        detector["getDeclarationWithComments"]("a", "b", "c");
      } catch (e) {
        err = e;
      }

      chai.assert.isUndefined(err);
      Reflect.set(detector, "definionFile", backupDefinionFile);
    });

    it("getDeclarationWithComments - condition 3", () => {
      let err = undefined;
      const detector = CodeIssueDetector.getInstance();
      const backupDefinionFile = Reflect.get(detector, "definionFile");

      Reflect.set(detector, "definionFile", {});
      sandbox.stub(ts, "forEachChild").callsFake(() => {});

      try {
        // Hack to direct call private methond
        detector["getDeclarationWithComments"]("a", "b", "c");
      } catch (e) {
        err = e;
      }

      chai.assert.isUndefined(err);
      Reflect.set(detector, "definionFile", backupDefinionFile);
    });

    it("processNamespace - Condition 1", () => {
      let err = undefined;
      const detector = CodeIssueDetector.getInstance();
      const backupGetMethodsAndProperties = Reflect.get(detector, "getMethodsAndProperties");

      Reflect.set(detector, "getMethodsAndProperties", () => ["a", undefined, "c"]);
      sandbox.stub(ts, "isModuleDeclaration").returns(true);
      sandbox.stub(ts, "isModuleBlock").returns(true);
      sandbox.stub(ts, "forEachChild").callsFake((node, fn) => {
        (node as unknown as any).children?.forEach((n: any) => {
          fn(n);
        });
      });

      try {
        // Hack to direct call private methond
        detector["processNamespace"]("a", "b", {
          name: { getText: () => "a" },
          children: [{ children: [{}] }],
        } as any);
      } catch (e) {
        err = e;
      }

      chai.assert.isUndefined(err);
      Reflect.set(detector, "getMethodsAndProperties", backupGetMethodsAndProperties);
    });

    it("processNamespace - Condition 2", () => {
      let err = undefined;
      const detector = CodeIssueDetector.getInstance();
      const backupGetMethodsAndProperties = Reflect.get(detector, "getMethodsAndProperties");

      Reflect.set(detector, "getMethodsAndProperties", () => undefined);
      sandbox.stub(ts, "isModuleDeclaration").returns(true);
      sandbox.stub(ts, "isModuleBlock").returns(true);
      sandbox.stub(ts, "forEachChild").callsFake((node, fn) => {
        (node as unknown as any).children?.forEach((n: any) => {
          fn(n);
        });
      });

      try {
        // Hack to direct call private methond
        detector["processNamespace"]("a", "b", {
          name: { getText: () => "a" },
          children: [{ children: [{}] }],
        } as any);
      } catch (e) {
        err = e;
      }

      chai.assert.isUndefined(err);
      Reflect.set(detector, "getMethodsAndProperties", backupGetMethodsAndProperties);
    });

    describe("Method: getPotentialRuntimeIssues", () => {
      let telemetryData: {
        properties: { [key: string]: string };
        measurements: { [key: string]: number };
      };

      beforeEach(() => {
        telemetryData = { properties: {}, measurements: {} };
      });

      afterEach(async () => {
        sandbox.restore();
      });

      it("condition when is Custom Function", () => {
        const detector = CodeIssueDetector.getInstance();

        const result = detector.getPotentialRuntimeIssues("Word", true, telemetryData);
        chai.assert.isDefined(result);
      });

      it("typeChecker undefined would return in the beginning", () => {
        const detector = CodeIssueDetector.getInstance();
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "typeChecker", undefined);
        const result = detector.getPotentialRuntimeIssues("Word", false, telemetryData);

        chai.assert.isDefined(result);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      it("runtime issue: findImportAndRequireStatements", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "typeChecker", {});
        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            isImportDeclaration: true,
            isVariableStatement: false,
            isExpressionStatement: false,
            getText: () => "import",
            getStart: () => 0,
            getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            children: [
              {
                isImportDeclaration: false,
                isVariableStatement: true,
                isExpressionStatement: false,
                getText: () => "require()",
                getStart: () => 0,
              },
              {
                isImportDeclaration: false,
                isVariableStatement: false,
                isExpressionStatement: true,
                getText: () => "require()",
                getStart: () => 0,
              },
            ],
          }),
        });
        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
        });
        sandbox
          .stub(ts, "isImportDeclaration")
          .callsFake((node) => (node as any).isImportDeclaration);
        sandbox
          .stub(ts, "isVariableStatement")
          .callsFake((node) => (node as any).isVariableStatement);
        sandbox
          .stub(ts, "isExpressionStatement")
          .callsFake((node) => (node as any).isExpressionStatement);

        try {
          // Hack to direct call private methond
          detector["findImportAndRequireStatements"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      it("runtime issue: findImportAndRequireStatements - Condition 1", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "program", { getSourceFile: () => false });
        Reflect.set(detector, "typeChecker", {});

        try {
          // Hack to direct call private methond
          detector["findImportAndRequireStatements"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      it("runtime issue: findEntryFunctionInGeneratedCode", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "typeChecker", {});
        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            isFunctionDeclaration: true,
            name: { text: "main", getText: () => "main" },
            parameters: [],
            modifiers: [],
          }),
        });
        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
          if (!t.name) throw new Error("name is undefined");
        });
        sandbox
          .stub(ts, "isFunctionDeclaration")
          .callsFake((node) => (node as any).isFunctionDeclaration);

        try {
          // Hack to direct call private methond
          detector["findEntryFunctionInGeneratedCode"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      it("runtime issue: findEntryFunctionInGeneratedCode - Condition 1", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "typeChecker", {});
        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            isFunctionDeclaration: true,
            name: { text: "main2", getText: () => "main2" },
            parameters: [1, 2],
            modifiers: [{ kind: ts.SyntaxKind.AsyncKeyword }],
          }),
        });
        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
          if (!t.name) throw new Error("name is undefined");
        });
        sandbox
          .stub(ts, "isFunctionDeclaration")
          .callsFake((node) => (node as any).isFunctionDeclaration);

        try {
          // Hack to direct call private methond
          detector["findEntryFunctionInGeneratedCode"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      it("runtime issue: findEntryFunctionInGeneratedCode - Condition 2", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "typeChecker", {});
        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            isFunctionDeclaration: true,
            name: { text: "main", getText: () => "main" },
            parameters: [1, 2],
            modifiers: undefined,
          }),
        });
        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
          if (!t.name) throw new Error("name is undefined");
        });
        sandbox
          .stub(ts, "isFunctionDeclaration")
          .callsFake((node) => (node as any).isFunctionDeclaration);

        try {
          // Hack to direct call private methond
          detector["findEntryFunctionInGeneratedCode"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      it("runtime issue: findEntryFunctionInGeneratedCode - Condition 3", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "typeChecker", {});
        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            isFunctionDeclaration: true,
            name: undefined,
          }),
        });
        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
          if (!t.name) throw new Error("name is undefined");
        });
        sandbox
          .stub(ts, "isFunctionDeclaration")
          .callsFake((node) => (node as any).isFunctionDeclaration);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          detector["findEntryFunctionInGeneratedCode"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      it("runtime issue: findEntryFunctionInGeneratedCode - Condition 4", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "typeChecker", {});
        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            isFunctionDeclaration: true,
            name: { text: "main" },
            parameters: [],
            modifiers: [{ kind: ts.SyntaxKind.AsyncKeyword }],
          }),
        });
        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
          if (!t.name) throw new Error("name is undefined");
        });
        sandbox
          .stub(ts, "isFunctionDeclaration")
          .callsFake((node) => (node as any).isFunctionDeclaration);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          detector["findEntryFunctionInGeneratedCode"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      it("runtime issue: findEntryFunctionInGeneratedCode - Condition 5", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "typeChecker", {});
        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            isFunctionDeclaration: false,
            name: "Yes",
            parent: {
              getText: () => "function main",
            },
            parameters: [1, 2, 3],
            modifiers: [{ kind: ts.SyntaxKind.AsyncKeyword }],
          }),
        });
        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
          if (!t.name) throw new Error("name is undefined");
        });
        sandbox
          .stub(ts, "isFunctionDeclaration")
          .onFirstCall()
          .returns(true)
          .onSecondCall()
          .returns(false);
        sandbox.stub(ts, "isArrowFunction").returns(true);
        sandbox.stub(ts, "isFunctionExpression").returns(true);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          detector["findEntryFunctionInGeneratedCode"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      it("runtime issue: findEntryFunctionInGeneratedCode - Condition 6", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "typeChecker", {});
        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            isFunctionDeclaration: false,
            name: "Yes",
            parent: undefined,
            parameters: [1, 2, 3],
            modifiers: [{ kind: ts.SyntaxKind.AsyncKeyword }],
          }),
        });
        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
          if (!t.name) throw new Error("name is undefined");
        });
        sandbox
          .stub(ts, "isFunctionDeclaration")
          .onFirstCall()
          .returns(true)
          .onSecondCall()
          .returns(false);
        sandbox.stub(ts, "isArrowFunction").returns(true);
        sandbox.stub(ts, "isFunctionExpression").returns(true);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          detector["findEntryFunctionInGeneratedCode"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      it("runtime issue: findEntryFunctionInGeneratedCode - Condition 7", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "typeChecker", {});
        Reflect.set(detector, "program", {
          getSourceFile: () => false,
        });

        try {
          // Hack to direct call private methond
          detector["findEntryFunctionInGeneratedCode"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      // eslint-disable-next-line no-secrets/no-secrets
      it("runtime issue: findPropertyAccessAfterCallExpression", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "typeChecker", {});
        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            getStart: () => 0,
            parent: true,
            expression: { getFullText: () => "main1" },
            name: { getText: () => "main1" },
            children: [{ name: undefined }],
          }),
        });
        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
          if (!t.name) throw new Error("expression is undefined");
        });
        sandbox.stub(ts, "isPropertyAccessExpression").returns(true);
        sandbox
          .stub(ts, "isCallExpression")
          .onFirstCall()
          .returns(false)
          .onSecondCall()
          .returns(true);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          // eslint-disable-next-line no-secrets/no-secrets
          detector["findPropertyAccessAfterCallExpression"]("Word");
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      // eslint-disable-next-line no-secrets/no-secrets
      it("runtime issue: findPropertyAccessAfterCallExpression - Condition 1", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "typeChecker", {});
        Reflect.set(detector, "program", {
          getSourceFile: () => false,
        });

        try {
          // Hack to direct call private methond
          // eslint-disable-next-line no-secrets/no-secrets
          detector["findPropertyAccessAfterCallExpression"]("Word");
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      it("runtime issue: findOfficeAPIObjectPropertyAccess", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            getStart: () => 0,
            expression: { getFullText: () => "main1" },
            name: { text: "main1" },
            children: [{ name: undefined }],
          }),
        });
        Reflect.set(detector, "typeChecker", {
          getTypeAtLocation: () => ({
            symbol: {
              escapedName: "Word",
            },
          }),
        });
        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
          if (!t.name) throw new Error("expression is undefined");
        });
        sandbox.stub(ts, "isPropertyAccessExpression").returns(true);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          detector["findOfficeAPIObjectPropertyAccess"]("Word");
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      it("runtime issue: findOfficeAPIObjectPropertyAccess - Condition 1", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            expression: {},
            name: { text: "main1" },
          }),
        });
        Reflect.set(detector, "typeChecker", {
          getTypeAtLocation: () => undefined,
        });
        sandbox.stub(ts, "forEachChild").callsFake(() => {});
        sandbox.stub(ts, "isPropertyAccessExpression").returns(true);

        try {
          // Hack to direct call private methond
          detector["findOfficeAPIObjectPropertyAccess"]("Word");
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      it("runtime issue: findOfficeAPIObjectPropertyAccess - Condition 2", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");

        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            getStart: () => 0,
            expression: {},
            name: { text: "main1" },
          }),
        });
        Reflect.set(detector, "typeChecker", {
          getTypeAtLocation: () => ({
            symbol: { escapedName: "" },
          }),
        });

        sandbox.stub(ts, "forEachChild").callsFake(() => {});
        sandbox.stub(ts, "isPropertyAccessExpression").returns(true);

        try {
          // Hack to direct call private methond
          detector["findOfficeAPIObjectPropertyAccess"]("");
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
      });

      it("runtime issue: findExcelA1NotationInStringConcatenation", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");
        const backupFunc = Reflect.get(detector, "isValidExcelA1Notation");

        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            getStart: () => 0,
            getText: () => "main1",
            left: { text: "main1", getFullText: () => "main1" },
            right: { text: "main1", getFullText: () => "main1" },
            name: { text: "main1" },
            children: [
              {
                getStart: () => 0,
                getText: () => "main1",
                left: { text: "main1", getFullText: () => "main1" },
                right: { text: "main1", getFullText: () => "main1" },
                name: { text: "main1" },
              },
              {
                getStart: () => 0,
                getText: () => "main1",
                left: { text: "main1", getFullText: () => "main1" },
                right: { text: "main1", getFullText: () => "main1" },
                name: { text: "main1" },
              },
              { name: undefined },
            ],
          }),
        });
        Reflect.set(detector, "typeChecker", {
          getTypeAtLocation: () => ({}),
          typeToString: () => "number",
        });
        Reflect.set(detector, "isValidExcelA1Notation", () => true);

        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
          if (!t.name) throw new Error("expression is undefined");
        });
        sandbox.stub(ts, "isBinaryExpression").returns(true);
        sandbox
          .stub(ts, "isStringLiteral")
          .onCall(0)
          .returns(true)
          .onCall(1)
          .returns(false)
          .onCall(2)
          .returns(true);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          detector["findExcelA1NotationInStringConcatenation"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
        Reflect.set(detector, "isValidExcelA1Notation", backupFunc);
      });

      it("runtime issue: findExcelA1NotationInStringConcatenation - Condition 1", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");
        const backupFunc = Reflect.get(detector, "isValidExcelA1Notation");

        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            getStart: () => 0,
            left: { text: "main1" },
            right: { text: "main1" },
            name: { text: "main1" },
            children: [
              {
                getStart: () => 0,
                left: { text: "main1" },
                right: { text: "main1" },
                name: { text: "main1" },
              },
              {
                getStart: () => 0,
                left: { text: "main1" },
                right: { text: "main1" },
                name: { text: "main1" },
              },
            ],
          }),
        });
        Reflect.set(detector, "typeChecker", {
          getTypeAtLocation: () => ({}),
          typeToString: () => "string",
        });
        Reflect.set(detector, "isValidExcelA1Notation", () => true);

        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
        });
        sandbox.stub(ts, "isBinaryExpression").returns(true);
        sandbox
          .stub(ts, "isStringLiteral")
          .onCall(0)
          .returns(true)
          .onCall(1)
          .returns(false)
          .onCall(2)
          .returns(true);

        try {
          // Hack to direct call private methond
          detector["findExcelA1NotationInStringConcatenation"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
        Reflect.set(detector, "isValidExcelA1Notation", backupFunc);
      });

      it("runtime issue: findExcelA1NotationInStringInterpolation", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");
        const backupFunc = Reflect.get(detector, "isValidExcelA1Notation");

        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            getStart: () => 0,
            getText: () => "main1",
            head: { text: "main1" },
            templateSpans: [
              {
                expression: {
                  getFullText: () => "main1",
                  operatorToken: { kind: ts.SyntaxKind.PlusToken },
                  left: { getFullText: () => "main1" },
                  right: { getFullText: () => "main1" },
                },
              },
            ],
            name: { text: "main1" },
            children: [{ name: undefined }],
          }),
        });
        Reflect.set(detector, "typeChecker", {
          getTypeAtLocation: () => ({ isNumberLiteral: () => false }),
          typeToString: () => "number",
        });
        Reflect.set(detector, "isValidExcelA1Notation", () => true);

        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
          if (!t.name) throw new Error("expression is undefined");
        });
        sandbox.stub(ts, "isTemplateExpression").returns(true);
        sandbox.stub(ts, "isBinaryExpression").returns(true);
        sandbox.stub(ts, "isPropertyAccessExpression").returns(true);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          detector["findExcelA1NotationInStringInterpolation"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
        Reflect.set(detector, "isValidExcelA1Notation", backupFunc);
      });

      it("runtime issue: findExcelA1NotationInStringInterpolation - Condition 1", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");
        const backupFunc = Reflect.get(detector, "isValidExcelA1Notation");

        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            getStart: () => 0,
            getText: () => "main1",
            head: { text: "main1" },
            templateSpans: [
              {
                expression: {
                  getFullText: () => "main1",
                  operatorToken: { kind: ts.SyntaxKind.PlusToken },
                  left: { getFullText: () => "main1", values: "main1" },
                  right: { getFullText: () => "main1", value: "main1" },
                },
              },
            ],
            name: { text: "main1" },
          }),
        });
        Reflect.set(detector, "typeChecker", {
          getTypeAtLocation: () => ({ isNumberLiteral: () => true }),
          typeToString: () => "number",
        });
        Reflect.set(detector, "isValidExcelA1Notation", () => true);

        sandbox.stub(ts, "forEachChild").callsFake(() => {});
        sandbox.stub(ts, "isTemplateExpression").returns(true);
        sandbox.stub(ts, "isBinaryExpression").returns(true);
        sandbox.stub(ts, "isPropertyAccessExpression").returns(false);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          detector["findExcelA1NotationInStringInterpolation"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
        Reflect.set(detector, "isValidExcelA1Notation", backupFunc);
      });

      it("runtime issue: findExcelA1NotationInStringInterpolation - Condition 2", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");
        const backupFunc = Reflect.get(detector, "isValidExcelA1Notation");

        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            getStart: () => 0,
            getText: () => "main1",
            head: { text: "main1" },
            templateSpans: [
              {
                expression: {
                  getFullText: () => "main1",
                  operatorToken: { kind: ts.SyntaxKind.PlusToken },
                  left: { getFullText: () => "main1", value: "main1", type: "string" },
                  right: { getFullText: () => "main1", value: "main1", type: "number" },
                },
              },
            ],
            name: { text: "main1" },
          }),
        });
        Reflect.set(detector, "typeChecker", {
          getTypeAtLocation: (x: object) => ({ ...x, isNumberLiteral: () => true }),
          typeToString: (x: any) => x.type,
        });
        Reflect.set(detector, "isValidExcelA1Notation", () => true);

        sandbox.stub(ts, "forEachChild").callsFake(() => {});
        sandbox.stub(ts, "isTemplateExpression").returns(true);
        sandbox.stub(ts, "isBinaryExpression").returns(true);
        sandbox.stub(ts, "isPropertyAccessExpression").returns(false);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          detector["findExcelA1NotationInStringInterpolation"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
        Reflect.set(detector, "isValidExcelA1Notation", backupFunc);
      });

      it("runtime issue: findExcelA1NotationInStringInterpolation - Condition 3", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");
        const backupFunc = Reflect.get(detector, "isValidExcelA1Notation");

        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            getStart: () => 0,
            getText: () => "main1",
            head: { text: "main1" },
            templateSpans: [
              {
                expression: {
                  getFullText: () => "main1",
                  operatorToken: { kind: ts.SyntaxKind.MinusToken },
                  left: { getFullText: () => "main1" },
                  right: { getFullText: () => "main1" },
                },
              },
            ],
            name: { text: "main1" },
            children: [{ name: undefined }],
          }),
        });
        Reflect.set(detector, "typeChecker", {
          getTypeAtLocation: () => ({ isNumberLiteral: () => false }),
          typeToString: () => "string",
        });
        Reflect.set(detector, "isValidExcelA1Notation", () => true);

        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
          if (!t.name) throw new Error("expression is undefined");
        });
        sandbox.stub(ts, "isTemplateExpression").returns(true);
        sandbox.stub(ts, "isBinaryExpression").returns(true);
        sandbox.stub(ts, "isPropertyAccessExpression").returns(false);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          detector["findExcelA1NotationInStringInterpolation"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
        Reflect.set(detector, "isValidExcelA1Notation", backupFunc);
      });

      it("runtime issue: findExcelA1NotationInStringInterpolation - Condition 4", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");
        const backupFunc = Reflect.get(detector, "isValidExcelA1Notation");

        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            getStart: () => 0,
            getText: () => "main1",
            head: { text: "main1" },
            templateSpans: [
              {
                expression: {
                  getFullText: () => "main1",
                  operatorToken: { kind: ts.SyntaxKind.PlusToken },
                  left: { getFullText: () => "main1", value: "main1", type: "number" },
                  right: { getFullText: () => "main1", value: "main1", type: "number" },
                },
              },
            ],
            name: { text: "main1" },
          }),
        });
        Reflect.set(detector, "typeChecker", {
          getTypeAtLocation: (x: object) => ({ ...x, isNumberLiteral: () => true }),
          typeToString: (x: any) => x.type,
        });
        Reflect.set(detector, "isValidExcelA1Notation", () => true);

        sandbox.stub(ts, "forEachChild").callsFake(() => {});
        sandbox.stub(ts, "isTemplateExpression").returns(true);
        sandbox.stub(ts, "isBinaryExpression").returns(true);
        sandbox.stub(ts, "isPropertyAccessExpression").returns(false);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          detector["findExcelA1NotationInStringInterpolation"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
        Reflect.set(detector, "isValidExcelA1Notation", backupFunc);
      });

      it("runtime issue: findExcelA1NotationInStringInterpolation - Condition 5", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");
        const backupFunc = Reflect.get(detector, "isValidExcelA1Notation");

        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            getStart: () => 0,
            getText: () => "main1",
            head: { text: "main1" },
            templateSpans: [
              {
                expression: {
                  getFullText: () => "main1",
                  operatorToken: { kind: ts.SyntaxKind.PlusToken },
                  left: { getFullText: () => "main1" },
                  right: { getFullText: () => "main1" },
                },
              },
            ],
            name: { text: "main1" },
            children: [
              {
                name: { text: "main1" },
                getStart: () => 0,
                getText: () => "main1",
                head: { text: "main1" },
                templateSpans: [
                  {
                    expression: {
                      getFullText: () => "main1",
                      operatorToken: { kind: ts.SyntaxKind.PlusToken },
                      left: { getFullText: () => "main1" },
                      right: { getFullText: () => "main1" },
                    },
                  },
                ],
              },
              {
                name: { text: "main1" },
                getStart: () => 0,
                getText: () => "main1",
                head: { text: "main1" },
                templateSpans: [
                  {
                    expression: {
                      getFullText: () => "main1",
                      operatorToken: { kind: ts.SyntaxKind.PlusEqualsToken },
                      left: { getFullText: () => "main1" },
                      right: { getFullText: () => "main1" },
                    },
                  },
                ],
              },
              { name: undefined },
            ],
          }),
        });
        Reflect.set(detector, "typeChecker", {
          getTypeAtLocation: () => ({ isNumberLiteral: () => false }),
          typeToString: () => "string",
        });
        Reflect.set(detector, "isValidExcelA1Notation", () => true);

        sandbox.stub(ts, "forEachChild").callsFake((node, visitNode) => {
          const t = node as any;
          if (t.children) t.children.forEach(visitNode);
          if (!t.name) throw new Error("expression is undefined");
        });
        sandbox.stub(ts, "isTemplateExpression").returns(true);
        sandbox.stub(ts, "isBinaryExpression").onCall(0).returns(true).onCall(1).returns(false);
        sandbox
          .stub(ts, "isPropertyAccessExpression")
          .onCall(0)
          .returns(true)
          .onCall(1)
          .returns(false)
          .onCall(2)
          .returns(false);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          detector["findExcelA1NotationInStringInterpolation"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
        Reflect.set(detector, "isValidExcelA1Notation", backupFunc);
      });

      it("runtime issue: findExcelA1NotationInStringInterpolation - Condition 6", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");
        const backupFunc = Reflect.get(detector, "isValidExcelA1Notation");

        Reflect.set(detector, "program", {
          getSourceFile: () => ({ head: { text: "main1" } }),
        });
        Reflect.set(detector, "typeChecker", {});
        Reflect.set(detector, "isValidExcelA1Notation", () => false);

        sandbox.stub(ts, "forEachChild").callsFake(() => {});
        sandbox.stub(ts, "isTemplateExpression").returns(true);
        sandbox.stub(ts, "isBinaryExpression").returns(true);
        sandbox.stub(ts, "isPropertyAccessExpression").returns(false);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          detector["findExcelA1NotationInStringInterpolation"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
        Reflect.set(detector, "isValidExcelA1Notation", backupFunc);
      });

      // eslint-disable-next-line no-secrets/no-secrets
      it("runtime issue: findExcelA1NotationInAllStringLiteral", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;
        const backupProgram = Reflect.get(detector, "program");
        const backupTypeChecker = Reflect.get(detector, "typeChecker");
        const backupFunc = Reflect.get(detector, "isValidExcelA1Notation");

        Reflect.set(detector, "program", {
          getSourceFile: () => ({
            getLineAndCharacterOfPosition: () => ({ line: 1, character: 1 }),
            getStart: () => 0,
            name: { text: "main1" },
            head: { text: "main1" },
          }),
        });
        Reflect.set(detector, "typeChecker", {});
        Reflect.set(detector, "isValidExcelA1Notation", () => true);

        sandbox.stub(ts, "forEachChild").callsFake(() => {
          throw new Error("expression is undefined");
        });
        sandbox.stub(ts, "isStringLiteral").returns(true);
        sandbox.stub(console, "error").callsFake(() => {});

        try {
          // Hack to direct call private methond
          // eslint-disable-next-line no-secrets/no-secrets
          detector["findExcelA1NotationInAllStringLiteral"]();
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
        Reflect.set(detector, "program", backupProgram);
        Reflect.set(detector, "typeChecker", backupTypeChecker);
        Reflect.set(detector, "isValidExcelA1Notation", backupFunc);
      });

      it("runtime issue: isValidExcelA1Notation", () => {
        const detector = CodeIssueDetector.getInstance();
        let err = undefined;

        try {
          // Hack to direct call private methond
          detector["isValidExcelA1Notation"]("A23:TK66");
          detector["isValidExcelA1Notation"]("A23");
          detector["isValidExcelA1Notation"](":");
        } catch (e) {
          err = e;
        }

        chai.assert.isUndefined(err);
      });
    });
  });
});
