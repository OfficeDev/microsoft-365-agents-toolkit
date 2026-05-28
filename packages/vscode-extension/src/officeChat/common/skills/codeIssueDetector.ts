// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import ts = require("typescript");
import { fetchRawFileContent } from "../utils";
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
  MeasurementCompilieErrorPropertyDoesNotExistOnTypeWithSuggestionsCount,
  MeasurementCompilieErrorTopLevelExpressionForbidenCount,
  MeasurementCompilieErrorTypeIsNotAssignableToTypeCount,
} from "../telemetryConsts";
import { ChatResponseStream } from "vscode";
import stringSimilarity = require("string-similarity");
import {
  getFixSuggestionArgumentCountMismatchGeneral,
  getFixSuggestionArgumentCountMismatchHasSignature,
  getFixSuggestionArgumentCountMismatchWithoutSignature,
  getFixSuggestionArgumentTypeMismatchGeneral,
  getFixSuggestionArgumentTypeMismatchWithDeclaration,
  getFixSuggestionArgumentTypeMismatchWithTypeDetail,
  getFixSuggestionCannotAssignToReadOnlyProperty,
  getFixSuggestionCannotFindModule,
  getFixSuggestionCannotFindName,
  getFixSuggestionConvertTypeToTypeMistake,
  getFixSuggestionExcelA1NotationInStringInterpolationBinaryExpressionGeneral,
  getFixSuggestionExcelA1NotationInStringInterpolationBinaryExpressionLeftNumberLiteral,
  getFixSuggestionExcelA1NotationInStringInterpolationBinaryExpressionRightNumberLiteral,
  getFixSuggestionExcelA1NotationInStringInterpolationPropertyAccess,
  getFixSuggestionExcelA1NotationInStringLiteralGeneral,
  getFixSuggestionExpressionExpectedHandlder,
  getFixSuggestionNoFunctionReturnOrNoimplementation,
  getFixSuggestionOperatorAddOnTypeMismatch,
  getFixSuggestionOverloadMismatchGeneral,
  getFixSuggestionOverloadMismatchWithDeclaration,
  getFixSuggestionPropertyDoesNotExistOnTypeFoundCandidateOfFixing,
  getFixSuggestionPropertyDoesNotExistOnTypeFoundConcreateMembership,
  getFixSuggestionPropertyDoesNotExistOnTypeFoundGeneralSuggestion,
  getFixSuggestionPropertyDoesNotExistOnTypeNoDetailSuggestion,
  getFixSuggestionPropertyDoesNotExistOnTypeUnionTypePrompt,
  getFixSuggestionTopLevelExpressionForbiden,
  getFixSuggestionTypeIsNotAssignableToType,
  getSuggestionOnAPIObjectPropertyAccessBeforeLoad,
  getSuggestionOnExcelA1NotationInStringConcatenationLeft,
  getSuggestionOnExcelA1NotationInStringConcatenationRight,
} from "../../officePrompts";

export class DetectionResult {
  public compileErrors: string[] = [];
  public runtimeErrors: string[] = [];
  public references: string[] = [];

  public merge(result: DetectionResult): void {
    this.compileErrors = this.compileErrors.concat(result.compileErrors);
    this.references = this.references.concat(result.references);
    this.runtimeErrors = this.runtimeErrors.concat(result.runtimeErrors);
  }

  public areSame(result: DetectionResult): boolean {
    return (
      this.compileErrors.length === result.compileErrors.length &&
      this.compileErrors.every((v, i) => v === result.compileErrors[i]) &&
      result.compileErrors.every((v, i) => v === this.compileErrors[i]) &&
      this.runtimeErrors.length === result.runtimeErrors.length &&
      this.runtimeErrors.every((v, i) => v === result.runtimeErrors[i]) &&
      result.runtimeErrors.every((v, i) => v === this.runtimeErrors[i]) &&
      this.references.length === result.references.length
    );
  }
}

export class CodeIssueDetector {
  static SOURCE_FILE_NAME = "source.ts";
  static DECLARATION_FILE_NAME = "office-js.d.ts";
  private static instance: CodeIssueDetector;
  private definionFile: ts.SourceFile | undefined;
  private program: ts.Program | undefined;
  private typeChecker: ts.TypeChecker | undefined;
  private completeMemberNames: string[] = [];

  private constructor() {}

  public static getInstance(): CodeIssueDetector {
    if (!CodeIssueDetector.instance) {
      CodeIssueDetector.instance = new CodeIssueDetector();
    }
    return CodeIssueDetector.instance;
  }

  public async detectIssuesAsync(
    response: ChatResponseStream,
    host: string,
    isCustomFunction: boolean,
    codeSnippet: string,
    telemetryData: {
      properties: { [key: string]: string };
      measurements: { [key: string]: number };
    }
  ): Promise<DetectionResult> {
    const result = new DetectionResult();
    // order is matther, don't swith the order
    await this.buildTypeDefAst(host);
    this.buildProgram(codeSnippet);
    this.typeChecker = this.program?.getTypeChecker();
    result.merge(this.getCompilationErrorsAsync(host, isCustomFunction, telemetryData));
    result.merge(this.getPotentialRuntimeIssues(host, isCustomFunction, telemetryData));

    return result;
  }

  private async buildTypeDefAst(host: string): Promise<void> {
    if (!this.definionFile) {
      const typeDefStr = await fetchRawFileContent(
        `https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts`
      );
      this.definionFile = ts.createSourceFile(
        CodeIssueDetector.DECLARATION_FILE_NAME,
        typeDefStr,
        ts.ScriptTarget.Latest,
        true
      );

      // Add this condition to check if self.definionFile is defined
      ts.forEachChild(this.definionFile, (node) => {
        const names = this.processNamespace(host, null, node);
        names?.forEach((name) => {
          this.completeMemberNames.push(name);
        });
      });
    }
  }

  private buildProgram(codeSnippet: string): void {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;
    // Add function definition to the code
    const code = `
    /// <reference path="${CodeIssueDetector.DECLARATION_FILE_NAME}" />

    ${codeSnippet}
    `;

    // Create a compiler host
    function createCustomCompilerHost(originalHost: ts.CompilerHost): ts.CompilerHost {
      return {
        ...originalHost,
        getSourceFile: (fileName, languageVersion, onError, shouldCreateNewSourceFile) => {
          if (fileName === CodeIssueDetector.SOURCE_FILE_NAME) {
            return ts.createSourceFile(fileName, code, ts.ScriptTarget.ES2015, true);
          } else if (fileName === "office-js.d.ts") {
            return self.definionFile;
          } else {
            // For all other files, use the original getSourceFile method.
            const libSource = originalHost.getSourceFile(
              fileName,
              languageVersion,
              onError,
              shouldCreateNewSourceFile
            );
            return libSource;
          }
        },
      };
    }

    const compilerOptions: ts.CompilerOptions = {
      allowJs: true,
      checkJs: true,
      noEmitOnError: true,
      target: ts.ScriptTarget.ES2015,
      lib: ["lib.es2015.d.ts", "lib.dom.d.ts"],
    };

    const originalHost = ts.createCompilerHost(compilerOptions);
    const customHost = createCustomCompilerHost(originalHost);

    // Create a program
    self.program = ts.createProgram(
      [CodeIssueDetector.SOURCE_FILE_NAME],
      compilerOptions,
      customHost
    );
  }

  // #region Compilation Error and suggestion Detection
  public getCompilationErrorsAsync(
    host: string,
    isCustomFunction: boolean,
    telemetryData: {
      properties: { [key: string]: string };
      measurements: { [key: string]: number };
    }
  ): DetectionResult {
    const result = new DetectionResult();
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;
    if (!self.program) {
      // TODO: log error in telemetry
      return result;
    }
    const diagnostics = ts.getPreEmitDiagnostics(self.program);

    diagnostics.forEach((diagnostic) => {
      if (diagnostic.file) {
        const { line, character } = diagnostic.file.getLineAndCharacterOfPosition(
          diagnostic.start || 0
        );
        const message = ts.flattenDiagnosticMessageText(diagnostic.messageText, "\n");
        const node = self.findNodeAtPosition(diagnostic.file, line, character);

        let lineText = "";
        let charStart = 0;
        let charEnd = 0;
        if (node) {
          charStart = diagnostic.file.getLineStarts()[line];
          charEnd = diagnostic.file.getLineEndOfPosition(node.getEnd());
          lineText = diagnostic.file.text.substring(charStart, charEnd);
        }

        const errorTreatment = self.getErrorTreatment(host, node, message, telemetryData);
        // let error = `Error: (line:${line + 1},character:${character + 1}):  ${message}`;
        let error = `Invalid code snippet at Char ${charStart}-${charEnd}:\n\`\`\`typescript\n${lineText}\n\`\`\`\n Error message:\n${message}`;
        if (errorTreatment) {
          error += `\nFix suggestion: ${errorTreatment}`;
        }
        error += "\n";
        result.compileErrors.push(error);
      }
    });

    return result;
  }

  private findNodeAtPosition(
    sourceFile: ts.SourceFile,
    line: number,
    character: number
  ): ts.Node | undefined {
    let foundNode: ts.Node | undefined = undefined;

    const position = ts.getPositionOfLineAndCharacter(sourceFile, line, character);

    function visit(node: ts.Node) {
      if (position >= node.getStart() && position < node.getEnd()) {
        foundNode = node;
        ts.forEachChild(node, visit);
      }
    }

    visit(sourceFile);
    return foundNode;
  }

  private getErrorTreatment(
    host: string,
    node: ts.Node | undefined,
    errorMsg: string,
    telemetryData: {
      properties: { [key: string]: string };
      measurements: { [key: string]: number };
    }
  ): string | undefined {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;
    let fixSuggestion: string | undefined;
    const treatments: {
      checker: (error: string) => boolean;
      callback: (node: ts.Node, error: string) => string | undefined;
    }[] = [];
    errorMsg = errorMsg.trim().replace(/(\r\n|\n|\r)/gm, "");

    const propertyDoesNotExistOnTypeWithSuggestions = {
      checker: (error: string) => {
        return error.includes("Did you mean");
      },
      callback: (node: ts.Node, error: string) => {
        if (
          !telemetryData.measurements[
            MeasurementCompilieErrorPropertyDoesNotExistOnTypeWithSuggestionsCount
          ]
        ) {
          telemetryData.measurements[
            MeasurementCompilieErrorPropertyDoesNotExistOnTypeWithSuggestionsCount
          ] = 0;
        }
        telemetryData.measurements[
          MeasurementCompilieErrorPropertyDoesNotExistOnTypeWithSuggestionsCount
        ] += 1;
        const matches = error.match(
          /Property '([^']+)' does not exist on type '[^']+'. Did you mean '([^']+)'?/
        );
        if (matches) {
          const invalidProperty = matches[1];
          const suggestedProperty = matches[2];
          return `Change code to use '${suggestedProperty}' instead of '${invalidProperty}'.`;
        }
        return fixSuggestion; // something went wrong
      },
    };
    treatments.push(propertyDoesNotExistOnTypeWithSuggestions);

    const propertyDoesNotExistOnType = {
      checker: (error: string) => {
        return error.includes("does not exist on type ");
      },
      callback: (node: ts.Node, error: string) => {
        if (!telemetryData.measurements[MeasurementCompilieErrorPropertyDoesNotExistOnTypeCount]) {
          telemetryData.measurements[MeasurementCompilieErrorPropertyDoesNotExistOnTypeCount] = 0;
        }
        telemetryData.measurements[MeasurementCompilieErrorPropertyDoesNotExistOnTypeCount] += 1;
        const matches = error.match(/Property '([^']+)' does not exist on type '([^']+)'./);
        if (matches) {
          const invalidProperty = matches[1];
          let className = matches[2];
          className = className.replace("typeof", "").trim(); // some type names have 'typeof' prefix
          const singleTypes = className.split("|"); // some types are union types like 'string | number'
          if (singleTypes.length > 1) {
            return getFixSuggestionPropertyDoesNotExistOnTypeUnionTypePrompt(singleTypes);
          } else {
            const memberNames: string[] = [];
            if (self.definionFile) {
              // Add this condition to check if self.definionFile is defined
              ts.forEachChild(self.definionFile, (node) => {
                const names = self.processNamespace(host, className, node);
                names?.forEach((name) => {
                  memberNames.push(name);
                });
              });
            }
            if (memberNames.length === 0) {
              return getFixSuggestionPropertyDoesNotExistOnTypeNoDetailSuggestion(
                className,
                invalidProperty
              );
            }
            const localPropertyMethodNames =
              memberNames.map((name) => name.split("property/method:")[1] ?? "") || [];
            const truncated = stringSimilarity.findBestMatch(
              `${invalidProperty}`,
              localPropertyMethodNames
            ).bestMatch.target;
            const sortedSimilarStringsLocal: string = memberNames.find((name) => {
              return name.indexOf(truncated) >= 0;
            }) as string;
            const sortedSimilarStringsGlobal: string[] = stringSimilarity
              .findBestMatch(
                `${invalidProperty}`,
                self.completeMemberNames.map((name) => name.split("property/method:")[1].trim())
              )
              .ratings.map((rating, index) => {
                rating.target = self.completeMemberNames[index];
                return rating;
              })
              .filter((rating) => rating.rating > 0.35)
              .sort((a, b) => b.rating - a.rating)
              .slice(0, 2)
              .map((rating) => rating.target);
            const foundCandidate: boolean =
              sortedSimilarStringsGlobal.find((name) => {
                return name.indexOf(sortedSimilarStringsLocal) >= 0;
              }) !== undefined;

            if (foundCandidate) {
              const declarationWithComments = self.getDeclarationWithComments(
                host,
                sortedSimilarStringsLocal.split("property/method:")[0].trim(),
                sortedSimilarStringsLocal.split("property/method:")[1].trim()
              );
              return getFixSuggestionPropertyDoesNotExistOnTypeFoundConcreateMembership(
                className,
                invalidProperty,
                declarationWithComments.comments,
                declarationWithComments.declaration
              );
            } else {
              sortedSimilarStringsGlobal.unshift(sortedSimilarStringsLocal);
              const suggestioons = sortedSimilarStringsGlobal.map((suggestion, index) => {
                const declarationWithComments = self.getDeclarationWithComments(
                  host,
                  suggestion.split("property/method:")[0].trim(),
                  suggestion.split("property/method:")[1].trim()
                );
                return getFixSuggestionPropertyDoesNotExistOnTypeFoundCandidateOfFixing(
                  index,
                  declarationWithComments.class,
                  declarationWithComments.comments,
                  declarationWithComments.declaration
                );
              });
              return getFixSuggestionPropertyDoesNotExistOnTypeFoundGeneralSuggestion(
                className,
                invalidProperty,
                suggestioons,
                memberNames
              );
            }
          }
        }
        return fixSuggestion; // something went wrong
      },
    };
    treatments.push(propertyDoesNotExistOnType);

    const noFunctionReturnOrNoimplementation = {
      checker: (error: string) => {
        return error.includes(
          "A function whose declared type is neither 'undefined', 'void', nor 'any' must return a value."
        );
      },
      callback: (node: ts.Node, error: string) => {
        if (
          !telemetryData.measurements[
            MeasurementCompilieErrorPropertyDoesNotExistOnTypeWithSuggestionCount
          ]
        ) {
          telemetryData.measurements[
            MeasurementCompilieErrorPropertyDoesNotExistOnTypeWithSuggestionCount
          ] = 0;
        }
        telemetryData.measurements[
          MeasurementCompilieErrorPropertyDoesNotExistOnTypeWithSuggestionCount
        ] += 1;
        return getFixSuggestionNoFunctionReturnOrNoimplementation();
      },
    };
    treatments.push(noFunctionReturnOrNoimplementation);

    const cannotFindModule = {
      checker: (error: string) => {
        return error.includes("Cannot find module");
      },
      callback: (node: ts.Node, error: string) => {
        if (!telemetryData.measurements[MeasurementCompilieErrorCannotFindModuleCount]) {
          telemetryData.measurements[MeasurementCompilieErrorCannotFindModuleCount] = 0;
        }
        telemetryData.measurements[MeasurementCompilieErrorCannotFindModuleCount] += 1;
        return getFixSuggestionCannotFindModule();
      },
    };
    treatments.push(cannotFindModule);

    const argumentCountMismatch = {
      checker: (error: string) => {
        return error.includes("arguments, but got ");
      },
      callback: (node: ts.Node, error: string) => {
        if (!telemetryData.measurements[MeasurementCompilieErrorArgumentCountMismatchCount]) {
          telemetryData.measurements[MeasurementCompilieErrorArgumentCountMismatchCount] = 0;
        }
        telemetryData.measurements[MeasurementCompilieErrorArgumentCountMismatchCount] += 1;
        let suggestion = "";
        // Get the TypeChecker from the Program
        const checker = self.program?.getTypeChecker();

        // search up until we find the CallExpression
        while (node && !ts.isCallExpression(node)) {
          node = node.parent;
        }

        if (!node) {
          return;
        }
        const callExpression = node;

        if (!ts.isCallExpression(callExpression)) {
          return getFixSuggestionArgumentCountMismatchGeneral();
        }

        const expression = callExpression.expression;
        const symbol = checker?.getSymbolAtLocation(expression);

        if (symbol) {
          // Use the Symbol to get the declarations
          const declarations = symbol.getDeclarations();
          if (declarations && declarations.length > 0) {
            // Get the first declaration
            const declaration = declarations[0];
            // Get the signature of the declaration
            // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
            const signature = checker!.getSignatureFromDeclaration(
              declaration as ts.SignatureDeclaration
            );

            if (signature) {
              // Get the number of parameters in the signature
              const expected = signature.parameters.length;
              // Get the number of arguments in the CallExpression
              const actual = callExpression.arguments.length;
              suggestion = getFixSuggestionArgumentCountMismatchHasSignature(
                expected,
                actual,
                signature.getDeclaration().getText()
              );
            } else {
              suggestion = getFixSuggestionArgumentCountMismatchWithoutSignature(
                declaration.getText()
              );
            }
            return suggestion;
          }
        }

        return getFixSuggestionArgumentCountMismatchGeneral();
      },
    };
    treatments.push(argumentCountMismatch);

    const argumentTypeMismatch = {
      checker: (error: string) => {
        return error.includes("Argument of type");
      },
      callback: (node: ts.Node, error: string) => {
        if (!telemetryData.measurements[MeasurementCompilieErrorArgumentTypeMismatchCount]) {
          telemetryData.measurements[MeasurementCompilieErrorArgumentTypeMismatchCount] = 0;
        }
        telemetryData.measurements[MeasurementCompilieErrorArgumentTypeMismatchCount] += 1;
        let suggestion = "";
        // Get the TypeChecker from the Program
        const checker = self.program?.getTypeChecker();

        // search up until we find the CallExpression
        while (node && !ts.isCallExpression(node)) {
          node = node.parent;
        }
        const callExpression = node;

        if (ts.isCallExpression(callExpression)) {
          const expression = callExpression.expression;
          const symbol = checker?.getSymbolAtLocation(expression);

          if (symbol) {
            // Use the Symbol to get the declarations
            const declarations = symbol.getDeclarations();
            if (declarations && declarations.length > 0) {
              // Get the first declaration
              const declaration = declarations[0];
              suggestion = getFixSuggestionArgumentTypeMismatchWithDeclaration(
                declaration.getFullText()
              );
            }
          }
        } else {
          const matches = error.match(
            /Argument of type '([^']+)' is not assignable to parameter of type '([^']+)'./
          );
          if (matches) {
            const invalidType = matches[1];
            const validType = matches[2];
            // return `The given argument is unexpected. It could be used a wrong object, or you should use an alternative format of the object, in order to match the expected type '${validType}'.`;
            suggestion = getFixSuggestionArgumentTypeMismatchWithTypeDetail(invalidType, validType);
          } else {
            suggestion = getFixSuggestionArgumentTypeMismatchGeneral();
          }
        }

        return suggestion;
      },
    };
    treatments.push(argumentTypeMismatch);

    const operatorAddOnTypeMismatch = {
      checker: (error: string) => {
        return error.includes("Operator '+' cannot be applied to types");
      },
      callback: (node: ts.Node, error: string) => {
        if (!telemetryData.measurements[MeasurementCompilieErrorOperatorAddOnTypeMismatchCount]) {
          telemetryData.measurements[MeasurementCompilieErrorOperatorAddOnTypeMismatchCount] = 0;
        }
        telemetryData.measurements[MeasurementCompilieErrorOperatorAddOnTypeMismatchCount] += 1;
        return getFixSuggestionOperatorAddOnTypeMismatch();
      },
    };
    treatments.push(operatorAddOnTypeMismatch);

    const typeIsNotAssignableToType = {
      checker: (error: string) => {
        return error.includes("is not assignable to type");
      },
      callback: (node: ts.Node, error: string) => {
        if (!telemetryData.measurements[MeasurementCompilieErrorTypeIsNotAssignableToTypeCount]) {
          telemetryData.measurements[MeasurementCompilieErrorTypeIsNotAssignableToTypeCount] = 0;
        }
        telemetryData.measurements[MeasurementCompilieErrorTypeIsNotAssignableToTypeCount] += 1;
        return getFixSuggestionTypeIsNotAssignableToType();
      },
    };
    treatments.push(typeIsNotAssignableToType);

    const convertTypeToTypeMistake = {
      checker: (error: string) => {
        return error.includes(
          "may be a mistake because neither type sufficiently overlaps with the other"
        );
      },
      callback: (node: ts.Node, error: string) => {
        if (!telemetryData.measurements[MeasurementCompilieErrorConvertTypeToTypeMistakeCount]) {
          telemetryData.measurements[MeasurementCompilieErrorConvertTypeToTypeMistakeCount] = 0;
        }
        telemetryData.measurements[MeasurementCompilieErrorConvertTypeToTypeMistakeCount] += 1;
        return getFixSuggestionConvertTypeToTypeMistake();
      },
    };
    treatments.push(convertTypeToTypeMistake);

    const overloadMismatch = {
      checker: (error: string) => {
        return error.includes("No overload matches this call. Overload 1 of ");
      },
      callback: (node: ts.Node, error: string) => {
        if (!telemetryData.measurements[MeasurementCompilieErrorOverloadMismatchCount]) {
          telemetryData.measurements[MeasurementCompilieErrorOverloadMismatchCount] = 0;
        }
        telemetryData.measurements[MeasurementCompilieErrorOverloadMismatchCount] += 1;
        let suggestion = "";
        // Get the TypeChecker from the Program
        const checker = self.program?.getTypeChecker();

        // search up until we find the CallExpression
        while (node && !ts.isCallExpression(node)) {
          node = node.parent;
        }
        const callExpression = node;

        if (ts.isCallExpression(callExpression)) {
          const expression = callExpression.expression;
          const symbol = checker?.getSymbolAtLocation(expression);

          if (symbol) {
            // Use the Symbol to get the declarations
            const declarations = symbol.getDeclarations();
            if (declarations && declarations.length > 0) {
              // Get the first declaration
              const declaration = declarations[0];
              suggestion = getFixSuggestionOverloadMismatchWithDeclaration(
                declaration.getFullText()
              );
            }
          }
        } else {
          const regex = /Overload (\d+) of (\d+), '([^']+)', gave the following error./;
          const match = error.match(regex);

          if (match) {
            // let currentOverload = match[1];
            // let inTotalOverload = match[2];
            const methodDeclaration = match[3];
            suggestion = getFixSuggestionOverloadMismatchWithDeclaration(methodDeclaration);
          } else {
            suggestion = getFixSuggestionOverloadMismatchGeneral();
          }
        }

        return suggestion;
      },
    };
    treatments.push(overloadMismatch);

    const cannotFindName = {
      checker: (error: string) => {
        return error.includes("Cannot find name");
      },
      callback: (node: ts.Node, error: string) => {
        if (!telemetryData.measurements[MeasurementCompilieErrorCannotFindNameCount]) {
          telemetryData.measurements[MeasurementCompilieErrorCannotFindNameCount] = 0;
        }
        telemetryData.measurements[MeasurementCompilieErrorCannotFindNameCount] += 1;
        return getFixSuggestionCannotFindName();
      },
    };
    treatments.push(cannotFindName);

    const cannotAssignToReadOnlyProperty = {
      checker: (error: string) => {
        return error.includes("Cannot assign to");
      },
      callback: (node: ts.Node, error: string) => {
        if (
          !telemetryData.measurements[MeasurementCompilieErrorCannotAssignToReadOnlyPropertyCount]
        ) {
          telemetryData.measurements[
            MeasurementCompilieErrorCannotAssignToReadOnlyPropertyCount
          ] = 0;
        }
        telemetryData.measurements[
          MeasurementCompilieErrorCannotAssignToReadOnlyPropertyCount
        ] += 1;
        return getFixSuggestionCannotAssignToReadOnlyProperty();
      },
    };
    treatments.push(cannotAssignToReadOnlyProperty);

    const topLevelExpressionForbiden = {
      checker: (error: string) => {
        return error.includes(
          "expressions are only allowed at the top level of a file when that file is a module"
        );
      },
      callback: (node: ts.Node, error: string) => {
        if (!telemetryData.measurements[MeasurementCompilieErrorTopLevelExpressionForbidenCount]) {
          telemetryData.measurements[MeasurementCompilieErrorTopLevelExpressionForbidenCount] = 0;
        }
        telemetryData.measurements[MeasurementCompilieErrorTopLevelExpressionForbidenCount] += 1;
        return getFixSuggestionTopLevelExpressionForbiden();
      },
    };
    treatments.push(topLevelExpressionForbiden);

    const expressionExpectedHandlder = {
      checker: (error: string) => {
        return error.includes("Expression expected");
      },
      callback: (node: ts.Node, error: string) => {
        if (!telemetryData.measurements[MeasurementCompilieErrorExpressionExpectedCount]) {
          telemetryData.measurements[MeasurementCompilieErrorExpressionExpectedCount] = 0;
        }
        telemetryData.measurements[MeasurementCompilieErrorExpressionExpectedCount] += 1;
        return getFixSuggestionExpressionExpectedHandlder();
      },
    };
    treatments.push(expressionExpectedHandlder);

    const treatment = treatments.find((t) => t.checker(errorMsg));
    if (treatment && node) {
      fixSuggestion = treatment.callback(node, errorMsg);
    } else {
      if (!telemetryData.measurements[MeasurementCompilieErrorOthersCount]) {
        telemetryData.measurements[MeasurementCompilieErrorOthersCount] = 0;
      }
      telemetryData.measurements[MeasurementCompilieErrorOthersCount] += 1;
    }

    return fixSuggestion;
  }

  private getMethodsAndProperties(classname: string | null, node: ts.Node): string[] {
    if (
      ts.isClassDeclaration(node) && !!classname
        ? node.name && node.name.getText() === classname
        : true
    ) {
      try {
        const declaredClassName = (node as ts.ClassDeclaration).name?.getText() || classname || "";
        const members = (node as ts.ClassDeclaration).members;
        if (!members) {
          return [];
        }
        const memberNames = members
          .map((member) => {
            if (ts.isMethodDeclaration(member) || ts.isPropertyDeclaration(member)) {
              return `class: ${declaredClassName}, property/method: ${member.name.getText()}`;
            }
            return undefined;
          })
          .filter((name): name is string => name !== undefined); // filter out undefined values
        return memberNames;
      } catch (error) {
        console.error("getMethodsAndProperties:" + (error as Error).toString());
      }
    }
    return [];
  }

  private getDeclarationWithComments(moduleName: string, className: string, memberName: string) {
    const sourceFile = this.definionFile;

    let declaration: ts.Node | undefined;
    let comments: string | undefined;

    function visit(node: ts.Node) {
      if (!declaration && ts.isModuleDeclaration(node) && node.name.getText() === moduleName) {
        ts.forEachChild(node, visit);
      } else if (
        !declaration &&
        ts.isClassDeclaration(node) &&
        node.name?.getText() === className
      ) {
        ts.forEachChild(node, visit);
      } else if (
        !declaration &&
        (ts.isPropertyDeclaration(node) || ts.isMethodDeclaration(node)) &&
        node.name.getText() === memberName
      ) {
        declaration = node;
        const commentRanges = ts.getLeadingCommentRanges(sourceFile!.text, node.pos);
        comments = commentRanges
          ? commentRanges
              .map((range) => sourceFile!.text.substring(range.pos, range.end).trim())
              .join("\n")
          : undefined;
      } else {
        ts.forEachChild(node, visit);
      }
    }

    ts.forEachChild(sourceFile!, visit);

    return { class: className, declaration: declaration?.getFullText(), comments };
  }

  private processNamespace(namespace: string, classname: string | null, node: ts.Node) {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;
    if (ts.isModuleDeclaration(node) && node.name && node.name.getText() == namespace) {
      // a namespace is a "module" in the AST
      const memberNames: string[] = [];
      ts.forEachChild(node, (childNode) => {
        if (ts.isModuleBlock(childNode)) {
          ts.forEachChild(childNode, (node) => {
            const names = self.getMethodsAndProperties(classname, node);
            names?.forEach((name) => {
              if (name) {
                memberNames.push(name);
              }
            });
          });
        }
      });
      return memberNames;
    }
    return null;
  }
  // #endregion

  // #region Styling Error and suggestion Detection
  public getPotentialRuntimeIssues(
    host: string,
    isCustomFunction: boolean,
    telemetryData: {
      properties: { [key: string]: string };
      measurements: { [key: string]: number };
    }
  ): DetectionResult {
    const result = new DetectionResult();
    if (!isCustomFunction) {
      result.merge(this.findEntryFunctionInGeneratedCode());
      // result.merge(this.findMainFunctionInvoke());
    }
    result.merge(this.findImportAndRequireStatements());
    result.merge(this.findPropertyAccessAfterCallExpression(host));
    result.merge(this.findOfficeAPIObjectPropertyAccess(host));
    result.merge(this.findExcelA1NotationInStringConcatenation());
    result.merge(this.findExcelA1NotationInStringInterpolation());
    result.merge(this.findExcelA1NotationInAllStringLiteral());
    return result;
  }

  private findImportAndRequireStatements(): DetectionResult {
    const result = new DetectionResult();

    if (!this.program) {
      return result;
    }
    const sourceFile = this.program.getSourceFile(CodeIssueDetector.SOURCE_FILE_NAME);
    if (!sourceFile || !this.typeChecker) {
      return result;
    }

    function visitNode(node: ts.Node) {
      if (
        sourceFile &&
        (ts.isImportDeclaration(node) ||
          ((ts.isVariableStatement(node) || ts.isExpressionStatement(node)) &&
            node.getText().includes("require(")))
      ) {
        {
          const line = sourceFile.getLineAndCharacterOfPosition(node.getStart()).line + 1;
          const warningMsg = `Error: Find "import" or "require" statement at line ${line}.`;
          const fixSuggestion = `Fix suggestion: Use mockup object or interface for dependencies.`;
          const warning = `${warningMsg} ${fixSuggestion}`;
          result.compileErrors.push(warning);
        }

        ts.forEachChild(node, visitNode);
      }
    }

    ts.forEachChild(sourceFile, visitNode);
    return result;
  }

  private findEntryFunctionInGeneratedCode(): DetectionResult {
    const result = new DetectionResult();

    if (!this.program) {
      return result;
    }
    const sourceFile = this.program.getSourceFile(CodeIssueDetector.SOURCE_FILE_NAME);
    if (!sourceFile || !this.typeChecker) {
      return result;
    }
    let foundTheMainFunction = false;
    let mainFunctionHasValidSignature = false;
    let definedAsAsync = false;
    function visit(node: ts.Node, checker: ts.TypeChecker) {
      // try to cover the arrow function, function expresson.
      if (
        ts.isFunctionDeclaration(node) ||
        ts.isArrowFunction(node) ||
        ts.isFunctionExpression(node)
      ) {
        const name = ts.isFunctionDeclaration(node)
          ? node.name?.getText()
          : node.parent?.getText().split(" ")[1];
        if (name === "main") {
          const isAsync = node.modifiers?.some(
            (modifier) => modifier.kind === ts.SyntaxKind.AsyncKeyword
          );
          const hasNoArguments = node.parameters.length === 0;
          foundTheMainFunction = true;
          mainFunctionHasValidSignature = hasNoArguments;
          definedAsAsync = !!isAsync;
        }
      }
      ts.forEachChild(node, (child) => visit(child, checker));
    }
    try {
      visit(sourceFile, this.typeChecker);

      if (!foundTheMainFunction) {
        const warningMsg = `Error: Entry function 'main' not found in the code. The entry function 'main' is the starting point of the code execution. It may missed, or has another name.`;
        const fixSuggestion = `Fix suggestion: Add a function named 'main' as the entry point of the code, wrap existing function call in right order.`;
        const warning = `${warningMsg} ${fixSuggestion}`;
        result.compileErrors.push(warning);
      } else {
        if (!mainFunctionHasValidSignature) {
          const warningMsg = `Error: Entry function 'main' has invalid signature. The entry function 'main' must not have any parameter.`;
          const fixSuggestion = `Fix suggestion: Remove the parameters from the 'main' function, and make sure it has no parameter.`;
          const warning = `${warningMsg} ${fixSuggestion}`;
          result.compileErrors.push(warning);
        }
        if (!definedAsAsync) {
          const warningMsg = `Error: Entry function 'main' is not defined as async function. The entry function 'main' must be defined as an async function.`;
          const fixSuggestion = `Fix suggestion: Add 'async' keyword before the 'main' function declaration to define it as an async function.`;
          const warning = `${warningMsg} ${fixSuggestion}`;
          result.compileErrors.push(warning);
        }
      }
    } catch (error) {
      // eslint-disable-next-line @typescript-eslint/restrict-plus-operands, no-secrets/no-secrets
      console.error("findEntryFunctionInGeneratedCode:" + (error as Error).toString());
    }

    return result;
  }

  private findPropertyAccessAfterCallExpression(host: string): DetectionResult {
    const result = new DetectionResult();

    if (!this.program) {
      return result;
    }
    const sourceFile = this.program.getSourceFile(CodeIssueDetector.SOURCE_FILE_NAME);
    if (!sourceFile || !this.typeChecker) {
      return result;
    }
    function visit(node: ts.Node, checker: ts.TypeChecker) {
      if (
        !!node.parent &&
        !ts.isCallExpression(node.parent) &&
        ts.isPropertyAccessExpression(node) &&
        ts.isCallExpression(node.expression) &&
        !!sourceFile
      ) {
        const expressionStr = node.expression.getFullText();
        const propertyStr = node.name.getText();
        const line = sourceFile.getLineAndCharacterOfPosition(node.getStart()).line + 1;
        const warningMsg = `Error: PropertyAccessExpression after CallExpression: ${expressionStr}.${propertyStr} at line ${line}.`;
        const fixSuggestion = `Fix suggestion: The immediate property access after a function call is forbidden. You must store the result of the function call ${expressionStr} in a variable first, prefer in previous line. Then access the property ${propertyStr} from the variable in the next line.`;
        const warning = `${warningMsg} ${fixSuggestion}`;
        result.runtimeErrors.push(warning);
      }
      ts.forEachChild(node, (child) => visit(child, checker));
    }
    try {
      visit(sourceFile, this.typeChecker);
    } catch (error) {
      // eslint-disable-next-line @typescript-eslint/restrict-plus-operands, no-secrets/no-secrets
      console.error("findPropertyAccessAfterCallExpression:" + (error as Error).toString());
    }

    return result;
  }

  private findOfficeAPIObjectPropertyAccess(host: string): DetectionResult {
    const result = new DetectionResult();
    const sourceFile = this.program?.getSourceFile(CodeIssueDetector.SOURCE_FILE_NAME);
    if (!sourceFile || !this.typeChecker) {
      return result;
    }
    function visit(node: ts.Node, checker: ts.TypeChecker) {
      if (ts.isPropertyAccessExpression(node) && sourceFile) {
        const objectType = checker.getTypeAtLocation(node.expression);
        if (objectType?.symbol && objectType.symbol.escapedName.toString().startsWith(host)) {
          const accessObjStr = objectType.symbol.escapedName;
          const propertyStr = node.name.text;
          const line = sourceFile.getLineAndCharacterOfPosition(node.getStart()).line + 1;

          if (!accessObjStr) {
            const warningMsg = getSuggestionOnAPIObjectPropertyAccessBeforeLoad(
              accessObjStr.toString(),
              propertyStr,
              line
            );
            result.runtimeErrors.push(warningMsg);
          }
        }
      }
      ts.forEachChild(node, (child) => visit(child, checker));
    }
    try {
      visit(sourceFile, this.typeChecker);
    } catch (error) {
      // eslint-disable-next-line @typescript-eslint/restrict-plus-operands, no-secrets/no-secrets
      console.error("findOfficeAPIObjectPropertyAccess:" + (error as Error).toString());
    }

    return result;
  }

  private findExcelA1NotationInStringConcatenation(): DetectionResult {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;
    const result = new DetectionResult();
    const sourceFile = this.program?.getSourceFile(CodeIssueDetector.SOURCE_FILE_NAME);
    if (!sourceFile || !this.typeChecker) {
      return result;
    }
    function visit(node: ts.Node, checker: ts.TypeChecker) {
      if (ts.isBinaryExpression(node)) {
        if (ts.isStringLiteral(node.left) && self.isValidExcelA1Notation(node.left.text)) {
          const rightType = checker.getTypeAtLocation(node.right);
          if (checker.typeToString(rightType) === "number" && !!sourceFile) {
            const line = sourceFile.getLineAndCharacterOfPosition(node.getStart()).line + 1;
            const warningMsg = getSuggestionOnExcelA1NotationInStringConcatenationRight(
              node.getText(),
              line,
              node.right.getFullText()
            );
            result.runtimeErrors.push(warningMsg);
          }
        } else if (ts.isStringLiteral(node.right) && self.isValidExcelA1Notation(node.right.text)) {
          const leftType = checker.getTypeAtLocation(node.left);
          if (checker.typeToString(leftType) === "number" && !!sourceFile) {
            const line = sourceFile.getLineAndCharacterOfPosition(node.getStart()).line + 1;
            const warningMsg = getSuggestionOnExcelA1NotationInStringConcatenationLeft(
              node.getText(),
              line,
              node.left.getFullText()
            );
            result.runtimeErrors.push(warningMsg);
          }
        }
      }
      ts.forEachChild(node, (child) => visit(child, checker));
    }
    try {
      visit(sourceFile, this.typeChecker);
    } catch (error) {
      // eslint-disable-next-line @typescript-eslint/restrict-plus-operands, no-secrets/no-secrets
      console.error("findExcelA1NotationInStringConcatenation:" + (error as Error).toString());
    }
    return result;
  }

  private findExcelA1NotationInStringInterpolation(): DetectionResult {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;
    const result = new DetectionResult();
    const sourceFile = this.program?.getSourceFile(CodeIssueDetector.SOURCE_FILE_NAME);
    if (!sourceFile || !this.typeChecker) {
      return result;
    }
    function visit(node: ts.Node, checker: ts.TypeChecker) {
      if (ts.isTemplateExpression(node)) {
        // target to all expression like: `A2:A${stockData.length + 1}`, `A2:A${stockData.length}`, `A2:A${1 + stockData.length}`
        const head = node.head.text;
        if (self.isValidExcelA1Notation(head)) {
          const span = node.templateSpans[0];
          if (ts.isPropertyAccessExpression(span.expression)) {
            const expressionStr = span.expression.getFullText();
            const type = checker.getTypeAtLocation(span.expression.name);
            if (!!sourceFile && checker.typeToString(type) === "number") {
              const line = sourceFile.getLineAndCharacterOfPosition(node.getStart()).line + 1;
              const warningMsg = getFixSuggestionExcelA1NotationInStringInterpolationPropertyAccess(
                node.getText(),
                line,
                expressionStr
              );
              result.runtimeErrors.push(warningMsg);
            }
          } else if (
            ts.isBinaryExpression(span.expression) &&
            (span.expression.operatorToken.kind === ts.SyntaxKind.PlusToken ||
              span.expression.operatorToken.kind === ts.SyntaxKind.MinusToken)
          ) {
            const leftType = checker.getTypeAtLocation(span.expression.left);
            const rightType = checker.getTypeAtLocation(span.expression.right);
            const expressionStr = span.expression.getFullText();
            if (
              checker.typeToString(leftType) === "number" &&
              rightType.isNumberLiteral() &&
              !!sourceFile
            ) {
              const line = sourceFile.getLineAndCharacterOfPosition(node.getStart()).line + 1;
              const warningMsg =
                getFixSuggestionExcelA1NotationInStringInterpolationBinaryExpressionLeftNumberLiteral(
                  node.getText(),
                  line,
                  expressionStr,
                  rightType.value.toString(),
                  span.expression.left.getFullText()
                );
              result.runtimeErrors.push(warningMsg);
            } else if (
              checker.typeToString(rightType) === "number" &&
              leftType.isNumberLiteral() &&
              !!sourceFile
            ) {
              const line = sourceFile.getLineAndCharacterOfPosition(node.getStart()).line + 1;
              const warningMsg =
                getFixSuggestionExcelA1NotationInStringInterpolationBinaryExpressionRightNumberLiteral(
                  node.getText(),
                  line,
                  expressionStr,
                  leftType.value.toString(),
                  span.expression.right.getFullText()
                );
              result.runtimeErrors.push(warningMsg);
            } else {
              // eslint-disable-next-line @typescript-eslint/no-non-null-assertion
              const line = sourceFile!.getLineAndCharacterOfPosition(node.getStart()).line + 1;
              const warningMsg =
                getFixSuggestionExcelA1NotationInStringInterpolationBinaryExpressionGeneral(
                  node.getText(),
                  line,
                  expressionStr,
                  span.expression.right.getFullText(),
                  span.expression.left.getFullText()
                );
              result.runtimeErrors.push(warningMsg);
            }
          }
        }
      }
      ts.forEachChild(node, (child) => visit(child, checker));
    }
    try {
      visit(sourceFile, this.typeChecker);
    } catch (error) {
      // eslint-disable-next-line @typescript-eslint/restrict-plus-operands, no-secrets/no-secrets
      console.error("findExcelA1NotationInStringInterpolation:" + (error as Error).toString());
    }
    return result;
  }

  private findExcelA1NotationInAllStringLiteral(): DetectionResult {
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;
    const result = new DetectionResult();
    const sourceFile = this.program?.getSourceFile(CodeIssueDetector.SOURCE_FILE_NAME);
    if (!sourceFile || !this.typeChecker) {
      return result;
    }
    function visit(node: ts.Node, checker: ts.TypeChecker): void {
      if (sourceFile && ts.isStringLiteral(node) && self.isValidExcelA1Notation(node.text)) {
        const line = sourceFile.getLineAndCharacterOfPosition(node.getStart()).line + 1;
        const warningMsg = getFixSuggestionExcelA1NotationInStringLiteralGeneral(node.text, line);
        result.runtimeErrors.push(warningMsg);
      }
      ts.forEachChild(node, (child) => visit(child, checker));
    }

    try {
      visit(sourceFile, this.typeChecker);
    } catch (error) {
      // eslint-disable-next-line @typescript-eslint/restrict-plus-operands, no-secrets/no-secrets
      console.error("findExcelA1NotationInAllStringLiteral:" + (error as Error).toString());
    }
    return result;
  }

  private columnToNumber(column: string) {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - "A".charCodeAt(0) + 1);
    }
    return result;
  }

  private isValidExcelA1Notation(range: string) {
    const match = range.match(/([A-Z]+)\d+(?::([A-Z]+)\d+)?/);
    if (!match) {
      return false;
    }
    if (match[2]) {
      const firstColumn = this.columnToNumber(match[1]);
      const secondColumn = this.columnToNumber(match[2]);
      return firstColumn <= secondColumn;
    }
    return true;
  }
  // #endregion
}
