// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  CancellationToken,
  ChatResponseStream,
  LanguageModelChatMessage,
  LanguageModelChatMessageRole,
} from "vscode";
import { CodeIssueDetector, DetectionResult } from "./codeIssueDetector";
import { ISkill } from "./iSkill"; // Add the missing import statement
import { Spec } from "./spec"; // Add the missing import statement
import { countMessagesTokens, getCopilotResponseAsString } from "../../../chat/utils";
import { ExecutionResultEnum } from "./executionResultEnum";
import {
  MeasurementSystemSelfReflectionAttemptCount,
  MeasurementSystemSelfReflectionAttemptSucceeded,
  MeasurementSelfReflectionExecutionTimeInTotalSec,
  MeasurementErrorsAfterCorrection,
} from "../telemetryConsts";
import {
  customFunctionSystemPrompt,
  excelSystemPrompt,
  getCodeSamplePrompt,
  getDeclarationsPrompt,
  getFixIssueDefaultSystemPrompt,
  getFixIssueUserPrompt,
} from "../../officePrompts";
import { localize } from "../../../utils/localizeUtils";
import { getTokenLimitation } from "../../consts";
import { SampleData } from "../samples/sampleData";
// import { writeLogToFile } from "../utils";

export class CodeIssueCorrector implements ISkill {
  static MAX_TRY_COUNT = 10; // From the observation from a small set of test, fix over 2 rounds leads to worse result, set it to a smal number so we can fail fast
  name: string;
  capability: string;

  constructor() {
    this.name = "codeIssueCorrector";
    this.capability = "Fix code issues";
  }

  public canInvoke(spec: Spec): boolean {
    return (
      !!spec.appendix.host &&
      !!spec.appendix.codeSnippet &&
      !!spec.appendix.codeTaskBreakdown &&
      spec.appendix.codeTaskBreakdown.length > 0
    );
  }

  public async invoke(
    languageModel: LanguageModelChatMessage,
    response: ChatResponseStream,
    token: CancellationToken,
    spec: Spec
  ): Promise<{ result: ExecutionResultEnum; spec: Spec }> {
    const host = spec.appendix.host;
    let codeSnippet = spec.appendix.codeSnippet;
    const codeTaskBreakdown = spec.appendix.codeTaskBreakdown;
    const t = performance.now();
    let t0 = performance.now();
    let baseLineResuult: DetectionResult = await CodeIssueDetector.getInstance().detectIssuesAsync(
      response,
      host,
      spec.appendix.isCustomFunction,
      codeSnippet,
      spec.appendix.telemetryData
    );
    let t1 = performance.now();
    let duration = Math.ceil((t1 - t0) / 1000);
    console.debug(
      `Baseline: [C] ${baseLineResuult.compileErrors.length}, [R] ${baseLineResuult.runtimeErrors.length}. Detect spend ${duration} seconds.`
    );

    const model: "copilot-gpt-3.5-turbo" | "copilot-gpt-4" = "copilot-gpt-3.5-turbo";
    let maxRetryCount: number;
    let issueTolerance: number;

    if (spec.appendix.complexity < 25) {
      maxRetryCount = 2;
      issueTolerance = 1;
    } else if (spec.appendix.complexity < 50) {
      maxRetryCount = 2;
      issueTolerance = 1;
    } else if (spec.appendix.complexity < 75) {
      maxRetryCount = 2;
      issueTolerance = 1;
    } else {
      maxRetryCount = 2;
      issueTolerance = 1;
    }

    if (baseLineResuult.compileErrors.length === 0 && baseLineResuult.runtimeErrors.length === 0) {
      console.debug("No issue found in baseline, skip the self reflection.");
      spec.appendix.telemetryData.measurements[MeasurementErrorsAfterCorrection] =
        baseLineResuult.compileErrors.length;
      return { result: ExecutionResultEnum.Success, spec: spec };
    }
    if (baseLineResuult.compileErrors.length > issueTolerance) {
      // Don't waste time on low quality code, fail fast
      console.debug(
        `${baseLineResuult.compileErrors.length} compile errors in baseline code that beyond our tolerance ${issueTolerance}, skip the self reflection.`
      );
      spec.appendix.telemetryData.measurements[MeasurementErrorsAfterCorrection] =
        baseLineResuult.compileErrors.length;
      return { result: ExecutionResultEnum.FailedAndGoNext, spec: spec };
    }

    // const setDeclartionPrompt = getDeclarationsPrompt();

    //     if (!!spec.appendix.apiDeclarationsReference && !!spec.appendix.apiDeclarationsReference.size) {
    //       const groupedMethodsOrProperties = new Map<string, SampleData[]>();
    //       for (const methodOrProperty of spec.appendix.apiDeclarationsReference) {
    //         if (!groupedMethodsOrProperties.has(methodOrProperty[1].definition)) {
    //           groupedMethodsOrProperties.set(methodOrProperty[1].definition, [methodOrProperty[1]]);
    //         }
    //         groupedMethodsOrProperties.get(methodOrProperty[1].definition)?.push(methodOrProperty[1]);
    //       }

    //       let tempClassDeclaration = "";
    //       groupedMethodsOrProperties.forEach((methodsOrPropertiesCandidates, className) => {
    //         tempClassDeclaration += `
    // class ${className} extends OfficeExtension.ClientObject {
    //   ${methodsOrPropertiesCandidates.map((sampleData) => sampleData.codeSample).join("\n\n")}
    // }
    // \n\n
    //         `;
    //       });

    //       setDeclartionPrompt += `

    //       \`\`\`typescript
    //       ${tempClassDeclaration};
    //       \`\`\`

    //       Let's think step by step.
    //       `;
    //     }
    //     const declarationMessage: LanguageModelChatMessage | null =
    //       spec.appendix.apiDeclarationsReference.size > 0
    //         ? new LanguageModelChatMessage(LanguageModelChatMessageRole.System, setDeclartionPrompt)
    //         : null;

    const sampleMessage: LanguageModelChatMessage | null =
      spec.appendix.codeSample.length > 0
        ? new LanguageModelChatMessage(
            LanguageModelChatMessageRole.User,
            getCodeSamplePrompt(spec.appendix.codeSample)
          )
        : null;

    let fixedCode: string | null = codeSnippet;
    const historicalErrors: string[] = [];
    let additionalInfo = "";
    for (let index = 0; index < maxRetryCount; index++) {
      // if (baseLineResuult.compileErrors.length > maxRetryCount - index) {
      //   // Let's fail fast, as if the error is too many, it's hard to fix in a few rounds
      //   console.debug(
      //     `${baseLineResuult.compileErrors.length} compile errors need to fix in next ${
      //       maxRetryCount - index
      //     } rounds, fail fast.`
      //   );
      //   break;
      // }
      response.progress(
        localize("teamstoolkit.chatParticipants.officeAddIn.issueDetector.fixingErrors")
      );
      t0 = performance.now();
      fixedCode = await this.fixIssueAsync(
        token,
        host,
        spec.appendix.isCustomFunction,
        codeSnippet,
        codeTaskBreakdown,
        baseLineResuult.compileErrors,
        baseLineResuult.runtimeErrors,
        historicalErrors,
        additionalInfo,
        model,
        null, //declarationMessage,
        sampleMessage,
        spec
      );
      t1 = performance.now();
      duration = Math.ceil((t1 - t0) / 1000);
      console.debug(`Self reflection iteration ${index + 1}, takes ${duration} seconds.`);
      if (!fixedCode) {
        // something wrong, just to the next round
        continue;
      }
      t0 = performance.now();
      const issuesAfterFix: DetectionResult =
        await CodeIssueDetector.getInstance().detectIssuesAsync(
          response,
          host,
          spec.appendix.isCustomFunction,
          fixedCode,
          spec.appendix.telemetryData
        );
      historicalErrors.push(
        ...baseLineResuult.compileErrors.map(
          (item) => item.replace(/at Char \d+-\d+:/g, "").split("\nFix suggestion")[0]
        )
      );
      t1 = performance.now();
      duration = Math.ceil((t1 - t0) / 1000);
      console.debug(
        `After fix: [C] ${issuesAfterFix.compileErrors.length}, [R] ${issuesAfterFix.runtimeErrors.length}. Detect spend ${duration} seconds.`
      );
      // const now = new Date();
      // const nowStr = `${now.getHours()}h:${now.getMinutes()}m:${now.getSeconds()}s`;
      // await writeLogToFile(`\n[${nowStr}]\n`);
      // await writeLogToFile(
      //   "-------- Compile Errors ----------------------------------------------------------------------------------------\n" +
      //     issuesAfterFix.compileErrors.join("\n")
      // );
      // await writeLogToFile(
      //   "-------- Runtime Errors ----------------------------------------------------------------------------------------\n" +
      //     issuesAfterFix.runtimeErrors.join("\n")
      // );
      const terminateResult = this.terminateFixIteration(
        spec.appendix.complexity,
        codeSnippet,
        baseLineResuult,
        fixedCode,
        issuesAfterFix
      );
      if (terminateResult.terminate) {
        additionalInfo = terminateResult.suggestion;
        continue;
      }

      //#region telemetry
      t1 = performance.now();
      duration = Math.ceil((t1 - t) / 1000);
      if (
        !spec.appendix.telemetryData.measurements[MeasurementSelfReflectionExecutionTimeInTotalSec]
      ) {
        spec.appendix.telemetryData.measurements[MeasurementSelfReflectionExecutionTimeInTotalSec] =
          duration;
      } else {
        spec.appendix.telemetryData.measurements[
          MeasurementSelfReflectionExecutionTimeInTotalSec
        ] += duration;
      }
      // console.debug(`Self reflection completed within ${duration} seconds.`);

      if (!spec.appendix.telemetryData.measurements[MeasurementSystemSelfReflectionAttemptCount]) {
        spec.appendix.telemetryData.measurements[MeasurementSystemSelfReflectionAttemptCount] = 0;
      }
      spec.appendix.telemetryData.measurements[MeasurementSystemSelfReflectionAttemptCount] += 1;
      //#endregion
      // In ideal case, we expect the result match the base line, however, if that is the last round, we accept the result
      // perhaps without check the equivalence of the base line
      if (
        issuesAfterFix.compileErrors.length === 0 &&
        (index == maxRetryCount - 1 || issuesAfterFix.areSame(baseLineResuult))
      ) {
        // no more issue, return the fixed code
        // A dirty hacky to remove the invacation of main function if any because LLM may generate it and hard to remove it
        const regex = /(await\s)?main\(\)(\..+)?;/gm;
        const matches = fixedCode.match(regex);
        if (matches && matches.length > 0) {
          fixedCode = fixedCode.replace(matches[0], "");
        }
        spec.appendix.codeSnippet = fixedCode;
        spec.appendix.telemetryData.properties[MeasurementSystemSelfReflectionAttemptSucceeded] =
          "true";
        spec.appendix.telemetryData.measurements[MeasurementErrorsAfterCorrection] =
          issuesAfterFix.compileErrors.length;
        return { result: ExecutionResultEnum.Success, spec: spec };
      }

      // Prepare for next iteration
      codeSnippet = fixedCode;
      baseLineResuult = issuesAfterFix;
    }

    spec.appendix.codeSnippet = fixedCode || codeSnippet;
    spec.appendix.telemetryData.properties[MeasurementSystemSelfReflectionAttemptSucceeded] =
      "false";
    spec.appendix.telemetryData.measurements[MeasurementErrorsAfterCorrection] =
      baseLineResuult.compileErrors.length;
    return { result: ExecutionResultEnum.FailedAndGoNext, spec: spec };
  }

  async fixIssueAsync(
    token: CancellationToken,
    host: string,
    isCustomFunctions: boolean,
    codeSnippet: string,
    substeps: string[],
    errorMessages: string[],
    warningMessage: string[],
    historicalErrors: string[],
    additionalInfo: string,
    model: "copilot-gpt-3.5-turbo" | "copilot-gpt-4",
    declarationMessage: LanguageModelChatMessage | null,
    sampleMessage: LanguageModelChatMessage | null,
    spec: Spec
  ) {
    if (errorMessages.length === 0) {
      return codeSnippet;
    }
    const tempUserInput = getFixIssueUserPrompt(codeSnippet, additionalInfo, historicalErrors);

    const defaultSystemPrompt = getFixIssueDefaultSystemPrompt(
      host,
      substeps,
      errorMessages,
      warningMessage
    );

    let referenceUserPrompt = "";
    switch (host) {
      case "Excel":
        if (!isCustomFunctions) {
          referenceUserPrompt = excelSystemPrompt;
        } else {
          referenceUserPrompt = customFunctionSystemPrompt;
        }
        break;
      default:
        referenceUserPrompt = "";
        break;
    }

    // Perform the desired operation
    // The order in array is matter, don't change it unless you know what you are doing
    const messages: LanguageModelChatMessage[] = [
      new LanguageModelChatMessage(LanguageModelChatMessageRole.User, tempUserInput),
      new LanguageModelChatMessage(LanguageModelChatMessageRole.User, defaultSystemPrompt),
    ];

    if (!!sampleMessage) {
      messages.push(sampleMessage);
    }

    // if (!!declarationMessage) {
    //   messages.push(declarationMessage);
    // }

    messages.push(
      new LanguageModelChatMessage(LanguageModelChatMessageRole.User, referenceUserPrompt)
    );

    let msgCount = countMessagesTokens(messages);
    while (msgCount > getTokenLimitation(model)) {
      messages.pop();
      msgCount = countMessagesTokens(messages);
    }
    // console.debug(`token count: ${msgCount}, number of messages remains: ${messages.length}.`);
    const copilotResponse = await getCopilotResponseAsString(model, messages, token);
    spec.appendix.telemetryData.chatMessages.push(...messages);
    spec.appendix.telemetryData.responseChatMessages.push(
      new LanguageModelChatMessage(LanguageModelChatMessageRole.Assistant, copilotResponse)
    );
    // extract the code snippet
    const regex = /```[\s]*typescript([\s\S]*?)```/gm;
    const matches = regex.exec(copilotResponse);
    if (!matches) {
      // something wrong with the LLM output
      // TODO: Add handling for this case
      console.error(
        "[Code issue fix] Failed to extract the code snippet from the response:",
        copilotResponse
      );
      return null;
    }

    const newCodeStr = matches[matches.length - 1].trim();
    if (codeSnippet.length - newCodeStr.length > newCodeStr.length) {
      // The code length reduced too much
      console.debug("Code length reduced too much.");
      return null;
    }

    return newCodeStr;
  }

  private terminateFixIteration(
    complexityScore: number,
    baselineCodeStr: string,
    baselineResult: DetectionResult,
    currentCodeStr: string,
    currentResult: DetectionResult
  ): { terminate: boolean; suggestion: string } {
    const codeLengthDelta: number = currentCodeStr.length - baselineCodeStr.length;
    const compileErrorDelta: number =
      currentResult.compileErrors.length - baselineResult.compileErrors.length;

    if (codeLengthDelta < 0) {
      // The code length reduced
      if (Math.abs(codeLengthDelta) >= currentCodeStr.length) {
        // The code length reduced too much
        console.debug("Terminate: code length reduced too much.");
        return {
          terminate: true,
          suggestion: "You should send back with the whole snippets without any explanasion.",
        };
      }
    }

    if (compileErrorDelta > 0) {
      // fix a ge jimo
      console.debug("Terminate: compile error increased.");
      return {
        terminate: true,
        suggestion: "The previous fix introduced more compile error.",
      };
    }

    return {
      terminate: false,
      suggestion: "",
    };
  }
}
