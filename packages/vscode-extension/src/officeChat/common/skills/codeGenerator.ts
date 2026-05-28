// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  CancellationToken,
  ChatResponseStream,
  LanguageModelChatMessage,
  LanguageModelChatMessageRole,
} from "vscode";
import { correctPropertyLoadSpelling } from "../utils";
import { SampleProvider } from "../samples/sampleProvider";
import { ISkill } from "./iSkill"; // Add the missing import statement
import { Spec } from "./spec";
import { countMessagesTokens, getCopilotResponseAsString } from "../../../chat/utils";
import { ExecutionResultEnum } from "./executionResultEnum";
import {
  MeasurementCodeGenAttemptCount,
  MeasurementCodeGenExecutionTimeInTotalSec,
  PropertySystemCodeGenResult,
  MeasurementSystemCodegenTaskBreakdownAttemptFailedCount,
  MeasurementCodeGenTaskBreakdownTimeInTotalSec,
  MeasurementCodeGenPreScanTimeInTotalSec,
  MeasurementCodeGenGetSampleTimeInTotalSec,
} from "../telemetryConsts";
import {
  excelSystemPrompt,
  customFunctionSystemPrompt,
  getUserInputBreakdownTaskUserPrompt,
  getUserAskPreScanningSystemPrompt,
  getUserSimpleAskBreakdownTaskSystemPrompt,
  getGenerateCodeUserPrompt,
  getGenerateCodeSamplePrompt,
  getCodeSamplePrompt,
} from "../../officePrompts";
import { localize } from "../../../utils/localizeUtils";
import { getTokenLimitation } from "../../consts";
// import { SampleData } from "../samples/sampleData";
// import { DeclarationFinder } from "../declarationFinder";

export class CodeGenerator implements ISkill {
  name: string;
  capability: string;

  constructor() {
    this.name = "Code Generator";
    this.capability = "Generate code";
  }

  public canInvoke(spec: Spec): boolean {
    return !!spec && !!spec.userInput && spec.userInput.trim().length > 0;
  }

  public async invoke(
    languageModel: LanguageModelChatMessage,
    response: ChatResponseStream,
    token: CancellationToken,
    spec: Spec
  ): Promise<{ result: ExecutionResultEnum; spec: Spec }> {
    response.progress("Identify code-generation scenarios...");
    if (
      (!spec.appendix.host || spec.appendix.host.length === 0) &&
      spec.appendix.complexity === 0
    ) {
      const t0 = performance.now();
      const scanResult = await this.userAskPreScanningAsync(spec, token);
      const t1 = performance.now();
      const duration = (t1 - t0) / 1000;
      spec.appendix.telemetryData.measurements[MeasurementCodeGenPreScanTimeInTotalSec] = duration;
      if (!scanResult) {
        return { result: ExecutionResultEnum.Failure, spec: spec };
      }
      spec.appendix.host = scanResult.host;
      spec.appendix.isCustomFunction = scanResult.customFunctions;
      spec.appendix.complexity = scanResult.complexity;
      spec.appendix.shouldContinue = scanResult.shouldContinue;
    }

    if (!spec.appendix.shouldContinue) {
      // Reject will make the whole request rejected
      return { result: ExecutionResultEnum.Rejected, spec: spec };
    }

    if (!spec.appendix.codeSample || spec.appendix.codeSample.length === 0) {
      const t0 = performance.now();
      const samples = await SampleProvider.getInstance().getTopKMostRelevantScenarioSampleCodesBM25(
        token,
        spec.appendix.host,
        spec.userInput,
        1
      );
      const t1 = performance.now();
      const duration = (t1 - t0) / 1000;
      spec.appendix.telemetryData.measurements[MeasurementCodeGenGetSampleTimeInTotalSec] =
        duration;
      if (samples.size > 0) {
        console.debug(`Sample code found: ${Array.from(samples.keys())[0]}`);
        spec.appendix.telemetryData.relatedSampleName = Array.from(samples.values()).map(
          (sample) => {
            // remove the '-1' behind the sample name
            const lastIndex = sample.name.lastIndexOf("-");
            return lastIndex !== -1 ? sample.name.substring(0, lastIndex) : sample.name;
          }
        );
        spec.appendix.codeSample = Array.from(samples.values())[0].codeSample;
      }
    }

    if (!spec.appendix.codeTaskBreakdown || !spec.appendix.codeExplanation) {
      const t0 = performance.now();
      const breakdownResult = await this.userAskBreakdownAsync(
        token,
        spec.appendix.complexity,
        spec.appendix.isCustomFunction,
        spec.appendix.host,
        spec.userInput,
        spec.appendix.codeSample,
        spec
      );
      const t1 = performance.now();
      const duration = (t1 - t0) / 1000;
      spec.appendix.telemetryData.measurements[MeasurementCodeGenTaskBreakdownTimeInTotalSec] =
        duration;

      console.debug(`functional spec: ${breakdownResult?.spec || ""}`);
      console.debug(breakdownResult?.funcs.map((task) => `- ${task}`).join("\n"));
      if (!breakdownResult || !breakdownResult.spec || breakdownResult.funcs.length === 0) {
        if (
          !spec.appendix.telemetryData.measurements[
            MeasurementSystemCodegenTaskBreakdownAttemptFailedCount
          ]
        ) {
          spec.appendix.telemetryData.measurements[
            MeasurementSystemCodegenTaskBreakdownAttemptFailedCount
          ] = 0;
        }
        spec.appendix.telemetryData.measurements[
          MeasurementSystemCodegenTaskBreakdownAttemptFailedCount
        ] += 1;
        return { result: ExecutionResultEnum.Failure, spec: spec };
      }
      spec.appendix.codeTaskBreakdown = breakdownResult.funcs;
      spec.appendix.codeExplanation = breakdownResult.spec;
      response.markdown(`
${spec.appendix.codeExplanation
  .substring(spec.appendix.codeExplanation.indexOf("1."))
  .replace(/\b\d+\./g, (match) => `\n${match}`)}
`);
    }
    if (!spec.appendix.telemetryData.measurements[MeasurementCodeGenAttemptCount]) {
      spec.appendix.telemetryData.measurements[MeasurementCodeGenAttemptCount] = 0;
    }
    spec.appendix.telemetryData.measurements[MeasurementCodeGenAttemptCount] += 1;
    let progressMessageStr = localize(
      "teamstoolkit.chatParticipants.officeAddIn.generateCode.hint"
    );
    if (spec.appendix.complexity >= 50) {
      progressMessageStr += localize(
        "teamstoolkit.chatParticipants.officeAddIn.generateCode.complex"
      );
    } else {
      progressMessageStr += localize(
        "teamstoolkit.chatParticipants.officeAddIn.generateCode.simple"
      );
    }
    response.progress(progressMessageStr);
    const t0 = performance.now();
    let codeSnippet: string | null = "";
    codeSnippet = await this.generateCode(
      token,
      spec.appendix.host,
      spec,
      spec.appendix.codeExplanation,
      spec.appendix.isCustomFunction,
      spec.appendix.codeTaskBreakdown,
      spec.appendix.codeSample
    );
    const t1 = performance.now();
    const duration = (t1 - t0) / 1000;
    if (!spec.appendix.telemetryData.measurements[MeasurementCodeGenExecutionTimeInTotalSec]) {
      spec.appendix.telemetryData.measurements[MeasurementCodeGenExecutionTimeInTotalSec] =
        duration;
    } else {
      spec.appendix.telemetryData.measurements[MeasurementCodeGenExecutionTimeInTotalSec] +=
        duration;
    }
    console.log(`Code generation took ${duration} seconds.`);
    if (!codeSnippet) {
      spec.appendix.telemetryData.properties[PropertySystemCodeGenResult] = "false";
      return { result: ExecutionResultEnum.Failure, spec: spec };
    }

    spec.appendix.telemetryData.properties[PropertySystemCodeGenResult] = "true";
    spec.appendix.codeSnippet = codeSnippet;
    return { result: ExecutionResultEnum.Success, spec: spec };
  }

  async userAskPreScanningAsync(
    spec: Spec,
    token: CancellationToken
  ): Promise<null | {
    host: string;
    shouldContinue: boolean;
    customFunctions: boolean;
    complexity: number;
  }> {
    const userPrompt = getUserInputBreakdownTaskUserPrompt(spec.userInput);
    const defaultSystemPrompt = getUserAskPreScanningSystemPrompt();

    // Perform the desired operation
    const messages: LanguageModelChatMessage[] = [
      new LanguageModelChatMessage(LanguageModelChatMessageRole.User, userPrompt),
      new LanguageModelChatMessage(LanguageModelChatMessageRole.User, defaultSystemPrompt),
    ];
    let copilotResponse = await getCopilotResponseAsString(
      "copilot-gpt-3.5-turbo", // "copilot-gpt-4", // "copilot-gpt-3.5-turbo",
      messages,
      token
    );
    spec.appendix.telemetryData.chatMessages.push(...messages);
    spec.appendix.telemetryData.responseChatMessages.push(
      new LanguageModelChatMessage(LanguageModelChatMessageRole.Assistant, copilotResponse)
    );
    let copilotRet: {
      host: string;
      shouldContinue: boolean;
      customFunctions: boolean;
      complexity: number;
    };

    try {
      if (!copilotResponse) {
        return null; // The response is empty
      }
      copilotResponse = copilotResponse.replace(/\\n/g, "").replace(/\n/g, "");
      const codeSnippetRet = copilotResponse.match(/```json([\s\S]*?)```/);
      if (!codeSnippetRet) {
        // try if the LLM already give a json object
        copilotRet = JSON.parse(copilotResponse.trim());
      } else {
        copilotRet = JSON.parse(codeSnippetRet[1].trim());
      }
      console.debug(
        `Custom functions: ${copilotRet.customFunctions ? "true" : "false"}, Complexity score: ${
          copilotRet.complexity
        }`
      );
    } catch (error) {
      console.error("[User task scanning] Failed to parse the response from Copilot:", error);
      return null;
    }

    return copilotRet;
  }

  async userAskBreakdownAsync(
    token: CancellationToken,
    complexity: number,
    isCustomFunctions: boolean,
    host: string,
    userInput: string,
    sampleCode: string,
    spec: Spec
  ): Promise<null | {
    spec: string;
    funcs: string[];
  }> {
    let userPrompt: string = getUserSimpleAskBreakdownTaskSystemPrompt(userInput);
    if (isCustomFunctions) {
      userPrompt = `This is a task about Excel custom functions, pay attention if this is a regular custom functions or streaming custom functions:\n\n ${userPrompt}`;
    }
    userPrompt += "\nDo not generate code snippets.\n\nThink about that step by step.";

    // Perform the desired operation
    const messages: LanguageModelChatMessage[] = [
      new LanguageModelChatMessage(LanguageModelChatMessageRole.User, userPrompt),
    ];

    //     let declarations = new Map<string, SampleData>();
    //     if (!spec.appendix.apiDeclarationsReference || !spec.appendix.apiDeclarationsReference.size) {
    //       declarations = await SampleProvider.getInstance().getMostRelevantDeclarationsUsingLLM(
    //         token,
    //         host,
    //         userInput,
    //         "" //sampleCode
    //       );

    //       spec.appendix.apiDeclarationsReference = declarations;
    //     } else {
    //       declarations = spec.appendix.apiDeclarationsReference;
    //     }

    //     if (declarations.size > 0) {
    //       const groupedMethodsOrProperties: Map<string, SampleData[]> = new Map<string, SampleData[]>();
    //       declarations.forEach((declaration) => {
    //         if (!groupedMethodsOrProperties.has(declaration.definition)) {
    //           groupedMethodsOrProperties.set(declaration.definition, []);
    //         }
    //         groupedMethodsOrProperties.get(declaration.definition)?.push(declaration);
    //       });

    //       let tempClassDeclaration = "\n```typescript\n";
    //       groupedMethodsOrProperties.forEach((methodsOrPropertiesCandidates, className) => {
    //         tempClassDeclaration += `
    // class ${className} extends OfficeExtension.ClientObject {
    //   ${methodsOrPropertiesCandidates.map((sampleData) => sampleData.codeSample).join("\n")}
    // }
    // \n
    //       `;
    //       });
    //       tempClassDeclaration += "```\n";

    //       console.debug(`API declarations: \n${tempClassDeclaration}`);
    //       const classPrompt = `Here are some API declaration that you may want to use as reference, you should only pick those relevant to the user's ask. List the name of used method, property with its class as part of the spec and function descriptions :\n\n${tempClassDeclaration}`;
    //       messages.push(new LanguageModelChatMessage(LanguageModelChatMessageRole.System, classPrompt));
    //     }

    if (sampleCode.length > 0) {
      messages.push(
        new LanguageModelChatMessage(
          LanguageModelChatMessageRole.User,
          getCodeSamplePrompt(sampleCode)
        )
      );
    }

    let copilotResponse = await getCopilotResponseAsString(
      "copilot-gpt-4", //"copilot-gpt-4", // "copilot-gpt-3.5-turbo",
      messages,
      token
    );
    spec.appendix.telemetryData.chatMessages.push(...messages);
    spec.appendix.telemetryData.responseChatMessages.push(
      new LanguageModelChatMessage(LanguageModelChatMessageRole.Assistant, copilotResponse)
    );
    let copilotRet: {
      spec: string;
      funcs: string[];
    };

    try {
      if (!copilotResponse) {
        return null; // The response is empty
      }
      copilotResponse = copilotResponse.replace(/\\n/g, "  ").replace(/\n/g, "  ");
      const codeSnippetRet = copilotResponse.match(/```json([\s\S]*?)```/);
      if (!codeSnippetRet) {
        // try if the LLM already give a json object
        copilotRet = JSON.parse(copilotResponse.trim());
      } else {
        copilotRet = JSON.parse(codeSnippetRet[1].trim());
      }
    } catch (error) {
      console.error("[User task breakdown] Failed to parse the response " + copilotResponse, error);
      return null;
    }
    // We're not able to control the LLM output very precisely, so we need to do some post-processing here
    // For non-custom functions, we need to make sure the entry function 'main' is included in the task breakdown
    // For custom functions, we need to make sure the entry function 'main' is not included in the task breakdown
    if (!isCustomFunctions) {
      copilotRet.funcs.push(
        "Create an entry function named 'main'. This function doesn't take any parameters and will call other functions in the list in right order. The function should be declared as 'async function'."
      );
    }

    return copilotRet;
  }

  async generateCode(
    token: CancellationToken,
    host: string,
    spec: Spec,
    codeSpec: string,
    isCustomFunctions: boolean,
    suggestedFunction: string[],
    sampleCode: string
  ) {
    const userPrompt = getGenerateCodeUserPrompt(codeSpec, host, suggestedFunction);

    //     if (!spec.appendix.apiDeclarationsReference || !spec.appendix.apiDeclarationsReference.size) {
    //       const declarations = await SampleProvider.getInstance().getMostRelevantDeclarationsUsingLLM(
    //         token,
    //         host,
    //         codeSpec,
    //         "" //sampleCode
    //       );

    //       spec.appendix.apiDeclarationsReference = declarations;
    //     }

    //     let declarationPrompt = getGenerateCodeDeclarationPrompt();
    //     if (spec.appendix.apiDeclarationsReference.size > 0) {
    //       const groupedMethodsOrProperties: Map<string, SampleData[]> = new Map<string, SampleData[]>();
    //       spec.appendix.apiDeclarationsReference.forEach((declaration) => {
    //         if (!groupedMethodsOrProperties.has(declaration.definition)) {
    //           groupedMethodsOrProperties.set(declaration.definition, []);
    //         }
    //         groupedMethodsOrProperties.get(declaration.definition)?.push(declaration);
    //       });

    //       let tempClassDeclaration = "\n```typescript\n";
    //       groupedMethodsOrProperties.forEach((methodsOrPropertiesCandidates, className) => {
    //         tempClassDeclaration += `
    // class ${className} extends OfficeExtension.ClientObject {
    //   ${methodsOrPropertiesCandidates.map((sampleData) => sampleData.codeSample).join("\n")}
    // }
    // \n
    //       `;
    //       });
    //       tempClassDeclaration += "```\n";

    //       declarationPrompt += tempClassDeclaration;
    //       // console.debug(`API declarations: \n${declarationPrompt}`);
    //     }
    const model: "copilot-gpt-4" | "copilot-gpt-3.5-turbo" = "copilot-gpt-4";
    let msgCount = 0;

    // Perform the desired operation
    // The order in array is matter, don't change it unless you know what you are doing
    const messages: LanguageModelChatMessage[] = [
      new LanguageModelChatMessage(LanguageModelChatMessageRole.User, userPrompt),
    ];

    let referenceUserPrompt = "";
    switch (host) {
      case "Excel":
        if (!isCustomFunctions) {
          referenceUserPrompt = excelSystemPrompt;
        } else {
          referenceUserPrompt = customFunctionSystemPrompt;
        }
        messages.push(
          new LanguageModelChatMessage(LanguageModelChatMessageRole.User, referenceUserPrompt)
        );
        break;
      default:
        referenceUserPrompt = "";
        break;
    }
    // // May sure for the custom functions, the reference user prompt is shown first so it has lower risk to be cut off
    // if (isCustomFunctions) {
    //   messages.push(
    //     new LanguageModelChatMessage(LanguageModelChatMessageRole.System, referenceUserPrompt)
    //   );
    //   messages.push(
    //     new LanguageModelChatMessage(LanguageModelChatMessageRole.System, declarationPrompt)
    //   );
    // } else {
    //   messages.push(
    //     new LanguageModelChatMessage(LanguageModelChatMessageRole.System, declarationPrompt)
    //   );
    //   messages.push(
    //     new LanguageModelChatMessage(LanguageModelChatMessageRole.System, referenceUserPrompt)
    //   );
    // }
    if (sampleCode.length > 0) {
      let samplePrompt = getGenerateCodeSamplePrompt();
      samplePrompt += `
      \n
      \`\`\`typescript
      ${sampleCode}
      \`\`\`

      Let's think step by step.
      `;
      messages.push(new LanguageModelChatMessage(LanguageModelChatMessageRole.User, samplePrompt));
    }
    // Because of the token window limitation, we have to cut off the messages if it exceeds the limitation
    msgCount = countMessagesTokens(messages);
    while (msgCount > getTokenLimitation(model)) {
      messages.pop();
      msgCount = countMessagesTokens(messages);
    }
    console.debug(`token count: ${msgCount}, number of messages remains: ${messages.length}.`);

    const copilotResponse = await getCopilotResponseAsString(model, messages, token);
    spec.appendix.telemetryData.chatMessages.push(...messages);
    spec.appendix.telemetryData.responseChatMessages.push(
      new LanguageModelChatMessage(LanguageModelChatMessageRole.Assistant, copilotResponse)
    );
    // extract the code snippet and the api list out
    const codeSnippetRet = copilotResponse.match(/```typescript([\s\S]*?)```/);
    if (!codeSnippetRet) {
      // something wrong with the LLM output
      // TODO: Add handling for this case
      console.error(
        "[Code generation] Failed to extract the code snippet from the response:",
        copilotResponse
      );
      return null;
    }

    return correctPropertyLoadSpelling(codeSnippetRet[1].trim());
  }
}
