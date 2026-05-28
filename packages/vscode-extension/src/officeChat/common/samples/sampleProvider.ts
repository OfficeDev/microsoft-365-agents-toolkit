// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CancellationToken, LanguageModelChatMessage, LanguageModelChatMessageRole } from "vscode";
import { BM25, BMDocument } from "../../retrievalUtil/BM25";
import { OfficeTemplateModelPorvider, WXPAppName } from "./officeTemplateModelPorvider";
import { SampleData } from "./sampleData";
import { prepareDiscription } from "../../retrievalUtil/retrievalUtil";
import { countMessagesTokens, getCopilotResponseAsString } from "../../../chat/utils";
import {
  getMostRelevantClassPrompt,
  getMostRelevantClassUsingNameOnlyPrompt,
  getMostRelevantMethodPropertyPrompt,
} from "../../officePrompts";
import { DeclarationFinder } from "../declarationFinder";
import { getTokenLimitation } from "../../consts";
import { Spec } from "../skills/spec";

// TODO: adjust the score threshold
const scoreThreshold = 0.5;

export class SampleProvider {
  private static instance: SampleProvider;

  private constructor() {
    // Private constructor to prevent direct instantiation
  }

  public static getInstance(): SampleProvider {
    if (!SampleProvider.instance) {
      SampleProvider.instance = new SampleProvider();
    }
    return SampleProvider.instance;
  }

  public async getTopKMostRelevantScenarioSampleCodesBM25(
    token: CancellationToken,
    host: string,
    scenario: string,
    k: number
  ): Promise<Map<string, SampleData>> {
    const samples: Map<string, SampleData> = new Map<string, SampleData>();
    const bm25: BM25 | null = await OfficeTemplateModelPorvider.getInstance().getBM25Model(
      host as WXPAppName
    );
    if (bm25) {
      const query = prepareDiscription(scenario.toLowerCase());
      const documents: BMDocument[] = bm25.search(query, k);

      for (const doc of documents) {
        if (doc.score >= scoreThreshold && doc.document.metadata) {
          const sampleData = doc.document.metadata as SampleData;
          samples.set(sampleData.name, sampleData);
        }
      }
    }
    return new Promise<Map<string, SampleData>>((resolve, reject) => {
      resolve(samples);
    });
  }

  /**
   *  Get the most relevant declarations using the language model.
   *  Due to the limitation of the token count, we have to split the process into a few steps.
   *  The first step is to get the most relevant classes based on the understanding of the code spec and the sample.
   *  The second step is to get the most relevant methods or properties from selected classes from previous step, based on the understanding of the code spec and the sample.
   */
  public async getMostRelevantDeclarationsUsingLLM(
    token: CancellationToken,
    host: string,
    codeSpec: string,
    sample: string,
    spec: Spec
  ): Promise<Map<string, SampleData>> {
    const pickedDeclarations: Map<string, SampleData> = new Map<string, SampleData>();
    const model: "copilot-gpt-3.5-turbo" | "copilot-gpt-4" = "copilot-gpt-4";
    const t1 = performance.now();
    let countOfLLMInvoke = 0;
    const classSummaries = await DeclarationFinder.getInstance().getClassSummariesForHost(host);
    if (classSummaries.length === 0) {
      return pickedDeclarations;
    }

    // It is possible that with the increase of the number of classes, the token count of the message will exceed the limitation. So if the token count exceeds the limitation, we will use the prompt that only contains the class name rather than the class's description to reduce the token count.
    let sampleMessage: LanguageModelChatMessage = new LanguageModelChatMessage(
      LanguageModelChatMessageRole.User,
      getMostRelevantClassPrompt(codeSpec, classSummaries, sample)
    );
    let msgCount = countMessagesTokens([sampleMessage]);
    if (msgCount > getTokenLimitation(model)) {
      sampleMessage = new LanguageModelChatMessage(
        LanguageModelChatMessageRole.User,
        getMostRelevantClassUsingNameOnlyPrompt(codeSpec, classSummaries, sample)
      );
      msgCount = countMessagesTokens([sampleMessage]);
    }

    if (msgCount > getTokenLimitation(model)) {
      console.debug(
        "[getMostRelevantDeclarationsUsingLLM] The token count of the message exceeds the limitation."
      );
      return pickedDeclarations;
    }

    countOfLLMInvoke += 1;
    const copilotResponse = await getCopilotResponseAsString(model, [sampleMessage], token);
    spec.appendix.telemetryData.chatMessages.push(sampleMessage);
    spec.appendix.telemetryData.responseChatMessages.push(
      new LanguageModelChatMessage(LanguageModelChatMessageRole.Assistant, copilotResponse)
    );
    const returnObject: { picked: string[] } = JSON.parse(
      copilotResponse.replace("```json", "").replace("```", "").replace(/\\n/g, "")
    );
    const classNames: string[] = returnObject.picked.map((value) => value.replace("- ", "").trim());

    if (classNames.length === 0) {
      console.debug("[getMostRelevantDeclarationsUsingLLM] No relevant class found for this task.");
      return pickedDeclarations;
    } else {
      console.debug("[getMostRelevantDeclarationsUsingLLM] The relevant classes are: ", classNames);
    }

    const t2 = performance.now();
    const classDeclarationPairs: [string, SampleData[]][] = [];
    for (const className of classNames) {
      const methodsOrProperties =
        await DeclarationFinder.getInstance().getMethodsOrPropertiesForClass(host, className);
      classDeclarationPairs.push([className, methodsOrProperties]);
    }

    const giantMethodsOrPropertiesSet: Map<string, SampleData[]>[] = [];
    // It is possible that the token count of the message will exceed the limitatiotn. So we have to split the process into a few steps. In some cases, a single class may has huge amount of methods or properties we can't afford, we have to skip it. For example, the class "Worksheet" in Excel has 100+ methods and properties.
    while (classDeclarationPairs.length > 0) {
      let msgCount = 0;
      // The following two variables are used to store the classes and methods/properties that will contains in the message send to copilot later, the token count of the message will be safe.
      let classNamesList: string[] = [];
      let methodsOrProperties: SampleData[] = [];
      // following two variables are temporary used to store the classes and methods/properties to calculate the token count. The token count of the message could exceed the limitation.
      const classNamesListTemp: string[] = [];
      const methodsOrPropertiesTemp: SampleData[] = [];

      let groupedMethodsOrProperties: Map<string, SampleData[]> = new Map<string, SampleData[]>();
      let getMoreRelevantMethodsOrPropertiesPrompt = "";

      // The while loop is used to get the classes and methods/properties that will contains in the message send to copilot later, the token count of the message will be safe. Those used classes will be removed from the classDeclarationPairs.
      let candidate: [string, SampleData[]] | undefined;
      while (msgCount < getTokenLimitation(model)) {
        classNamesList = classNamesListTemp.map((value) => value);
        methodsOrProperties = methodsOrPropertiesTemp.map((value) => value);

        candidate = classDeclarationPairs.pop();
        if (!candidate) {
          break;
        }

        classNamesListTemp.unshift(candidate[0]);
        methodsOrPropertiesTemp.unshift(...candidate[1]);
        // group the methods or properties by class name
        groupedMethodsOrProperties = new Map<string, SampleData[]>();
        for (const methodOrProperty of methodsOrPropertiesTemp) {
          if (!groupedMethodsOrProperties.has(methodOrProperty.definition)) {
            groupedMethodsOrProperties.set(methodOrProperty.definition, []);
          }
          groupedMethodsOrProperties.get(methodOrProperty.definition)?.push(methodOrProperty);
        }
        getMoreRelevantMethodsOrPropertiesPrompt = getMostRelevantMethodPropertyPrompt(
          codeSpec,
          classNamesListTemp,
          groupedMethodsOrProperties,
          sample
        );
        sampleMessage = new LanguageModelChatMessage(
          LanguageModelChatMessageRole.User,
          getMoreRelevantMethodsOrPropertiesPrompt
        );
        msgCount = countMessagesTokens([sampleMessage]);
      }
      if (msgCount > getTokenLimitation(model)) {
        if (methodsOrProperties.length === 0) {
          giantMethodsOrPropertiesSet.push(groupedMethodsOrProperties);
          continue;
        } else {
          classDeclarationPairs.push(candidate as [string, SampleData[]]);
        }
      }
      // group the methods or properties by class name
      groupedMethodsOrProperties = new Map<string, SampleData[]>();
      for (const methodOrProperty of methodsOrProperties) {
        if (!groupedMethodsOrProperties.has(methodOrProperty.definition)) {
          groupedMethodsOrProperties.set(methodOrProperty.definition, []);
        }
        groupedMethodsOrProperties.get(methodOrProperty.definition)?.push(methodOrProperty);
      }

      countOfLLMInvoke += 1;
      const picked = await this.getMostRelevantPropertiesOrMethodsDeclaratitons(
        codeSpec,
        classNamesList,
        groupedMethodsOrProperties,
        sample,
        methodsOrProperties,
        token,
        model,
        spec
      );
      picked.forEach((value, key) => {
        if (!pickedDeclarations.has(key)) {
          pickedDeclarations.set(key, value);
        }
      });
    }

    for (const groupedMethodsOrProperties of giantMethodsOrPropertiesSet) {
      for (const key of Array.from(groupedMethodsOrProperties.keys())) {
        const classNamesListTemp = [key];
        let methodOrPropertyDeclarationsTemp: SampleData[] = [];
        let methodOrPropertyDeclarations: SampleData[] = [];

        while ((groupedMethodsOrProperties.get(key) || []).length > 0) {
          methodOrPropertyDeclarationsTemp = [];
          do {
            methodOrPropertyDeclarations =
              groupedMethodsOrProperties.get(key) || ([] as SampleData[]);
            if (methodOrPropertyDeclarations.length > 1) {
              methodOrPropertyDeclarationsTemp.push(
                methodOrPropertyDeclarations.pop() as SampleData
              );
            }

            const getMoreRelevantMethodsOrPropertiesPrompt = getMostRelevantMethodPropertyPrompt(
              codeSpec,
              classNamesListTemp,
              groupedMethodsOrProperties,
              sample
            );
            sampleMessage = new LanguageModelChatMessage(
              LanguageModelChatMessageRole.User,
              getMoreRelevantMethodsOrPropertiesPrompt
            );
            msgCount = countMessagesTokens([sampleMessage]);
          } while (msgCount > getTokenLimitation(model));

          countOfLLMInvoke += 1;
          const picked = await this.getMostRelevantPropertiesOrMethodsDeclaratitons(
            codeSpec,
            classNamesListTemp,
            groupedMethodsOrProperties,
            sample,
            methodOrPropertyDeclarations,
            token,
            model,
            spec
          );
          picked.forEach((value, key) => {
            if (!pickedDeclarations.has(key)) {
              pickedDeclarations.set(key, value);
            }
          });

          groupedMethodsOrProperties.delete(key);
          groupedMethodsOrProperties.set(key, methodOrPropertyDeclarationsTemp);
        }
      }
    }

    const t3 = performance.now();
    console.log(
      `Pick relevant classes: ${(t2 - t1) / 1000} seconds, get ${
        pickedDeclarations.size
      } methods/properties: ${(t3 - t2) / 1000}, count of LLM invoking: ${countOfLLMInvoke}.`
    );
    return new Promise<Map<string, SampleData>>((resolve, reject) => {
      resolve(pickedDeclarations);
    });
  }

  private async getMostRelevantPropertiesOrMethodsDeclaratitons(
    codeSpec: string,
    classNamesList: string[],
    groupedMethodsOrProperties: Map<string, SampleData[]>,
    sample: string,
    methodsOrProperties: SampleData[],
    token: CancellationToken,
    model: "copilot-gpt-3.5-turbo" | "copilot-gpt-4",
    spec: Spec
  ): Promise<Map<string, SampleData>> {
    const pickedDeclarations: Map<string, SampleData> = new Map<string, SampleData>();
    const getMoreRelevantMethodsOrPropertiesPrompt = getMostRelevantMethodPropertyPrompt(
      codeSpec,
      classNamesList,
      groupedMethodsOrProperties,
      sample
    );
    const sampleMessage = new LanguageModelChatMessage(
      LanguageModelChatMessageRole.User,
      getMoreRelevantMethodsOrPropertiesPrompt
    );
    const copilotResponse = await getCopilotResponseAsString(model, [sampleMessage], token);
    spec.appendix.telemetryData.chatMessages.push(sampleMessage);
    spec.appendix.telemetryData.responseChatMessages.push(
      new LanguageModelChatMessage(LanguageModelChatMessageRole.Assistant, copilotResponse)
    );
    let returnObject: { picked: string[] } = { picked: [] };
    try {
      returnObject = JSON.parse(
        copilotResponse.replace("```json", "").replace("```", "").replace(/\\n/g, "")
      );
    } catch (error) {
      console.log(copilotResponse);
    }

    returnObject.picked.forEach((value: string) => {
      // The return may contains encoded characters, we need to decode them.
      const parts = value
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .split(";")
        .map((part) => part.trim())
        .filter((part) => part.length > 0);
      if (parts.length == 1) {
        // Sometimes the return in the format of "method1;" without class name
        const methodPropertyDeclaration = parts[0].trim() + ";";
        // methodPropertyDeclaration = methodPropertyDeclaration.endsWith(";")
        //   ? methodPropertyDeclaration
        //   : methodPropertyDeclaration + ";";
        const sampleData = methodsOrProperties.find(
          (sample) => sample.codeSample.trim() === methodPropertyDeclaration
        );
        if (sampleData) {
          pickedDeclarations.set(sampleData.description, sampleData);
        } else {
          console.debug("[parts.length == 1]: " + methodPropertyDeclaration);
        }
      } else if (parts.length > 2) {
        // Sometimes the return in the format of "class: className; method1; method2; ...; methodN;"
        const className = parts[0].replace("class:", "").trim();
        for (let i = 1; i < parts.length - 1; i++) {
          const methodPropertyDeclaration = parts[i].trim() + ";";
          // methodPropertyDeclaration = methodPropertyDeclaration.endsWith(";")
          //   ? methodPropertyDeclaration
          //   : methodPropertyDeclaration + ";";
          const sampleData = methodsOrProperties.find(
            (sample) =>
              sample.definition.trim() === className &&
              sample.codeSample.trim() === methodPropertyDeclaration
          );
          if (sampleData) {
            pickedDeclarations.set(sampleData.description, sampleData);
          }
        }
      } else if (parts.length === 2) {
        // in the format of "class: className; methodOrPropertyDeclaration;"
        const className = parts[0].replace("class:", "").trim();
        const methodPropertyDeclaration = parts[1].trim() + ";";
        // methodPropertyDeclaration = methodPropertyDeclaration.endsWith(";")
        //   ? methodPropertyDeclaration
        //   : methodPropertyDeclaration + ";";
        const sampleData = methodsOrProperties.find(
          (sample) =>
            sample.definition.trim() === className &&
            sample.codeSample.trim() === methodPropertyDeclaration
        );
        if (sampleData) {
          pickedDeclarations.set(sampleData.description, sampleData);
        }
      }
    });

    return pickedDeclarations;
  }
}
