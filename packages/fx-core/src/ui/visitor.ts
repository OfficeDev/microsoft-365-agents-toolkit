// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FxError,
  IQTreeNode,
  InputResult,
  Inputs,
  MultiSelectQuestion,
  OptionItem,
  Question,
  Result,
  SingleSelectQuestion,
  StaticOptions,
  TelemetryEvent,
  TelemetryProperty,
  TelemetryReporter,
  UserError,
  UserInteraction,
  Void,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import { assign, cloneDeep } from "lodash";
import {
  EmptyOptionError,
  InputValidationError,
  MissingRequiredInputError,
  UserCancelError,
  assembleError,
} from "../error";
import { getValidationFunction, validate, validationUtils } from "./validationUtils";

async function isAutoSkipSelect(q: Question, inputs: Inputs): Promise<boolean> {
  let skipSingle = false;
  if (q.type === "singleSelect" || q.type === "multiSelect") {
    if (q.skipSingleOption !== undefined) {
      if (typeof q.skipSingleOption === "function") {
        skipSingle = await q.skipSingleOption(inputs);
      } else {
        skipSingle = q.skipSingleOption;
      }
    }
  }
  return skipSingle;
}

export function getSingleOption(
  q: SingleSelectQuestion | MultiSelectQuestion,
  option?: StaticOptions
): string | OptionItem | [string | OptionItem] {
  if (!option) option = q.staticOptions;
  const optionIsString = typeof option[0] === "string";
  let returnResult: string | OptionItem;
  if (optionIsString) returnResult = option[0] as string;
  else {
    if (q.returnObject === true) returnResult = option[0];
    else returnResult = (option[0] as OptionItem).id;
  }
  if (q.type === "singleSelect") return returnResult;
  else return [returnResult];
}

async function getCallFuncValue(inputs: Inputs, raw?: unknown): Promise<unknown> {
  if (raw && typeof raw === "function") {
    return await raw(inputs);
  }
  return raw;
}

export async function loadOptions(
  question: SingleSelectQuestion | MultiSelectQuestion,
  inputs: Inputs
): Promise<StaticOptions> {
  let options = question.staticOptions;
  if (question.dynamicOptions) {
    options = await question.dynamicOptions(inputs);
  }
  return options;
}

export type QuestionTreeVisitor = (
  question: Question,
  ui: UserInteraction,
  inputs: Inputs,
  step?: number,
  totalSteps?: number
) => Promise<Result<InputResult<any>, FxError>>;

export async function singleSelectCallback(
  question: SingleSelectQuestion,
  answer: string | OptionItem,
  inputs: Inputs,
  options: StaticOptions
): Promise<void> {
  if (!question.onDidSelection) return;
  if (typeof answer !== "string") {
    await question.onDidSelection(answer, inputs);
  } else {
    const selected = (options as (string | OptionItem)[]).find((o: string | OptionItem) => {
      if (typeof o === "string") return o === answer;
      return o.id === answer;
    });
    if (selected) {
      await question.onDidSelection(selected, inputs);
    }
  }
}

/**
 * ask question when visiting the question tree
 * @param question
 * @param core
 * @param inputs
 */
export const questionVisitor: QuestionTreeVisitor = async function (
  question: Question,
  ui: UserInteraction,
  inputs: Inputs,
  step?: number,
  totalSteps?: number
): Promise<Result<InputResult<any>, FxError>> {
  if (inputs[question.name] !== undefined) {
    //if answer is preset, validate it and quick return the preset answer
    const res = await validationUtils.validateInputs(question, inputs[question.name], inputs);
    if (res) return err(new InputValidationError(question.name, res, "questionVisitor"));
    if (question.type === "singleSelect" && question.onDidSelection) {
      const options = await loadOptions(question, inputs);
      await singleSelectCallback(question, inputs[question.name], inputs, options);
    }
    return ok({ type: "skip", result: inputs[question.name] });
  }

  if (inputs.nonInteractive) {
    // if no preset answer and non-interactive mode
    // first priority: use single option as value
    if (question.type === "singleSelect" || question.type === "multiSelect") {
      const skipSingle = await isAutoSkipSelect(question, inputs);
      if (skipSingle) {
        const options = await loadOptions(question, inputs);
        if (options.length === 0) {
          return err(new EmptyOptionError(question.name, "questionVisitor"));
        }
        if (options.length === 1) {
          const value = getSingleOption(question, options);
          if (value) {
            if (question.type === "singleSelect") {
              await singleSelectCallback(question, inputs[question.name], inputs, options);
            }
            return ok({ type: "skip", result: value });
          }
        }
      }
    }
    // second priority: use default as value
    if (question.default) {
      const value = (await getCallFuncValue(inputs, question.default)) as
        | string
        | string[]
        | OptionItem
        | OptionItem[];
      if (value) {
        const validateRes = await validationUtils.validateInputs(question, value, inputs);
        if (validateRes) {
          return err(new InputValidationError(question.name, validateRes, "questionVisitor"));
        } else {
          if (question.type === "singleSelect") {
            const options = await loadOptions(question, inputs);
            await singleSelectCallback(question, value as string, inputs, options);
          }
          return ok({ type: "skip", result: value });
        }
      }
    }
    if (question.required)
      return err(new MissingRequiredInputError(question.name, "questionVisitor"));
    else return ok({ type: "skip", result: undefined });
  }

  //no preset answer and interactive mode, call UI
  const title = (await getCallFuncValue(inputs, question.title)) as string;
  let defaultValue:
    | string
    | string[]
    | (() => Promise<string>)
    | (() => Promise<string[]>)
    | boolean
    | undefined = undefined;
  if (question.forgetLastValue !== true && question.value)
    defaultValue = question.value as string | string[];
  else {
    if (question.default) {
      if (typeof question.default === "function") {
        defaultValue = async () => {
          return await (question as any).default(inputs);
        };
      } else {
        defaultValue = question.default;
      }
    }
  }
  const placeholder = (await getCallFuncValue(inputs, question.placeholder)) as string;
  const prompt = (await getCallFuncValue(inputs, question.prompt)) as string;
  if (question.type === "text") {
    const validationFunc = question.validation
      ? getValidationFunction<string>(question.validation, inputs)
      : undefined;
    const additionalValidationOnAcceptFunc = question.additionalValidationOnAccept
      ? getValidationFunction<string>(question.additionalValidationOnAccept, inputs)
      : undefined;

    return await ui.inputText({
      name: question.name,
      title: title,
      password: question.password,
      default: defaultValue as string | (() => Promise<string>),
      placeholder: placeholder,
      prompt: prompt,
      validation: validationFunc,
      step: step,
      totalSteps: totalSteps,
      additionalValidationOnAccept: additionalValidationOnAcceptFunc,
    });
  } else if (question.type === "singleSelect" || question.type === "multiSelect") {
    const skipSingle = await isAutoSkipSelect(question, inputs);
    let options: StaticOptions | (() => Promise<StaticOptions>) | undefined = undefined;
    if (question.dynamicOptions) {
      options = async () => {
        return question.dynamicOptions!(inputs);
      };
    } else {
      if (!question.staticOptions || question.staticOptions.length === 0) {
        return err(new EmptyOptionError(question.name, "questionVisitor"));
      }
      if (skipSingle && question.staticOptions.length === 1) {
        // quick return for static options with only one item
        const returnResult = getSingleOption(question, question.staticOptions);
        if (question.type === "singleSelect" && question.onDidSelection) {
          let selected = returnResult as string | OptionItem;
          if (typeof selected === "string") {
            selected = question.staticOptions[0];
          }
          await question.onDidSelection(selected, inputs);
        }
        return ok({ type: "skip", result: returnResult });
      }
      options = question.staticOptions;
    }
    if (question.type === "singleSelect") {
      const validationFunc = question.validation
        ? getValidationFunction<string>(question.validation, inputs)
        : undefined;
      const res = await ui.selectOption({
        name: question.name,
        title: title,
        options: options,
        returnObject: question.returnObject,
        default: defaultValue as string | (() => Promise<string>),
        placeholder: placeholder,
        prompt: prompt,
        step: step,
        totalSteps: totalSteps,
        buttons: question.buttons,
        validation: validationFunc,
        skipSingleOption: skipSingle,
      });
      if (res.isOk() && res.value.type === "success") {
        let selected = res.value.result;
        if (typeof selected === "string") {
          const options = res.value.options as string[] | OptionItem[];
          if (options && options.length > 0 && typeof options[0] !== "string") {
            selected = (options as OptionItem[]).find((o: OptionItem) => o.id === selected);
          }
        }
        await question.onDidSelection?.(selected!, inputs);
      }
      return res;
    } else {
      const validationFunc = question.validation
        ? getValidationFunction<string[]>(question.validation, inputs)
        : undefined;
      return await ui.selectOptions({
        name: question.name,
        title: title,
        options: options,
        returnObject: question.returnObject,
        default: defaultValue as string[] | (() => Promise<string[]>),
        placeholder: placeholder,
        prompt: prompt,
        onDidChangeSelection: question.onDidChangeSelection,
        step: step,
        totalSteps: totalSteps,
        validation: validationFunc,
        skipSingleOption: skipSingle,
      });
    }
  } else if (question.type === "multiFile") {
    const validationFunc = question.validation
      ? getValidationFunction<string[]>(question.validation, inputs)
      : undefined;
    return await ui.selectFiles({
      name: question.name,
      title: title,
      placeholder: placeholder,
      prompt: prompt,
      default: defaultValue as string[] | (() => Promise<string[]>),
      step: step,
      totalSteps: totalSteps,
      validation: validationFunc,
    });
  } else if (question.type === "singleFile") {
    const validationFunc = question.validation
      ? getValidationFunction<string>(question.validation, inputs)
      : undefined;
    let defaultFolder;
    if (question.defaultFolder) {
      if (typeof question.defaultFolder === "function") {
        defaultFolder = async () => {
          return await (question as any).defaultFolder(inputs);
        };
      } else {
        defaultFolder = question.defaultFolder;
      }
    }
    return await ui.selectFile({
      name: question.name,
      title: title,
      placeholder: placeholder,
      prompt: prompt,
      default: defaultValue as string | (() => Promise<string>),
      step: step,
      totalSteps: totalSteps,
      validation: validationFunc,
      filters: question.filters,
      innerStep: question.innerStep,
      innerTotalStep: question.innerTotalStep,
      defaultFolder,
    });
  } else if (question.type === "folder") {
    const validationFunc = question.validation
      ? getValidationFunction<string>(question.validation, inputs)
      : undefined;
    return await ui.selectFolder({
      name: question.name,
      title: title,
      placeholder: placeholder,
      prompt: prompt,
      default: defaultValue as string | (() => Promise<string>),
      step: step,
      totalSteps: totalSteps,
      validation: validationFunc,
    });
  } else if (question.type === "singleFileOrText" && !!ui.selectFileOrInput) {
    const validationFunc = question.validation
      ? getValidationFunction<string>(question.validation, inputs)
      : undefined;
    const inputValidationFunc = question.inputBoxConfig.validation
      ? getValidationFunction<string>(question.inputBoxConfig.validation, inputs)
      : undefined;
    const innerTitle = (await getCallFuncValue(inputs, question.inputBoxConfig.title)) as string;
    const innerPlaceholder = (await getCallFuncValue(
      inputs,
      question.inputBoxConfig.placeholder
    )) as string;
    const innerPrompt = (await getCallFuncValue(inputs, question.inputBoxConfig.prompt)) as string;
    const res = await ui.selectFileOrInput({
      name: question.name,
      title: title,
      placeholder: placeholder,
      prompt: prompt,
      inputOptionItem: question.inputOptionItem,
      inputBoxConfig: {
        name: question.inputBoxConfig.name,
        title: innerTitle,
        placeholder: innerPlaceholder,
        prompt: innerPrompt,
        validation: inputValidationFunc,
        step: question.inputBoxConfig.step,
      },
      filters: question.filters,
      step: step,
      totalSteps: totalSteps,
      validation: validationFunc,
    });
    return res;
  } else if (question.type === "confirm" && ui.confirm) {
    const res = await ui.confirm({
      name: question.name,
      title: title,
      default: defaultValue as boolean,
      step: step,
      totalSteps: totalSteps,
    });
    return res;
  }
  return err(
    new UserError(
      "API",
      "UnsupportedNodeType",
      `Unsupported question node type:${JSON.stringify(question)}`,
      `Unsupported question node type:${JSON.stringify(question)}`
    )
  );
};

/**
 * serialize the tree node into array in DFS order
 */
export function collect(
  node: IQTreeNode,
  list: IQTreeNode[],
  parentMap: Map<IQTreeNode, IQTreeNode>
): void {
  list.push(node);
  if (node.children) {
    for (const child of node.children) {
      if (child) {
        parentMap.set(child, node);
        collect(child, list, parentMap);
      }
    }
  }
}

export async function traverse(
  root: IQTreeNode,
  inputs: Inputs,
  ui: UserInteraction,
  telemetryReporter?: TelemetryReporter,
  visitor: QuestionTreeVisitor = questionVisitor
): Promise<Result<Void, FxError>> {
  // The reason to clone is that we don't want to change the original inputs if user cancel the process
  const clonedInputs = cloneDeep(inputs);

  // 1. collect all nodes into array
  const parentMap = new Map<IQTreeNode, IQTreeNode>();
  const nodeList: IQTreeNode[] = [];
  collect(root, nodeList, parentMap);

  const visitedNodeSet = new Set<IQTreeNode>();

  const visitedInputNodeArray: IQTreeNode[] = [];

  let i = 0;
  for (; i < nodeList.length; ++i) {
    const node = nodeList[i];

    // if parent node is not visited, current node should not be visited
    const parent = parentMap.get(node);
    if (parent) {
      if (!visitedNodeSet.has(parent)) {
        continue;
      }
    }

    // 1. check condition
    if (node.condition) {
      let parentValue: any = undefined;
      // const parent = parentMap.get(node);
      if (parent) {
        parentValue = findValue(parent, parentMap);
      }
      const validRes = await validate(
        node.condition,
        parentValue as string | string[] | OptionItem | OptionItem[],
        clonedInputs
      );
      if (validRes !== undefined) {
        continue;
      }
    }

    // 2. visit node if not group
    if (node.data.type !== "group") {
      const question = node.data;
      let res;
      try {
        res = await visitor(
          question,
          ui,
          clonedInputs,
          visitedInputNodeArray.length + 1,
          undefined
        );
        sendTelemetryEvent(telemetryReporter, res, question, clonedInputs);
      } catch (e) {
        return err(assembleError(e));
      }
      if (res.isErr()) {
        // Cancel or Error
        return err(res.error);
      }
      const inputResult = res.value;
      if (inputResult.type === "back") {
        const prevNode = visitedInputNodeArray.pop();
        if (!prevNode) {
          return err(new UserCancelError());
        }
        for (--i; i >= 0; --i) {
          const tmpNode = nodeList[i];
          visitedNodeSet.delete(tmpNode);
          // clear prevNode data
          if (tmpNode.data.type !== "group") {
            delete tmpNode.data.value;
            delete tmpNode.data.valueType;
            delete clonedInputs[tmpNode.data.name];
          }
          if (tmpNode === prevNode) {
            break;
          }
        }
        --i;
        continue;
      } else {
        //success or skip: set value
        question.value = inputResult.result;
        question.valueType = inputResult.type;
        clonedInputs[question.name] = question.value;
        visitedNodeSet.add(node);
        if (question.valueType === "success") {
          visitedInputNodeArray.push(node);
        }
      }
    } else {
      visitedNodeSet.add(node);
    }
  }
  assign(inputs, clonedInputs);
  return ok(Void);
}

export function findValue(curr: IQTreeNode, parentMap: Map<IQTreeNode, IQTreeNode>): any {
  if (curr.data.type !== "group") {
    // need to convert OptionItem value into id for validation
    if (curr.data.type === "singleSelect") {
      const sq: SingleSelectQuestion = curr.data;
      if (sq.value && typeof sq.value !== "string" && sq.value.id) {
        return sq.value.id;
      }
    } else if (curr.data.type === "multiSelect") {
      const mq: MultiSelectQuestion = curr.data;
      if (mq.value && typeof mq.value[0] !== "string") {
        return (mq.value as OptionItem[]).map((i) => i.id);
      }
    }
    return curr.data.value;
  }
  const parent = parentMap.get(curr);
  if (parent) {
    return findValue(parent, parentMap);
  }
  return undefined;
}

function sendTelemetryEvent(
  telemetryReporter: TelemetryReporter | undefined,
  qvres: Result<InputResult<any>, FxError>,
  question: Question,
  inputs: Inputs
) {
  if (qvres.isErr()) {
    telemetryReporter?.sendTelemetryEvent(TelemetryEvent.askQuestion, {
      [TelemetryProperty.answerType]: qvres.error.name,
      [TelemetryProperty.question]: question.name,
      [TelemetryProperty.platform]: inputs.platform,
      [TelemetryProperty.stage]: inputs.stage ? inputs.stage : "",
    });
  } else {
    telemetryReporter?.sendTelemetryEvent(TelemetryEvent.askQuestion, {
      [TelemetryProperty.answerType]: qvres.value.type,
      [TelemetryProperty.question]: question.name,
      [TelemetryProperty.answer]:
        question.type == "singleSelect" || question.type == "multiSelect"
          ? qvres.value.result?.toString()
          : "",
      [TelemetryProperty.platform]: inputs.platform,
      [TelemetryProperty.stage]: inputs.stage ? inputs.stage : "",
    });
  }
}
