// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ok, err, Result } from "neverthrow";
import type { IQTreeNode, Question, Group } from "../api/qm/question";
import type { UserInteraction, InputResult } from "../api/qm/ui";
import type { Inputs, OptionItem } from "../api/types";
import type { FxError } from "../api/error";
import type { ConditionFunc, StringValidation, StringArrayValidation } from "../api/qm/validation";
import type { AtkError } from "../core/error";
import { userError } from "../core/error";

/**
 * Traverse a question tree, collecting user answers into the inputs record.
 *
 * Algorithm: iterative DFS with back-stack for undo.
 * - Group nodes pass through silently (no UI prompt).
 * - Question nodes call the appropriate UI method.
 * - Conditions are evaluated against the current inputs.
 * - "back" responses restore the previous state.
 * - Pre-filled inputs are accepted without prompting.
 *
 * @param root - The question tree root node.
 * @param ui - The platform's UserInteraction implementation.
 * @param inputs - Mutable inputs record; answers are written here.
 * @returns ok(undefined) on completion, err on cancel/error.
 */
export async function traverseQuestionTree(
  root: IQTreeNode,
  ui: UserInteraction,
  inputs: Inputs
): Promise<Result<undefined, AtkError>> {
  // Clone inputs so we can restore on back
  const currentInputs: Record<string, unknown> = { ...inputs };

  // Flatten tree into a visit order with parent tracking
  const visitOrder = flattenTree(root);

  let idx = 0;
  const history: { idx: number; snapshot: Record<string, unknown> }[] = [];

  while (idx < visitOrder.length) {
    const { node, parentQuestionName, descendantCount } = visitOrder[idx];

    // Evaluate condition
    if (node.condition) {
      const conditionPassed = await evaluateCondition(
        node.condition,
        parentQuestionName,
        currentInputs as Inputs
      );
      if (!conditionPassed) {
        // Skip this node and all its descendants
        idx += 1 + descendantCount;
        continue;
      }
    }

    // Skip cliOptionDisabled or inputsDisabled nodes
    if (node.inputsDisabled === "self" || node.inputsDisabled === "all") {
      const skip = node.inputsDisabled === "all" ? 1 + descendantCount : 1;
      idx += skip;
      continue;
    }

    const data = node.data;

    // Group nodes pass through
    if (isGroup(data)) {
      idx++;
      continue;
    }

    const question = data;

    // If already answered (pre-filled), skip prompting
    if (currentInputs[question.name] !== undefined) {
      idx++;
      continue;
    }

    // Save snapshot for back navigation
    history.push({ idx, snapshot: { ...currentInputs } });

    // Ask the question
    const result = await askQuestion(question, ui, currentInputs as Inputs);
    if (result.isErr()) {
      return err(
        userError("QuestionTraversalError", result.error.message ?? "User interaction failed.", {
          source: "questions/traverse",
        })
      );
    }

    const inputResult = result.value;

    if (inputResult.type === "back") {
      // Restore previous state
      history.pop(); // remove the snapshot we just pushed
      const prev = history.pop();
      if (!prev) {
        return err(
          userError("UserCancelled", "User cancelled the operation.", {
            source: "questions/traverse",
          })
        );
      }
      Object.keys(currentInputs).forEach((k) => delete currentInputs[k]);
      Object.assign(currentInputs, prev.snapshot);
      idx = prev.idx;
      continue;
    }

    if (inputResult.type === "success" && inputResult.result !== undefined) {
      // Store the answer (extract id from OptionItem if needed)
      const value = inputResult.result;
      currentInputs[question.name] =
        typeof value === "object" && value !== null && "id" in value
          ? (value as OptionItem).id
          : value;
    }

    idx++;
  }

  // Write collected answers back to inputs
  for (const [key, value] of Object.entries(currentInputs)) {
    (inputs as Record<string, unknown>)[key] = value;
  }

  return ok(undefined);
}

/**
 * Flatten a tree into a linear visit order with parent tracking.
 * Each entry also tracks how many descendants follow it, so the
 * traversal can skip an entire subtree when a condition fails.
 */
interface FlatNode {
  node: IQTreeNode;
  parentQuestionName?: string;
  /** Number of descendants that follow this node in the flat list */
  descendantCount: number;
}

function flattenTree(root: IQTreeNode): FlatNode[] {
  const result: FlatNode[] = [];

  function walk(node: IQTreeNode, parentName?: string): number {
    const currentName = isGroup(node.data) ? parentName : node.data.name;
    const entry: FlatNode = { node, parentQuestionName: parentName, descendantCount: 0 };
    const _myIndex = result.length;
    result.push(entry);

    let totalDescendants = 0;
    if (node.children) {
      for (const child of node.children) {
        totalDescendants += walk(child, currentName);
      }
    }
    entry.descendantCount = totalDescendants;
    return totalDescendants + 1; // include self
  }

  walk(root);
  return result;
}

/**
 * Check if a node data is a Group.
 */
function isGroup(data: Question | Group): data is Group {
  return (data as Group).type === "group";
}

/**
 * Evaluate a condition against current inputs.
 */
async function evaluateCondition(
  condition: StringValidation | StringArrayValidation | ConditionFunc,
  parentQuestionName: string | undefined,
  inputs: Inputs
): Promise<boolean> {
  // ConditionFunc
  if (typeof condition === "function") {
    return condition(inputs);
  }

  // StringValidation / StringArrayValidation with `equals`
  if ("equals" in condition && condition.equals !== undefined) {
    const parentValue = parentQuestionName ? inputs[parentQuestionName] : undefined;
    if (Array.isArray(condition.equals)) {
      return Array.isArray(parentValue) && arraysEqual(parentValue, condition.equals);
    }
    return parentValue === condition.equals;
  }

  // StringValidation with `enum`
  if ("enum" in condition && condition.enum !== undefined) {
    const parentValue = parentQuestionName ? inputs[parentQuestionName] : undefined;
    return typeof parentValue === "string" && condition.enum.includes(parentValue);
  }

  // StringArrayValidation with `contains`
  if ("contains" in condition && condition.contains !== undefined) {
    const parentValue = parentQuestionName ? inputs[parentQuestionName] : undefined;
    return Array.isArray(parentValue) && parentValue.includes(condition.contains);
  }

  // Default: pass
  return true;
}

function arraysEqual(a: unknown[], b: unknown[]): boolean {
  return a.length === b.length && a.every((v, i) => v === b[i]);
}

/**
 * Ask a single question via the appropriate UI method.
 */
async function askQuestion(
  question: Question,
  ui: UserInteraction,
  inputs: Inputs
): Promise<Result<InputResult<unknown>, FxError>> {
  const title =
    typeof question.title === "function" ? await question.title(inputs) : question.title;
  const placeholder =
    "placeholder" in question && typeof question.placeholder === "function"
      ? await question.placeholder(inputs)
      : "placeholder" in question
        ? question.placeholder
        : undefined;

  switch (question.type) {
    case "singleSelect": {
      const q = question;
      let options = q.staticOptions;
      if (q.dynamicOptions) {
        options = await q.dynamicOptions(inputs);
      }

      // Auto-skip if single option and skipSingleOption is true
      const skipSingle =
        typeof q.skipSingleOption === "function"
          ? await q.skipSingleOption(inputs)
          : q.skipSingleOption;
      if (skipSingle && options.length === 1) {
        const val = typeof options[0] === "string" ? options[0] : options[0].id;
        (question as any).value = val;
        return ok({ type: "success", result: val } as InputResult<unknown>);
      }

      const defaultVal = typeof q.default === "function" ? await q.default(inputs) : q.default;
      return (await ui.selectOption({
        name: question.name,
        title: title ?? "",
        options,
        default: defaultVal,
        placeholder: placeholder as string | undefined,
        returnObject: q.returnObject,
      })) as Result<InputResult<unknown>, FxError>;
    }

    case "multiSelect": {
      const q = question as any;
      let options = q.staticOptions;
      if (q.dynamicOptions) {
        options = await q.dynamicOptions(inputs);
      }
      const defaultVal = typeof q.default === "function" ? await q.default(inputs) : q.default;
      return (await ui.selectOptions({
        name: question.name,
        title: title ?? "",
        options,
        default: defaultVal,
        placeholder: placeholder as string | undefined,
      })) as Result<InputResult<unknown>, FxError>;
    }

    case "text": {
      const q = question as any;
      const defaultVal = typeof q.default === "function" ? await q.default(inputs) : q.default;
      return (await ui.inputText({
        name: question.name,
        title: title ?? "",
        default: defaultVal,
        password: q.password,
        placeholder: placeholder as string | undefined,
      })) as Result<InputResult<unknown>, FxError>;
    }

    case "singleFile": {
      const q = question as any;
      const defaultVal = typeof q.default === "function" ? await q.default(inputs) : q.default;
      return (await ui.selectFile({
        name: question.name,
        title: title ?? "",
        default: defaultVal,
        filters: q.filters,
      })) as Result<InputResult<unknown>, FxError>;
    }

    case "multiFile": {
      return (await ui.selectFiles({
        name: question.name,
        title: title ?? "",
      })) as Result<InputResult<unknown>, FxError>;
    }

    case "folder": {
      const q = question as any;
      const defaultVal = typeof q.default === "function" ? await q.default(inputs) : q.default;
      return (await ui.selectFolder({
        name: question.name,
        title: title ?? "",
        default: defaultVal,
      })) as Result<InputResult<unknown>, FxError>;
    }

    case "singleFileOrText": {
      const q = question as any;
      if (ui.selectFileOrInput) {
        return (await ui.selectFileOrInput({
          name: question.name,
          title: title ?? "",
          inputOptionItem: q.inputOptionItem,
          inputBoxConfig: {
            name: q.inputBoxConfig.name,
            title: q.inputBoxConfig.title ?? "",
            placeholder: q.inputBoxConfig.placeholder,
          },
          filters: q.filters,
        })) as Result<InputResult<unknown>, FxError>;
      }
      // Fallback to selectFile
      return (await ui.selectFile({
        name: question.name,
        title: title ?? "",
        filters: q.filters,
      })) as Result<InputResult<unknown>, FxError>;
    }

    case "confirm": {
      if (ui.confirm) {
        return (await ui.confirm({
          name: question.name,
          title: title ?? "",
        })) as Result<InputResult<unknown>, FxError>;
      }
      return ok({ type: "success", result: true } as InputResult<unknown>);
    }

    default:
      return ok({ type: "skip" } as InputResult<unknown>);
  }
}
