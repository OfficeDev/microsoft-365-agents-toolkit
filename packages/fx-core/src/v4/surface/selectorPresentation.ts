// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { SystemError, UserInteraction } from "@microsoft/teamsfx-api";
import {
  PresentationOption,
  PresentationQuestion,
  SelectorPresentation,
} from "../buildTarget/parseSelector";
import { PromptResult, RouteQuestion, RouteResolverPort } from "../buildTarget/resolveBuildTarget";
import {
  ExpressionRuntimePort,
  Scope,
  collectFeatureFlagReferences,
  evaluateExpression,
} from "../expression/evaluateExpression";
import { getFeatureFlaggedLabel } from "../../common/localizeUtils";
import { localizePrefixedText } from "./localizePrompt";

function labelWithIcon(label: string, iconPath: string | undefined): string {
  return iconPath === undefined ? label : `$(${iconPath}) ${label}`;
}

function optionLabel(option: PresentationOption, featureFlagNames: ReadonlySet<string>): string {
  let label = localizePrefixedText(option.keyPrefix, "label", option.label);
  for (const featureFlagName of featureFlagNames) {
    label = getFeatureFlaggedLabel(label, featureFlagName);
  }
  return labelWithIcon(label, option.iconPath);
}

export function createSelectorPrompt(
  presentation: SelectorPresentation,
  ui: UserInteraction,
  surface: string,
  flagReader: (name: string) => boolean
): RouteResolverPort["prompt"] {
  const exprPort: ExpressionRuntimePort = { functions: () => undefined, flags: flagReader };
  const byName = new Map<string, PresentationQuestion>(
    presentation.questions.map((question) => [question.name, question])
  );

  return async function prompt(question: RouteQuestion, step: number): Promise<PromptResult> {
    const pq = byName.get(question.name);
    if (pq === undefined) {
      throw new SystemError({
        source: "Scaffold",
        name: "MissingSelectorPresentation",
        message: `The selector has no presentation for question '${question.name}'.`,
      });
    }
    const scope: Scope = { surface };
    const visible: Array<{
      option: PresentationOption;
      featureFlagNames: ReadonlySet<string>;
    }> = [];
    for (const option of pq.staticOptions) {
      let featureFlagNames: ReadonlySet<string> = new Set();
      if (option.condition !== undefined) {
        const references = collectFeatureFlagReferences(option.condition);
        if (references.isErr()) {
          throw references.error;
        }
        featureFlagNames = references.value;
        const gate = evaluateExpression(option.condition, scope, exprPort);
        if (gate.isErr()) {
          throw gate.error;
        }
        if (gate.value !== true) {
          continue;
        }
      }
      visible.push({ option, featureFlagNames });
    }
    const selected = await ui.selectOption({
      name: pq.name,
      title: localizePrefixedText(pq.keyPrefix, "title", pq.title) ?? pq.name,
      placeholder: localizePrefixedText(pq.keyPrefix, "placeholder", pq.placeholder),
      step,
      options: visible.map(({ option, featureFlagNames }) => ({
        id: option.id,
        label: optionLabel(option, featureFlagNames),
        detail: localizePrefixedText(option.keyPrefix, "detail", option.detail),
        groupName: localizePrefixedText(option.keyPrefix, "groupName", option.groupName),
      })),
      returnObject: false,
    });
    if (selected.isErr()) {
      throw selected.error;
    }
    if (selected.value.type === "back") {
      return { kind: "back" };
    }
    const result = selected.value.result;
    if (typeof result === "string") {
      return { kind: "value", value: result };
    }
    return { kind: "value", value: result === undefined ? "" : result.id };
  };
}
