// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FxError,
  InputTextConfig,
  MultiSelectConfig,
  OptionItem as SurfaceOptionItem,
  SingleSelectConfig,
  SystemError,
  UserInteraction,
} from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { Asked, OptionItem, PromptUI, QuestionSpec } from "../collectInputs/collectInputs";

/**
 * The create-Q2 prompt bridge: adapt the v4 surface-neutral `PromptUI`
 * (`ask` / `askMulti`, the face `collect-inputs` drives) onto the host
 * `UserInteraction` (`selectOption` / `inputText` / `selectOptions`).
 *
 * Spec: docs/03-specs/operations/scaffolding/collect-create-inputs.md
 *       (CCI-06 singleSelect, CCI-07 text, CCI-08 multiSelect)
 *
 * v4-owned (INV-7): `UserInteraction` is `@microsoft/teamsfx-api`, upstream of
 * both the v4 and v3 worlds — not a v3 symbol — so this bridge imports no v3
 * type. A v4 identity-only `OptionItem` carries only presentational fields and
 * its `id` across the bridge (spec INV-3): no v3 `option.data` payload round
 * trips, and the surface result is projected straight back to the selected `id`.
 *
 * The create templates use `singleSelect` / `text` (via `ask`) and `multiSelect`
 * (via `askMulti`) today; any other kind is an explicit `SystemError` rather than
 * a silent mismatch, until a template needs it (spec Notes).
 */

const SOURCE = "Scaffold";

/** Map a v4 identity-only option to the surface option shape (label defaults to its id). */
function toSurfaceOptions(options: OptionItem[]): SurfaceOptionItem[] {
  return options.map((option) => ({
    id: option.id,
    label: option.label ?? option.id,
    description: option.description,
    detail: option.detail,
    groupName: option.groupName,
  }));
}

/** Project a single-select surface result back to the selected `id` string. */
function selectedId(result: string | SurfaceOptionItem | undefined): string {
  if (typeof result === "string") {
    return result;
  }
  if (result === undefined) {
    return "";
  }
  return result.id;
}

/** Project a multi-select surface result back to the selected `id` strings. */
function selectedIds(result: string[] | SurfaceOptionItem[] | undefined): string[] {
  if (result === undefined) {
    return [];
  }
  return result.map((item) => (typeof item === "string" ? item : item.id));
}

function unsupportedKind(question: QuestionSpec): FxError {
  return new SystemError({
    source: SOURCE,
    name: "UnsupportedQuestionKind",
    message: `Question '${question.name}' has kind '${question.type}', which the create surface does not render.`,
  });
}

/**
 * Build a `PromptUI` over the host `UserInteraction`. The caller's 1-based `step`
 * is threaded onto each config so the host shows a Back button past the first
 * prompt; a host `back` result is projected to `{ kind: "back" }` (spec INPUT-16),
 * otherwise the selected `id` / text is returned, keeping `collect-inputs`
 * surface-neutral.
 */
export function createUiPromptUI(ui: UserInteraction): PromptUI {
  return {
    async ask(
      question: QuestionSpec,
      options: OptionItem[] | undefined,
      step?: number
    ): Promise<Result<Asked<string>, FxError>> {
      if (question.type === "singleSelect") {
        if (options === undefined) {
          return err(unsupportedKind(question));
        }
        const config: SingleSelectConfig = {
          name: question.name,
          title: question.title ?? question.name,
          placeholder: question.placeholder,
          prompt: question.prompt,
          options: toSurfaceOptions(options),
          returnObject: false,
          step,
        };
        const result = await ui.selectOption(config);
        if (result.isErr()) {
          return err(result.error);
        }
        if (result.value.type === "back") {
          return ok({ kind: "back" });
        }
        return ok({ kind: "value", value: selectedId(result.value.result) });
      }
      if (question.type === "text") {
        const config: InputTextConfig = {
          name: question.name,
          title: question.title ?? question.name,
          placeholder: question.placeholder,
          prompt: question.prompt,
          step,
        };
        const result = await ui.inputText(config);
        if (result.isErr()) {
          return err(result.error);
        }
        if (result.value.type === "back") {
          return ok({ kind: "back" });
        }
        return ok({ kind: "value", value: result.value.result ?? "" });
      }
      return err(unsupportedKind(question));
    },

    async askMulti(
      question: QuestionSpec,
      options: OptionItem[] | undefined,
      step?: number
    ): Promise<Result<Asked<string[]>, FxError>> {
      if (question.type !== "multiSelect" || options === undefined) {
        return err(unsupportedKind(question));
      }
      const config: MultiSelectConfig = {
        name: question.name,
        title: question.title ?? question.name,
        placeholder: question.placeholder,
        prompt: question.prompt,
        options: toSurfaceOptions(options),
        returnObject: false,
        step,
      };
      const result = await ui.selectOptions(config);
      if (result.isErr()) {
        return err(result.error);
      }
      if (result.value.type === "back") {
        return ok({ kind: "back" });
      }
      return ok({ kind: "value", value: selectedIds(result.value.result) });
    },
  };
}
