// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  select as inquirerSelect,
  checkbox as inquirerCheckbox,
  input as inquirerInput,
  confirm as inquirerConfirm,
} from "@inquirer/prompts";
import {
  UserInteraction,
  SingleSelectConfig,
  MultiSelectConfig,
  InputTextConfig,
  SelectFileConfig,
  SelectFilesConfig,
  SelectFolderConfig,
  SingleFileOrInputConfig,
  InputResult,
  FxError,
  OptionItem,
  StaticOptions,
  IProgressHandler,
  ConfirmConfig,
  ConfirmResult,
  ok,
  Result,
} from "@microsoft/teamsfx-core-next";

/**
 * CLI User Interaction adapter.
 *
 * In interactive mode: delegates to @inquirer/prompts.
 * In non-interactive mode: returns defaults without prompting.
 */
export class CLIUserInteraction implements UserInteraction {
  private _interactive = true;

  get interactive(): boolean {
    return this._interactive && process.env.CI_ENABLED !== "true";
  }

  set interactive(value: boolean) {
    this._interactive = value;
  }

  async selectOption(config: SingleSelectConfig): Promise<Result<InputResult<string>, FxError>> {
    const options = typeof config.options === "function" ? await config.options() : config.options;
    const defaultVal = await resolveDefault(options, config.default);
    if (!this.interactive) {
      return ok({ type: "success", result: defaultVal });
    }

    const choices = options.map((opt) => {
      if (typeof opt === "string") {
        return { name: opt, value: opt };
      }
      const item = opt as OptionItem;
      return {
        name: item.label ?? item.id,
        value: item.id,
        description: item.description,
      };
    });

    const answer = await inquirerSelect({
      message: config.title ?? config.name ?? "Select an option",
      choices,
      default: defaultVal || undefined,
    });
    return ok({ type: "success", result: answer });
  }

  async selectOptions(config: MultiSelectConfig): Promise<Result<InputResult<string[]>, FxError>> {
    const defaults = await resolveArrayDefault(config.default);
    if (!this.interactive) {
      return ok({ type: "success", result: defaults });
    }

    const options = typeof config.options === "function" ? await config.options() : config.options;
    const choices = options.map((opt) => {
      if (typeof opt === "string") {
        return { name: opt, value: opt, checked: defaults.includes(opt) };
      }
      const item = opt as OptionItem;
      return {
        name: item.label ?? item.id,
        value: item.id,
        checked: defaults.includes(item.id),
      };
    });

    const answers = await inquirerCheckbox({
      message: config.title ?? config.name ?? "Select options",
      choices,
    });
    return ok({ type: "success", result: answers });
  }

  async inputText(config: InputTextConfig): Promise<Result<InputResult<string>, FxError>> {
    const defaultVal: string = await resolveStringDefault(config.default, "");
    if (!this.interactive) {
      return ok({ type: "success", result: defaultVal });
    }

    const answer = await inquirerInput({
      message: config.title ?? config.name ?? "Enter a value",
      default: defaultVal || undefined,
      validate: config.validation
        ? async (input: string) => {
            const err = await config.validation!(input);
            return err ?? true;
          }
        : undefined,
    });
    return ok({ type: "success", result: answer });
  }

  async selectFile(config: SelectFileConfig): Promise<Result<InputResult<string>, FxError>> {
    const defaultVal: string = await resolveStringDefault(config.default, "");
    if (!this.interactive) {
      return ok({ type: "success", result: defaultVal });
    }

    const answer = await inquirerInput({
      message: config.title ?? config.name ?? "Enter a file path",
      default: defaultVal || undefined,
    });
    return ok({ type: "success", result: answer });
  }

  async selectFiles(config: SelectFilesConfig): Promise<Result<InputResult<string[]>, FxError>> {
    const defaults: string[] = await resolveArrayDefault(config.default);
    if (!this.interactive) {
      return ok({ type: "success", result: defaults });
    }

    const answer = await inquirerInput({
      message: config.title ?? config.name ?? "Enter file paths (comma-separated)",
      default: defaults.join(", ") || undefined,
    });
    return ok({
      type: "success",
      result: answer
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean),
    });
  }

  async selectFolder(config: SelectFolderConfig): Promise<Result<InputResult<string>, FxError>> {
    const defaultVal: string = await resolveStringDefault(config.default, ".");
    if (!this.interactive) {
      return ok({ type: "success", result: defaultVal });
    }

    const answer = await inquirerInput({
      message: config.title ?? config.name ?? "Enter a folder path",
      default: defaultVal || undefined,
    });
    return ok({ type: "success", result: answer });
  }

  async selectFileOrInput(
    config: SingleFileOrInputConfig
  ): Promise<Result<InputResult<string>, FxError>> {
    return this.inputText(config.inputBoxConfig);
  }

  async confirm(config: ConfirmConfig): Promise<Result<ConfirmResult, FxError>> {
    const defaultVal: boolean =
      typeof config.default === "function"
        ? await config.default()
        : typeof config.default === "boolean"
          ? config.default
          : true;
    if (!this.interactive) {
      return ok({ type: "success" as const, result: defaultVal });
    }

    const answer = await inquirerConfirm({
      message: config.title ?? config.name ?? "Confirm?",
      default: defaultVal,
    });
    return ok({ type: "success" as const, result: answer });
  }

  async openUrl(url: string): Promise<Result<boolean, FxError>> {
    console.log(`  Open URL: ${url}`);
    return ok(true);
  }

  async showMessage(
    level: "info" | "warn" | "error",
    message: string | Array<{ content: string }>,
    _modal: boolean,
    ...items: string[]
  ): Promise<Result<string | undefined, FxError>> {
    const text = typeof message === "string" ? message : message.map((m) => m.content).join("");
    if (level === "error") {
      console.error(text);
    } else if (level === "warn") {
      console.warn(text);
    } else {
      console.log(text);
    }
    return ok(items.length > 0 ? items[0] : undefined);
  }

  createProgressBar(title: string, totalSteps: number): IProgressHandler {
    let currentStep = 0;
    return {
      start: async (detail?: string) => {
        console.log(`[0/${totalSteps}] ${title}${detail ? ": " + detail : ""}`);
      },
      next: async (detail?: string) => {
        currentStep++;
        console.log(`[${currentStep}/${totalSteps}] ${title}${detail ? ": " + detail : ""}`);
      },
      end: async (success: boolean) => {
        console.log(`${success ? "✓" : "✗"} ${title}`);
      },
    };
  }
}

async function resolveDefault(options: StaticOptions, defaultValue?: unknown): Promise<string> {
  if (typeof defaultValue === "string") return defaultValue;
  if (typeof defaultValue === "function") {
    const resolved = await defaultValue();
    return String(resolved);
  }
  if (options.length > 0) {
    const first = options[0];
    return typeof first === "string" ? first : (first as OptionItem).id;
  }
  return "";
}

async function resolveStringDefault(
  value: string | (() => Promise<string>) | undefined,
  fallback: string
): Promise<string> {
  if (typeof value === "function") return value();
  return value ?? fallback;
}

async function resolveArrayDefault(
  value: string[] | (() => Promise<string[]>) | "none" | "all" | undefined
): Promise<string[]> {
  if (typeof value === "function") return value();
  if (Array.isArray(value)) return value;
  return [];
}

export const cliUI = new CLIUserInteraction();
