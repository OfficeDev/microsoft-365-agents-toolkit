// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface IDynamicPromptSettings {
  [templateName: string]: IDynamicPromptFormat<unknown>;
}

export type IDynamicPromptFormat<TArgs> = {
  templates: Record<string, string>;
  messages: IDynamicPromptMessageFormat[];
  version: string;
  presets?: IDynamicPromptPresets;
  $__args_type_helper__?: TArgs;
};

export interface IDynamicPromptMessageFormat {
  role: MessageRole;
  entryTemplate: string;
}

export type MessageRole = "system" | "user" | "assistant";

export interface IDynamicPromptParams<TArgs = never> {
  args: TArgs;

  templates: Record<string, string>;
  common: Record<string, string>;
  presets?: IDynamicPromptPresets;

  item?: unknown;
  itemIndex?: number;
  itemOrdinal?: number;
}

export interface IDynamicPromptPresets {
  [key: string]: SingleOrArray<string | number | boolean | undefined | IDynamicPromptPresets>;
}

type SingleOrArray<T> = T | T[];
