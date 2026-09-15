// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { OptionItem } from "../collectInputs/collectInputs";

export interface DescriptorLanguages {
  languages: string[];
  options: OptionItem[];
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isPresentation(value: unknown): value is OptionItem {
  return (
    isRecord(value) &&
    typeof value.id === "string" &&
    Object.keys(value).every((key) => ["id", "label", "description", "keyPrefix"].includes(key)) &&
    ["label", "description", "keyPrefix"].every(
      (key) => value[key] === undefined || typeof value[key] === "string"
    )
  );
}

export function readDescriptorLanguages(
  descriptor: unknown
): DescriptorLanguages | { error: string } {
  const data = isRecord(descriptor) ? descriptor : {};
  const declared = Array.isArray(data.languages)
    ? data.languages.filter((value): value is string => typeof value === "string")
    : [];
  const languages = declared.length > 0 ? declared : ["common"];
  const overrides = new Map<string, OptionItem>();
  if (data.languageOptions !== undefined) {
    if (!Array.isArray(data.languageOptions)) {
      return { error: "languageOptions must be an array" };
    }
    for (const option of data.languageOptions) {
      if (!isPresentation(option)) {
        return { error: "languageOptions may contain only an id and string presentation fields" };
      }
      if (!declared.includes(option.id)) {
        return { error: `languageOptions references undeclared language '${option.id}'` };
      }
      if (overrides.has(option.id)) {
        return { error: `languageOptions repeats language '${option.id}'` };
      }
      overrides.set(option.id, option);
    }
  }
  return {
    languages,
    options: languages.map((id) => ({ id, ...overrides.get(id) })),
  };
}
