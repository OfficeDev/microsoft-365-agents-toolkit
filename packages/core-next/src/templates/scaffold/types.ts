// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Convert a full language name to the short key used in template zip archives.
 * Matches the fx-core `convertToLangKey()` convention.
 */
export function convertToLangKey(language: string): string {
  switch (language) {
    case "typescript":
      return "ts";
    case "javascript":
      return "js";
    case "csharp":
      return "csharp";
    case "python":
      return "python";
    case "common":
    case "none":
      return "common";
    default:
      return language;
  }
}

/**
 * Information needed to scaffold one template from the repository.
 */
export interface TemplateInfo {
  /** The template name (e.g. "copilot-gpt-basic") matching the folder in templates/ */
  templateName: string;
  /** Language variant (e.g. "ts", "js", "python", "csharp", "common") */
  language: string;
  /** Mustache variable replacements applied to .tpl files and file names */
  replaceMap?: Record<string, string>;
  /** Optional file filter — return true to include a file from the zip */
  filterFn?: (fileName: string) => boolean;
  /** Optional subfolder under destPath to scaffold into */
  subFolder?: string;
}

/**
 * Context passed through the scaffold pipeline.
 */
export interface ScaffoldContext {
  /** Template name (same as TemplateInfo.templateName) */
  name: string;
  /** Language key */
  language: string;
  /** Destination directory */
  destination: string;
  /** Retry limit for remote download */
  tryLimits: number;
  /** Timeout in ms for remote download */
  timeoutInMs: number;
  /** Whether to use local fallback */
  useFallback: boolean;
  /** List of files written (populated by scaffold actions) */
  outputs: string[];
  /** File name transform (e.g. strip .tpl, render Mustache in name) */
  fileNameReplaceFn?: (name: string, data: Buffer) => string;
  /** File content transform (render Mustache in .tpl files) */
  fileDataReplaceFn?: (name: string, data: Buffer) => string | Buffer;
  /** File inclusion filter */
  filterFn?: (name: string) => boolean;
}

/**
 * Configuration for the template download URL pattern.
 */
export interface TemplateConfig {
  /** Base URL for downloading template zips (e.g. GitHub releases) */
  downloadBaseUrl: string;
  /** Tag prefix for versioned releases */
  tagPrefix: string;
  /** File extension for template archives */
  archiveExt: string;
  /** Version to download — "local" means skip remote, use fallback */
  version: string;
}
