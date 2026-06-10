// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
export interface DeclarativeCopilotManifestValidationResult {
  id: string;
  filePath: string;
  validationResult: string[];
  actionValidationResult: PluginManifestValidationResult[];
  skillValidationResult: SkillValidationResult[];
}

export interface PluginManifestValidationResult {
  id: string;
  filePath: string;
  validationResult: string[];
}

export interface SkillValidationResult {
  folder: string;
  filePath: string;
  validationResult: string[];
}
