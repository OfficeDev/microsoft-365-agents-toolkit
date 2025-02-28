// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export const defaultTryLimits = 1;
export const defaultTimeoutInMs = 20000;
export const sampleDefaultTimeoutInMs = 40000;
export const templateFileExt = ".tpl";
export const errorSource = "GEN";
export const componentName = "generator";
export const commonTemplateName = "common";
export const placeholderDelimiters: [string, string] = ["{{", "}}"];
export const oldPlaceholderDelimiters: [string, string] = ["{%", "%}"];
export const sampleConcurrencyLimits = 20;
export const sampleDefaultRetryLimits = 2;

export const declarativeAgentInstructionFileName = "instruction.txt";

export enum OneDriveSharePointItemType {
  File = "File",
  Folder = "Folder",
  Site = "Site",
}
