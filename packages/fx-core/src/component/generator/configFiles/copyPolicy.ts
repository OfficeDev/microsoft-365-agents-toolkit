// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface CopyPolicy {
  allowExistingFile: boolean;
  policy: "add" | "skip" | "error";
}

export const playgroundNode: Record<string, CopyPolicy> = {
  "package.json": { allowExistingFile: true, policy: "add" },
  ".vscode/launch.json": { allowExistingFile: true, policy: "add" },
  ".vscode/tasks.json": { allowExistingFile: true, policy: "add" },
  "env/.env.playground": { allowExistingFile: false, policy: "skip" },
  "env/.env.playground.user": { allowExistingFile: false, policy: "skip" },
  "m365agents.playground.yml": { allowExistingFile: false, policy: "error" },
  ".localConfigs.playground": { allowExistingFile: true, policy: "skip" },
};

export const localNode: Record<string, CopyPolicy> = {
  "package.json": { allowExistingFile: true, policy: "add" },
  ".vscode/launch.json.tpl": { allowExistingFile: true, policy: "add" },
  ".vscode/tasks.json.tpl": { allowExistingFile: true, policy: "add" },
  "env/.env.local": { allowExistingFile: false, policy: "skip" },
  "m365agents.local.yml.tpl": { allowExistingFile: false, policy: "error" },
};

export const playgroundPython: Record<string, CopyPolicy> = {
  ".vscode/launch.json": { allowExistingFile: true, policy: "add" },
  ".vscode/tasks.json": { allowExistingFile: true, policy: "add" },
  "env/.env.playground": { allowExistingFile: false, policy: "skip" },
  "env/.env.playground.user": { allowExistingFile: false, policy: "skip" },
  ".localConfigs.playground": { allowExistingFile: true, policy: "skip" },
  "m365agents.playground.yml": { allowExistingFile: false, policy: "error" },
};

export const localPython: Record<string, CopyPolicy> = {
  ".vscode/launch.json.tpl": { allowExistingFile: true, policy: "add" },
  ".vscode/tasks.json.tpl": { allowExistingFile: true, policy: "add" },
  "env/.env.local": { allowExistingFile: false, policy: "skip" },
  "m365agents.local.yml.tpl": { allowExistingFile: false, policy: "error" },
};

export const policys: Record<string, Record<string, CopyPolicy>> = {
  "playground-typescript": playgroundNode,
  "local-typescript": localNode,
  "playground-python": playgroundPython,
  "local-python": localPython,
};
