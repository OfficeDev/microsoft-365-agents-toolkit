// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export interface CopyPolicy {
  allowExistingFile: boolean;
  policy: "add" | "skip" | "error";
}

export const playground: Record<string, CopyPolicy> = {
  "package.json": { allowExistingFile: true, policy: "add" },
  ".vscode/launch.json": { allowExistingFile: true, policy: "add" },
  ".vscode/tasks.json": { allowExistingFile: true, policy: "add" },
  "env/.env.playground": { allowExistingFile: false, policy: "skip" },
  "env/.env.playground.user": { allowExistingFile: false, policy: "skip" },
  "m365agents.playground.json": { allowExistingFile: false, policy: "error" },
  ".localConfigs.playground": { allowExistingFile: true, policy: "skip" },
};

export const policys: Record<string, Record<string, CopyPolicy>> = {
  "playground-typescript": playground,
};
