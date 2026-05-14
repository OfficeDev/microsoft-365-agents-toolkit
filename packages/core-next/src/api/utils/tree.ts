// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result } from "neverthrow";
import { FxError } from "../error";

export interface TreeItem {
  commandId: string;
  label: string;
  callback?: (args: any) => Promise<Result<null, FxError>>;
  parent?: TreeCategory | string;
  contextValue?: string;
  icon?: string;
  subTreeItems?: TreeItem[];
  tooltip?: { value: string; isMarkdown: boolean };
  description?: string;
  isCustom?: boolean;
  expanded?: boolean;
}

export interface TreeProvider {
  refresh: (tree: TreeItem[]) => Promise<Result<null, FxError>>;
  add: (tree: TreeItem[]) => Promise<Result<null, FxError>>;
  remove: (tree: TreeItem[]) => Promise<Result<null, FxError>>;
}

export enum TreeCategory {
  GettingStarted,
  Account,
  Feedback,
  Project,
  Provision,
  Environment,
}
