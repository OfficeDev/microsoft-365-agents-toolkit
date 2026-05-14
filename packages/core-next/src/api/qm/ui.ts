// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Result } from "neverthrow";
import { FxError } from "../error";
import { Inputs, OptionItem } from "../types";
import { Colors } from "../utils/log";
import { LocalFunc, OnSelectionChangeFunc, StaticOptions } from "./question";

export interface UIConfig<T> {
  name: string;
  title: string;
  placeholder?: string;
  prompt?: string;
  step?: number;
  totalSteps?: number;
  default?: T | (() => Promise<T>) | string;
  validation?: (input: T) => string | undefined | Promise<string | undefined>;
  buttons?: { icon: string; tooltip: string; command: string }[];
  innerStep?: number;
  innerTotalStep?: number;
}

export interface ConfirmConfig extends UIConfig<boolean> {
  transformer?: (value: boolean) => string;
}

export interface SingleSelectConfig extends UIConfig<string> {
  options: StaticOptions | (() => Promise<StaticOptions>);
  default?: string | (() => Promise<string>);
  returnObject?: boolean;
  skipSingleOption?: boolean;
}

export interface MultiSelectConfig extends UIConfig<string[]> {
  options: StaticOptions | (() => Promise<StaticOptions>);
  default?: string[] | (() => Promise<string[]>) | "none" | "all";
  returnObject?: boolean;
  onDidChangeSelection?: OnSelectionChangeFunc;
  skipSingleOption?: boolean;
}

export interface InputTextConfig extends UIConfig<string> {
  password?: boolean;
  default?: string | (() => Promise<string>);
  additionalValidationOnAccept?: (
    input: string
  ) => string | undefined | Promise<string | undefined>;
}

export type SelectFileConfig = UIConfig<string> & {
  filters?: { [name: string]: string[] };
  default?: string | (() => Promise<string>);
  possibleFiles?: { id: string; label: string; description?: string }[];
  defaultFolder?: string | (() => Promise<string>);
};

export type SelectFilesConfig = UIConfig<string[]> & {
  filters?: { [name: string]: string[] };
  default?: string[] | (() => Promise<string[]>);
};

export type SelectFolderConfig = UIConfig<string> & {
  default?: string | (() => Promise<string>);
};

export interface ExecuteFuncConfig extends UIConfig<string> {
  func: LocalFunc<any>;
  inputs: Inputs;
}

export interface SingleFileOrInputConfig extends UIConfig<string> {
  inputOptionItem: OptionItem;
  inputBoxConfig: UIConfig<string>;
  filters?: { [name: string]: string[] };
}

export interface InputResult<T> {
  type: "success" | "skip" | "back";
  result?: T;
  options?: StaticOptions;
}

export type ConfirmResult = InputResult<boolean>;
export type SingleSelectResult = InputResult<string | OptionItem>;
export type MultiSelectResult = InputResult<StaticOptions>;
export type InputTextResult = InputResult<string>;
export type SelectFileResult = InputResult<string>;
export type SelectFilesResult = InputResult<string[]>;
export type SelectFolderResult = InputResult<string>;

export interface UserInteraction {
  confirm?: (config: ConfirmConfig) => Promise<Result<ConfirmResult, FxError>>;
  selectOption: (config: SingleSelectConfig) => Promise<Result<SingleSelectResult, FxError>>;
  selectOptions: (config: MultiSelectConfig) => Promise<Result<MultiSelectResult, FxError>>;
  inputText: (config: InputTextConfig) => Promise<Result<InputTextResult, FxError>>;
  selectFile: (config: SelectFileConfig) => Promise<Result<SelectFileResult, FxError>>;
  selectFiles: (config: SelectFilesConfig) => Promise<Result<SelectFilesResult, FxError>>;
  selectFolder: (config: SelectFolderConfig) => Promise<Result<SelectFolderResult, FxError>>;
  openUrl(link: string): Promise<Result<boolean, FxError>>;
  showMessage(
    level: "info" | "warn" | "error",
    message: string,
    modal: boolean,
    ...items: string[]
  ): Promise<Result<string | undefined, FxError>>;
  showMessage(
    level: "info" | "warn" | "error",
    message: Array<{ content: string; color: Colors }>,
    modal: boolean,
    ...items: string[]
  ): Promise<Result<string | undefined, FxError>>;
  createProgressBar: (title: string, totalSteps: number) => IProgressHandler;
  reload?(): Promise<Result<boolean, FxError>>;
  executeFunction?(config: ExecuteFuncConfig): any | Promise<any>;
  openFile?(filePath: string): Promise<Result<boolean, FxError>>;
  runCommand?(args: {
    cmd: string;
    workingDirectory?: string;
    shell?: string;
    timeout?: number;
    env?: { [k: string]: string };
    shellName?: string;
    iconPath?: string;
  }): Promise<Result<string, FxError>>;
  selectFileOrInput?(
    config: SingleFileOrInputConfig
  ): Promise<Result<InputResult<string>, FxError>>;
  showDiagnosticInfo?(diagnostics: IDiagnosticInfo[]): void;
}

export interface IProgressHandler {
  start: (detail?: string) => Promise<void> | void;
  next: (detail?: string) => Promise<void> | void;
  end: (success: boolean, hideAfterFinish?: boolean) => Promise<void> | void;
  text?: (detail: string) => Promise<void> | void;
}

export enum DiagnosticSeverity {
  Error = 0,
  Warning = 1,
  Information = 2,
  Hint = 3,
}

export interface IDiagnosticInfo {
  filePath: string;
  startLine: number;
  startIndex: number;
  endLine: number;
  endIndex: number;
  message: string;
  severity: DiagnosticSeverity;
  code?: { value: string; link: string };
}
