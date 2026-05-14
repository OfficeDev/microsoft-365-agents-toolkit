// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, OptionItem } from "../types";
import {
  ConditionFunc,
  FuncValidation,
  StringArrayValidation,
  StringValidation,
  ValidationSchema,
} from "./validation";

export interface FunctionRouter {
  namespace: string;
  method: string;
}

export interface Func extends FunctionRouter {
  params?: any;
}

export type LocalFunc<T> = (inputs: Inputs) => T | Promise<T>;

export type OnSelectionChangeFunc = (
  currentSelectedIds: Set<string>,
  previousSelectedIds: Set<string>
) => Promise<Set<string>>;

export type StaticOptions = string[] | OptionItem[];

export type DynamicOptions = LocalFunc<StaticOptions>;

export interface BaseQuestion {
  name: string;
  title?: string | LocalFunc<string | undefined>;
  value?: unknown;
  valueType?: "skip" | "success";
  default?: unknown;
  step?: number;
  totalSteps?: number;
  innerStep?: number;
  innerTotalStep?: number;
  forgetLastValue?: boolean;
  buttons?: { icon: string; tooltip: string; command: string }[];
}

export interface UserInputQuestion extends BaseQuestion {
  type:
    | "singleSelect"
    | "multiSelect"
    | "singleFile"
    | "multiFile"
    | "folder"
    | "text"
    | "singleFileOrText"
    | "innerText"
    | "confirm";
  title: string | LocalFunc<string | undefined>;
  placeholder?: string | LocalFunc<string | undefined>;
  prompt?: string | LocalFunc<string | undefined>;
  default?: string | string[] | boolean | LocalFunc<string | string[] | boolean | undefined>;
  validation?: ValidationSchema;
  validationHelp?: string;
  required?: boolean;
  alternativeNames?: string[];
  cliName?: string;
  cliShortName?: string;
  isBoolean?: boolean;
  cliType?: "option" | "argument";
  cliDescription?: string;
  cliHidden?: boolean;
}

export interface SingleSelectQuestion extends UserInputQuestion {
  type: "singleSelect";
  staticOptions: StaticOptions;
  dynamicOptions?: DynamicOptions;
  value?: string | OptionItem;
  default?: string | LocalFunc<string | undefined>;
  returnObject?: boolean;
  skipSingleOption?: boolean | LocalFunc<boolean>;
  cliChoiceListCommand?: string;
  skipValidation?: boolean;
  onDidSelection?: (itemOrId: string | OptionItem, inputs: Inputs) => Promise<void> | void;
}

export interface ConfirmQuestion extends UserInputQuestion {
  type: "confirm";
  transformer?: (value: boolean) => string;
  value?: boolean;
  default?: boolean | LocalFunc<boolean>;
}

export interface MultiSelectQuestion extends UserInputQuestion {
  type: "multiSelect";
  staticOptions: StaticOptions;
  dynamicOptions?: DynamicOptions;
  value?: string[] | OptionItem[];
  default?: string[] | LocalFunc<string[] | undefined> | "none" | "all";
  returnObject?: boolean;
  skipSingleOption?: boolean;
  onDidChangeSelection?: OnSelectionChangeFunc;
  validation?: StringArrayValidation | FuncValidation<string[]>;
  cliChoiceListCommand?: string;
  skipValidation?: boolean;
}

export interface TextInputQuestion extends UserInputQuestion {
  type: "text";
  password?: boolean;
  value?: string;
  default?: string | LocalFunc<string | undefined>;
  validation?: StringValidation | FuncValidation<string>;
  additionalValidationOnAccept?: StringValidation | FuncValidation<string>;
}

export interface InnerTextInputQuestion extends UserInputQuestion {
  type: "innerText";
  password?: boolean;
  value?: string;
  default?: string | LocalFunc<string | undefined>;
  validation?: StringValidation | FuncValidation<string>;
}

export interface SingleFileQuestion extends UserInputQuestion {
  type: "singleFile";
  value?: string;
  default?: string | LocalFunc<string | undefined>;
  validation?: FuncValidation<string>;
  filters?: { [name: string]: string[] };
  defaultFolder?: string | LocalFunc<string | undefined>;
}

export interface MultiFileQuestion extends UserInputQuestion {
  type: "multiFile";
  value?: string[];
  default?: string | LocalFunc<string | undefined>;
  validation?: FuncValidation<string[]>;
}

export interface FolderQuestion extends UserInputQuestion {
  type: "folder";
  value?: string;
  default?: string | LocalFunc<string | undefined>;
  validation?: FuncValidation<string>;
}

export interface SingleFileOrInputQuestion extends UserInputQuestion {
  type: "singleFileOrText";
  inputOptionItem: OptionItem;
  inputBoxConfig: InnerTextInputQuestion;
  filters?: { [name: string]: string[] };
}

export interface Group {
  type: "group";
  name?: string;
}

export type Question =
  | SingleSelectQuestion
  | MultiSelectQuestion
  | TextInputQuestion
  | SingleFileQuestion
  | MultiFileQuestion
  | FolderQuestion
  | SingleFileOrInputQuestion
  | ConfirmQuestion;

export interface IQTreeNode {
  data: Question | Group;
  condition?: StringValidation | StringArrayValidation | ConditionFunc;
  conditionResult?: boolean;
  children?: IQTreeNode[];
  cliOptionDisabled?: "self" | "children" | "all";
  inputsDisabled?: "self" | "children" | "all";
}
