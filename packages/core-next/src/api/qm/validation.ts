// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, OptionItem } from "../types";

export type ValidateFunc<T> = (
  input: T,
  inputs?: Inputs
) => string | undefined | Promise<string | undefined>;

export interface StaticValidation {
  required?: boolean;
  equals?: unknown;
}

export interface StringValidation extends StaticValidation {
  maxLength?: number;
  minLength?: number;
  pattern?: string;
  enum?: string[];
  startsWith?: string;
  endsWith?: string;
  includes?: string;
  equals?: string;
  notEquals?: string;
  excludesEnum?: string[];
}

export interface StringArrayValidation extends StaticValidation {
  maxItems?: number;
  minItems?: number;
  uniqueItems?: boolean;
  equals?: string[];
  enum?: string[];
  excludes?: string;
  contains?: string;
  containsAll?: string[];
  containsAny?: string[];
}

export interface FuncValidation<
  T extends string | string[] | OptionItem | OptionItem[] | undefined,
> {
  validFunc: ValidateFunc<T>;
}

export type ConditionFunc = (inputs: Inputs) => boolean | Promise<boolean>;

export type ValidationSchema = StringValidation | StringArrayValidation | FuncValidation<any>;
