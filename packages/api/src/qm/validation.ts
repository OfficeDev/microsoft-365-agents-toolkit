// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Inputs, OptionItem } from "../types";

export type ValidateFunc<T> = (
  input: T,
  inputs?: Inputs
) => string | undefined | Promise<string | undefined>;

/**
 * Validation for Any Instance Type
 * JSON Schema Validation reference: http://json-schema.org/draft/2019-09/json-schema-validation.html
 */
export interface StaticValidation {
  /**
   * whether the value is required or not, default value is true if it is undefined
   */
  required?: boolean;
  /**
   * An instance validates successfully against this keyword if its value is equal to the value of the keyword.
   */
  equals?: unknown;
}

/**
 * Validation for Strings
 */
export interface StringValidation extends StaticValidation {
  /**
   * A string instance is valid against this keyword if its length is less than, or equal to, the value of this keyword.
   */
  maxLength?: number;
  /**
   * A string instance is valid against this keyword if its length is greater than, or equal to, the value of this keyword.
   */
  minLength?: number;
  /**
   * A string instance is considered valid if the regular expression matches the instance successfully.
   */
  pattern?: string;
  /**
   * A string instance validates successfully against this keyword if its value is equal to one of the elements in this keyword's array value.
   */
  enum?: string[]; // the value must be contained in this list
  /**
   * A string instance is valid against this keyword if the string starts with the value of this keyword.
   */
  startsWith?: string;
  /**
   * A string instance is valid against this keyword if the string ends with the value of this keyword.
   */
  endsWith?: string;
  /**
   * A string instance is valid against this keyword if the string contains the value of this keyword.
   */
  includes?: string;
  /**
   * An instance validates successfully against this keyword if its value is equal to the value of the keyword.
   */
  equals?: string;
  /**
   * An instance validates successfully against this keyword if its value is not equal to the value of the keyword.
   */
  notEquals?: string;

  /**
   * A string instance validates successfully against this keyword if its value does not equal to any of the elements in this keyword's array value.
   */
  excludesEnum?: string[];
}

/**
 * Validation for String Arrays
 */
export interface StringArrayValidation extends StaticValidation {
  /**
   * The value of this keyword MUST be a non-negative integer.
   * An array instance is valid against "maxItems" if its size is less than, or equal to, the value of this keyword.
   */
  maxItems?: number;
  /**
   * The value of this keyword MUST be a non-negative integer.
   * An array instance is valid against "minItems" if its size is greater than, or equal to, the value of this keyword.
   */
  minItems?: number;
  /**
   * If this keyword has boolean value false, the instance validates successfully. If it has boolean value true, the instance validates successfully if all of its elements are unique.
   */
  uniqueItems?: boolean;
  /**
   * An instance validates successfully against this string array if they have the exactly the same elements.
   */
  equals?: string[];
  /**
   * An array instance is valid against "enum" array if all of the elements of the array is contained in the `enum` array.
   */
  enum?: string[];

  /**
   * An array instance is valid against "excludes" if it doesn't contains the value of `excludes`
   */
  excludes?: string;

  /**
   * An array instance is valid against "contains" if it contains the value of `contains`
   */
  contains?: string;
  /**
   * An array instance is valid against "containsAll" array if it contains all of the elements of `containsAll` array.
   */
  containsAll?: string[];
  /**
   * An array instance is valid against "containsAny" array if it contains any one of the elements of `containsAny` array.
   */
  containsAny?: string[]; ///non-standard, the values must contains any one in the array
}

/**
 * The validation is checked by a validFunc provided by user
 */
export interface FuncValidation<
  T extends string | string[] | OptionItem | OptionItem[] | undefined,
> {
  /**
   * A function that will be called to validate input and to give a hint to the user.
   *
   * @param input The current value of the input to be validated.
   * @return A human-readable string which is presented as diagnostic message.
   * Return `undefined` when 'value' is valid.
   */
  validFunc: ValidateFunc<T>;
}

export type ConditionFunc = (inputs: Inputs) => boolean | Promise<boolean>;

/**
 * Definition of validation schema, which is a union of `StringValidation`, `StringArrayValidation` and `FuncValidation<any>`
 */
export type ValidationSchema = StringValidation | StringArrayValidation | FuncValidation<any>;
