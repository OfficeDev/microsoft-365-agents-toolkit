// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { ProgrammingLanguage } from "../../../question";

export interface TemplateInfo {
  templateName: string;
  language: ProgrammingLanguage;
  replaceMap?: { [key: string]: string }; // key is the placeholder in the template file, value is the value to replace
  filterFn?: (fileName: string) => boolean; // return true to include the file, false to exclude
  subFolder?: string; // subfolder under the template folder
}
