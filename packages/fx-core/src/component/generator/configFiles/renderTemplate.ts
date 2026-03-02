// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import Handlebars from "handlebars";

export function renderTemplate(templateFilePath: string, context: Record<string, any>): string {
  const templateContent = fs.readFileSync(templateFilePath, "utf-8");
  const template = Handlebars.compile(templateContent);
  const renderedContent = template(context);
  return renderedContent;
}
