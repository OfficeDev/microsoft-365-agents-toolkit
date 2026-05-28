// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IDynamicPromptParams } from "./types";

export function buildDynamicPromptInternal(
  expression: string,
  params: IDynamicPromptParams<unknown>
): string {
  expression = expression && expression.trim();
  if (!expression) {
    return "";
  }

  for (const builder of functionBuilders) {
    const match = expression.match(builder.regex);
    if (match) {
      const functionArgs = match[1].split(",").map((arg) => arg.trim());

      return builder.build(functionArgs, params);
    }
  }

  // no other function supported: no '(' or ')' in expression
  if (/[()]/.test(expression)) {
    throw new Error(`Expression "${expression}" is not valid.`);
  }

  let template = getDeepValue(expression, params);
  if (typeof template === "number" || typeof template === "boolean") {
    template = template.toString();
  }

  if (typeof template !== "string") {
    throw new Error(
      `The value of expression "${expression}" is not a string, but typed as "${typeof template}".`
    );
  }

  if (expression.startsWith("args.")) {
    // for args.xxx, use the original value directly for prompt leak prevention
    return template;
  }

  return template.replace(/{{[^{}]+}}/g, (macker) => {
    const subExpression = macker.substring(2, macker.length - 2).trim();
    return buildDynamicPromptInternal(subExpression, params);
  });
}

function getDeepValue<T>(expression: string, params: IDynamicPromptParams<unknown>) {
  // expression should include ony '\w', '_', '$' and '.' in this case.
  if (/[^\w_\$.]/.test(expression)) {
    throw new Error(`Expression "${expression}" is not valid.`);
  }

  const parts = expression.split(".");
  let value: unknown = params;
  for (let i = 0; i < parts.length; i++) {
    if (!value) {
      return undefined;
    }

    value = (value as Record<string, unknown>)[parts[i]];
  }

  return value as T;
}

interface IFunctionBuilder {
  regex: RegExp;
  build: (functionArgs: string[], dynamicPromptParams: IDynamicPromptParams<unknown>) => string;
}

const functionBuilders: IFunctionBuilder[] = [
  {
    // iff(condition, trueValue, falseValue)
    regex: /^\s*iff\s*\(\s*(.+)\s*\)\s*$/,
    build: (args, params) => {
      const conditionExpression = args[0];
      if (getDeepValue(conditionExpression, params)) {
        return buildDynamicPromptInternal(args[1], params);
      } else {
        return buildDynamicPromptInternal(args[2], params);
      }
    },
  },
  {
    // arrayJoin(arrayExpression, itemTemplate, separator)
    regex: /^\s*arrayJoin\s*\(\s*(.+)\s*\)\s*$/,
    build: (args, params) => {
      const [arrayExpression, itemTemplate = "item", separatorExpression = ""] = args;
      const array = getDeepValue(arrayExpression, params) || [];
      if (!Array.isArray(array)) {
        throw new Error(`Expression "${arrayExpression}" is not an array.`);
      }

      if (!array?.length) {
        return "";
      }

      const builtArray = array.map((item, index) =>
        buildDynamicPromptInternal(itemTemplate, {
          ...params,
          item,
          itemIndex: index,
          itemOrdinal: index + 1,
        })
      );

      const separator =
        (separatorExpression && getDeepValue<string>(separatorExpression, params)) || "";

      return builtArray.filter((item) => !!item).join(separator);
    },
  },
  {
    // stringify(value)
    regex: /^\s*stringify\s*\(\s*(.+)\s*\)\s*$/,
    build: (args, params) => {
      const value = getDeepValue(args[0], params);
      if (value === undefined) {
        return "";
      }

      return JSON.stringify(value);
    },
  },
];
