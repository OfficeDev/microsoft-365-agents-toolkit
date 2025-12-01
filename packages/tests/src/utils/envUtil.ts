// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
export type DotenvOutput = {
  [k: string]: string;
};

const KEY_VALUE_PAIR_RE = /^\s*([\w.-]+)\s*=\s*(.*)?\s*$/;
const NEW_LINE_RE = /\\n/g;
const NEW_LINE_SPLITTER = /\r?\n/;
const NEW_LINE = "\n";
type DotenvParsedLine =
  | string
  | { key: string; value: string; comment?: string };
export interface DotenvParseResult {
  lines?: DotenvParsedLine[];
  obj: DotenvOutput;
}

export class DotenvUtil {
  deserialize(src: string | Buffer): DotenvParseResult {
    const lines: DotenvParsedLine[] = [];
    const obj: DotenvOutput = {};
    const stringLines = src.toString().split(NEW_LINE_SPLITTER);
    for (const line of stringLines) {
      const kvMatchArray = line.match(KEY_VALUE_PAIR_RE);
      if (kvMatchArray !== null) {
        // match key-value pair
        const key = kvMatchArray[1];
        let value = kvMatchArray[2] || "";
        let inlineComment;
        const dQuoted = value[0] === '"' && value[value.length - 1] === '"';
        const sQuoted = value[0] === "'" && value[value.length - 1] === "'";
        if (sQuoted || dQuoted) {
          value = value.substring(1, value.length - 1);
          if (dQuoted) {
            value = value.replace(NEW_LINE_RE, NEW_LINE);
          }
        } else {
          value = value.trim();
          //try to match comment starter
          const index = value.indexOf("#");
          if (index >= 0) {
            inlineComment = value.substring(index);
            value = value.substring(0, index).trim();
          }
        }
        if (value) obj[key] = value;
        lines.push(
          inlineComment
            ? { key: key, value: value, comment: inlineComment }
            : { key: key, value: value }
        );
      } else {
        lines.push(line);
      }
    }
    return { lines: lines, obj: obj };
  }
}

export const dotenvUtil = new DotenvUtil();

// const original = `# Built-in environment variables
// TEAMSFX_ENV=dev2
// AZURE_SUBSCRIPTION_ID=
// AZURE_RESOURCE_GROUP_NAME=
// RESOURCE_SUFFIX=

// # Generated during provision, you can also add your own variables. If you're adding a secret value, add SECRET_ prefix to the name so Teams Toolkit can handle them properly
// BOT_ID=
// SECRET_APP_CLIENT_SECRET=
// TEAMS_APP_ID=
// BOT_AZURE_FUNCTION_APP_RESOURCE_ID=
// APP_DOMAIN=
// BOT_FUNCTION_ENDPOINT=
// TEAMS_APP_TENANT_ID=
// `;

// const parsed = dotenvUtil.deserialize(original);
// console.log(parsed)
