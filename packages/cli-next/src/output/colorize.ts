// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import chalk from "chalk";

export enum TextType {
  Success = "success",
  Error = "error",
  Warning = "warning",
  Info = "info",
  Hyperlink = "hyperlink",
  Email = "email",
  Important = "important",
  Details = "details",
  Commands = "commands",
}

export function colorize(message: string, type: TextType): string {
  if (!process.stdout.isTTY) {
    return message;
  }
  switch (type) {
    case TextType.Success:
      return chalk.greenBright(message);
    case TextType.Error:
      return chalk.redBright(message);
    case TextType.Warning:
      return chalk.yellowBright(message);
    case TextType.Hyperlink:
      return chalk.cyanBright.underline(message);
    case TextType.Email:
    case TextType.Important:
      return chalk.magentaBright(message);
    case TextType.Details:
      return chalk.gray(message);
    case TextType.Commands:
      return chalk.blueBright(message);
    case TextType.Info:
    default:
      return message;
  }
}
