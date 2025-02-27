// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, OptionItem, ok } from "@microsoft/teamsfx-api";
import { allCapabilities } from "@microsoft/teamsfx-core";
import chalk from "chalk";
import Table from "cli-table3";
import { logger } from "../../commonlib/logger";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { ListFormatOption } from "../common";

export const listTemplatesCommand: CLICommand = {
  name: "templates",
  description: commands["list.templates"].description,
  options: [ListFormatOption],
  defaultInteractiveOption: false,
  handler: (ctx) => {
    const format = ctx.optionValues.format;
    let result;
    if (format === "table") {
      result = jsonToTable(allCapabilities());
    } else {
      result = JSON.stringify(allCapabilities(), null, 2);
    }
    logger.info(result);
    return ok(undefined);
  },
  telemetry: {
    event: TelemetryEvent.ListSample,
  },
};

function jsonToTable(capabilities: OptionItem[]): string {
  let maxUrlLength = 0;
  let maxIdLength = 0;
  let maxLabelLength = 0;
  capabilities.forEach((item) => {
    if (item.data && (item.data as string).length > maxUrlLength) {
      maxUrlLength = (item.data as string).length;
    }
    if (("id: " + item.id).length > maxIdLength) {
      maxIdLength = ("id: " + item.id).length;
    }
    if (item.label.length > maxLabelLength) {
      maxLabelLength = item.label.length;
    }
  });
  maxUrlLength += 2;
  maxIdLength += 2;
  maxLabelLength += 2;

  const col1Length = Math.max(maxIdLength, maxLabelLength);

  maxUrlLength = Math.max(80, maxUrlLength);

  const terminalWidth = process.stdout.isTTY ? process.stdout.columns : 80;

  const table = new Table({
    head: [chalk.cyanBright("Template"), chalk.cyanBright("Description")],
    colAligns: ["left", "left"],
    colWidths: [col1Length, Math.min(maxUrlLength, terminalWidth - col1Length - 3)],
    wordWrap: true,
  });
  capabilities.forEach((item) => {
    const row = [item.label + chalk.gray("\nid: " + item.id)];
    row.push(
      chalk.gray([item.description, item.detail].filter((i) => !!i).join(". ")) +
        "\n" +
        (item.data ? chalk.underline.blue(item.data) : "")
    );
    table.push(row);
  });
  return table.toString();
}
