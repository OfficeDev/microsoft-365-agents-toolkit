// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, ok, Platform } from "@microsoft/teamsfx-api";
import {
  featureFlagManager,
  FeatureFlags,
  getAllTemplatesOnPlatform,
} from "@microsoft/teamsfx-core";
import chalk from "chalk";
import Table from "cli-table3";
import { logger } from "../../commonlib/logger";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { ListFormatOption } from "../common";

interface TemplateGroup {
  name: string;
  displayName: string;
  description: string;
  language: string;
}

export function listAllTemplates(): TemplateGroup[] {
  let templates = getAllTemplatesOnPlatform(Platform.VSCode);
  if (featureFlagManager.getBooleanValue(FeatureFlags.CLIDotNet)) {
    templates = getAllTemplatesOnPlatform(Platform.VS);
  }

  // Group by template name, ignoring programming language
  const groupedTemplates = new Map<string, TemplateGroup>();

  templates.forEach((template) => {
    if (!groupedTemplates.has(template.name)) {
      const templateWithDisplay = template as typeof template & { displayName?: string };
      groupedTemplates.set(template.name, {
        name: template.name,
        displayName: templateWithDisplay.displayName || template.name,
        description: template.description,
        language: template.language,
      });
    }
  });

  return Array.from(groupedTemplates.values());
}

export const listTemplatesCommand: CLICommand = {
  name: "templates",
  description: commands["list.templates"].description,
  options: [ListFormatOption],
  defaultInteractiveOption: false,
  handler: (ctx) => {
    const format = ctx.optionValues.format;
    const templates = listAllTemplates();
    let result;
    if (format === "table") {
      result = jsonToTable(templates);
    } else {
      result = JSON.stringify(templates, null, 2);
    }
    logger.info(result);
    return ok(undefined);
  },
  telemetry: {
    event: TelemetryEvent.ListSample,
  },
};

function jsonToTable(templates: TemplateGroup[]): string {
  let maxIdLength = 0;
  let maxNameLength = 0;
  let maxDescriptionLength = 0;
  templates.forEach((template) => {
    if (template.name.length > maxIdLength) {
      maxIdLength = template.name.length;
    }
    if (template.displayName.length > maxNameLength) {
      maxNameLength = template.displayName.length;
    }
    if (template.description.length > maxDescriptionLength) {
      maxDescriptionLength = template.description.length;
    }
  });
  maxIdLength += 2;
  maxNameLength += 2;
  maxDescriptionLength += 2;

  const terminalWidth = process.stdout.isTTY ? process.stdout.columns : 80;
  const idColWidth = Math.max(15, maxIdLength);
  const nameColWidth = Math.max(20, maxNameLength);
  const descColWidth = Math.min(
    maxDescriptionLength,
    terminalWidth - idColWidth - nameColWidth - 4
  );

  const table = new Table({
    head: [
      chalk.cyanBright("Capability(Id)"),
      chalk.cyanBright("Name"),
      chalk.cyanBright("Description"),
    ],
    colAligns: ["left", "left", "left"],
    colWidths: [idColWidth, nameColWidth, descColWidth],
    wordWrap: true,
  });

  templates.forEach((template) => {
    table.push([template.name, template.displayName, template.description]);
  });

  return table.toString();
}
