// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/* eslint-disable @typescript-eslint/no-explicit-any */

import path from "path";
import fse from "fs-extra";
import {
  AppManifestUtils,
  AppPackageFolderName,
  DeclarativeAgentManifestV1D3,
  ManifestTemplateFileName,
  TeamsManifestVDevPreview,
} from "@microsoft/teamsfx-api";
import { dotenvUtil } from "../../utils/envUtil";
import { getUuid } from "../../../common/stringUtils";

const NOT_COPY_FILES = [
  "README.md",
  "teamsapp.yml",
  "m365agents.yml",
  "package-lock.json",
  "pnpm-lock.yaml",
  "yarn.lock",
];
const NOT_COPY_FOLDERS = ["node_modules", "env"];

const DEFAULT_MANIFEST_ID = "${{TEAMS_APP_ID}}";
const DEFAULT_DA_ID = "declarativeCopilotAlc";

const ENV_FOLDER = "env";
const ENV_FILE_NAME = ".env.dev";

const PKG_JSON_FILE_NAME = "package.json";

const DEFAULT_CMD_NAME = "fillcolor";
const DEFAULT_CMD_FILE_NAME = "commands.js";

const DEFAULT_DA_FILENAME = "declarativeCopilot";
const DEFAULT_ACTION_FILENAME = "Excel-API-local-plugin";
const FILE_EXTENSION = ".json";

export class MetaOSHelper {
  static copyFilterFn(filePath: string): boolean {
    for (const item of NOT_COPY_FILES) {
      if (filePath.endsWith(item)) {
        return false;
      }
    }

    for (const item of NOT_COPY_FOLDERS) {
      if (filePath.includes(item)) {
        return false;
      }
    }

    return true;
  }

  static async copyExistMetaOSProject(sourceFolder: string, targetFolder: string): Promise<void> {
    await fse.copy(sourceFolder, targetFolder, {
      filter: MetaOSHelper.copyFilterFn,
    });
  }

  static getNameWithSuffix(name: string, suffix: number): string {
    return suffix ? `${name}${suffix}` : name;
  }

  static ensureFunctionNameIsNotExist(jsonObj: any[], key: string, functionName: string): string {
    let suffix = 0;
    let nameConflict = false;

    do {
      nameConflict = false;
      for (const obj of jsonObj) {
        if (obj?.[key] === MetaOSHelper.getNameWithSuffix(functionName, suffix)) {
          suffix++;
          nameConflict = true;
          break;
        }
      }
    } while (nameConflict);

    return MetaOSHelper.getNameWithSuffix(functionName, suffix);
  }

  static ensureFileNameIsNotExist(filePath: string, filename: string, ext: string): string {
    let suffix = 0;

    while (
      fse.existsSync(path.join(filePath, MetaOSHelper.getNameWithSuffix(filename, suffix), ext))
    ) {
      suffix++;
    }

    return `${MetaOSHelper.getNameWithSuffix(filename, suffix)}${ext}`;
  }

  static async unifyProjectID(projectFolder: string): Promise<void> {
    const manifestPath = path.join(projectFolder, AppPackageFolderName, ManifestTemplateFileName);
    const envFilePath = path.join(projectFolder, ENV_FOLDER, ENV_FILE_NAME);

    const manifest = (await AppManifestUtils.readTeamsManifest(
      manifestPath
    )) as TeamsManifestVDevPreview;

    // use dotenvUtil rather than envUtil to avoid touch to the process.env
    const envVars = dotenvUtil.deserialize(await fse.readFile(envFilePath, { encoding: "utf8" }));

    const newUUID = getUuid();
    manifest.id = newUUID;
    envVars.obj.TEAMS_APP_ID = newUUID;

    await AppManifestUtils.writeTeamsManifest(manifestPath, manifest);
    await fse.writeFile(envFilePath, dotenvUtil.serialize(envVars), { encoding: "utf8" });
  }

  static async extendToDA(projectFolder: string, appName: string): Promise<void> {
    // Ensure schema files name
    const DAFilename = MetaOSHelper.ensureFileNameIsNotExist(
      projectFolder,
      DEFAULT_DA_FILENAME,
      FILE_EXTENSION
    );
    const ActionFilename = MetaOSHelper.ensureFileNameIsNotExist(
      projectFolder,
      DEFAULT_ACTION_FILENAME,
      FILE_EXTENSION
    );

    // Modify manifest.json
    const commandName = await MetaOSHelper.modifyManifest(projectFolder, DAFilename);

    // generate DA files
    await MetaOSHelper.generateDAFile(projectFolder, DAFilename, ActionFilename, appName);
    await MetaOSHelper.generateActionFile(projectFolder, ActionFilename, appName, commandName);

    // Add functions to command.ts
    await MetaOSHelper.addCodeToCommands(projectFolder, commandName);

    // Upgrade office-addin-debugging
    await MetaOSHelper.upgradeOfficeAddInDebugging(projectFolder);
  }

  static async modifyManifest(projectFolder: string, DAFilename: string): Promise<string> {
    let commandName = DEFAULT_CMD_NAME;

    const manifestPath = path.join(projectFolder, AppPackageFolderName, ManifestTemplateFileName);
    const manifest = (await AppManifestUtils.readTeamsManifest(
      manifestPath
    )) as TeamsManifestVDevPreview;

    // Update manifest GUID
    manifest.id = DEFAULT_MANIFEST_ID;

    // Add the DA definition
    manifest.copilotAgents = {
      declarativeAgents: [
        {
          id: DEFAULT_DA_ID,
          file: DAFilename,
        },
      ],
    };

    // Add the command to command's runtime
    const runtimes = manifest.extensions?.[0]?.runtimes;
    if (runtimes) {
      let added = false;
      for (const runtime of runtimes) {
        if (runtime?.code?.script?.includes(DEFAULT_CMD_FILE_NAME)) {
          if (runtime.actions) {
            commandName = MetaOSHelper.ensureFunctionNameIsNotExist(
              runtime.actions,
              "id",
              commandName
            );
            runtime.actions.push({
              id: commandName,
              type: "executeDataFunction",
            });
          } else {
            runtime.actions = [
              {
                id: commandName,
                type: "executeDataFunction",
              },
            ];
          }
          added = true;
          break;
        }
      }
      if (!added) {
        throw new Error("No command's runtime found in manifest.extensions!");
      }
    } else {
      throw new Error("No runtimes found in manifest.extensions!");
    }

    // save file and return
    await AppManifestUtils.writeTeamsManifest(manifestPath, manifest);

    return commandName;
  }

  static async generateDAFile(
    projectFolder: string,
    DAFilename: string,
    ActionFilename: string,
    appName: string
  ): Promise<void> {
    const fileJson: DeclarativeAgentManifestV1D3 = {
      version: "v1.3",
      name: `Excel Add-in + Agent for ${appName}`,
      description: "Agent for working with Excel cells.",
      instructions:
        "You are an agent for working with an add-in. You can work with any cells, not just a well-formatted table.",
      conversation_starters: [
        {
          title: "Change cell color",
          text: "I want to change the color of cell B2 to orange",
        },
      ],
      actions: [
        {
          id: "localExcelPlugin",
          file: ActionFilename,
        },
      ],
    };

    await AppManifestUtils.writeDeclarativeAgentManifest(
      path.join(projectFolder, AppPackageFolderName, DAFilename),
      fileJson
    );
  }

  static async upgradeOfficeAddInDebugging(projectFolder: string): Promise<void> {
    const pkgJsonPath = path.join(projectFolder, PKG_JSON_FILE_NAME);
    if (fse.existsSync(pkgJsonPath)) {
      const pkgJson = await fse.readJSON(pkgJsonPath);
      pkgJson["devDependencies"]["office-addin-debugging"] = "^6.0.4";
      await fse.writeJSON(pkgJsonPath, pkgJson, { spaces: 2 });
    } else {
      throw new Error(`package.json file doesn't exist!`);
    }
  }

  static async generateActionFile(
    projectFolder: string,
    ActionFilename: string,
    appName: string,
    commandName: string
  ): Promise<void> {
    // TODO: as any for temporary, since the runtime type `localPlugin` is not type defined yet
    const fileJson: any = {
      schema_version: "v2.2",
      name_for_human: `Excel Add-in + Agent for ${appName}`,
      description_for_human: `Add-in Actions in Agents for ${appName}`,
      namespace: "addin_function",
      functions: [
        {
          name: `${commandName}`,
          description: `${commandName} changes a single cell location to a specific color.`,
          parameters: {
            type: "object",
            properties: {
              Cell: {
                type: "string",
                description: "A cell location in the format of A1, B2, etc.",
                default: "B2",
              },
              Color: {
                type: "string",
                description: "A color in hex format, e.g. #30d5c8",
                default: "#30d5c8",
              },
            },
            required: ["Cell", "Color"],
          },
          returns: {
            type: "string",
            description: "A string indicating the result of the action.",
          },
          states: {
            reasoning: {
              description: `\`${commandName}\` changes the color of a single cell based on the grid location and a color value.`,
              instructions:
                "The user will pass ask for a color that isn't in the hex format needed in most cases, make sure to convert to the closest approximation in the right format.",
            },
            responding: {
              description: `\`${commandName}\` changes the color of a single cell based on the grid location and a color value.`,
              instructions:
                "If there is no error present, tell the user the cell location and color that was set.",
            },
          },
        },
      ],
      runtimes: [
        {
          type: "LocalPlugin",
          spec: {
            local_endpoint: "Microsoft.Office.Addin",
          },
          run_for_functions: [`${commandName}`],
        },
      ],
    };

    const filePath = path.join(projectFolder, AppPackageFolderName, ActionFilename);
    // directly write JSON to avoid type check for not type defined runtime.type `LocalPlugin`
    await fse.writeJSON(filePath, fileJson, { spaces: 2 });
  }

  static async addCodeToCommands(projectFolder: string, commandName: string): Promise<void> {
    if (!fse.existsSync(path.join(projectFolder, "src", "commands", "commands.ts"))) {
      throw new Error("command.ts file doesn't exist!");
    }

    const codeToAppend = `
/* global Excel, performance, console */

async function ${commandName}(cell, color) {
  await Excel.run(async (context) => {
    context.workbook.worksheets.getActiveWorksheet().getRange(cell).format.fill.color = color;
    await context.sync();
  });
}
// eslint-disable-next-line @typescript-eslint/no-unused-vars
Office.onReady((info) => {
  Office.actions.associate("${commandName}", async (message) => {
    const start = performance.now();
    const { Cell: cell, Color: color } = JSON.parse(message);
    await ${commandName}(cell, color);
    const duration = performance.now() - start;
    const result = \`Demo add-in: Action completed! completed in \$\{duration.toFixed(0)\} ms.\`;
    console.log(\`Returning result: "\$\{result\}"\`);
    return result;
  });
});
`;

    await fse.appendFile(path.join(projectFolder, "src", "commands", "commands.ts"), codeToAppend);
  }
}
