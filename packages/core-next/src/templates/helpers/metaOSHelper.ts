// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as fs from "fs/promises";
import * as fsSync from "fs";
import * as path from "path";
import * as crypto from "crypto";

const APP_PACKAGE_FOLDER = "appPackage";
const MANIFEST_FILE = "manifest.json";

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
const DEFAULT_DA_ID = "declarativeAgentAlc";

const ENV_FOLDER = "env";
const ENV_FILE_NAME = ".env.dev";

const DEFAULT_CMD_NAME_W = "addfooter";
const DEFAULT_CMD_NAME_X = "fillcolor";
const DEFAULT_CMD_NAME_P = "addtexttoslide";
const DEFAULT_CMD_FILE_NAME = "commands.js";

const DEFAULT_DA_FILENAME = "declarativeAgent";
const DEFAULT_ACTION_FILENAME = "alchemy-plugin";
const FILE_EXTENSION = ".json";

function copyFilterFn(filePath: string): boolean {
  for (const item of NOT_COPY_FILES) {
    if (filePath.endsWith(item)) return false;
  }
  for (const item of NOT_COPY_FOLDERS) {
    if (filePath.includes(item)) return false;
  }
  return true;
}

/**
 * Copy an existing MetaOS project to a new location, filtering out
 * node_modules, env, and lock files.
 */
export async function copyExistMetaOSProject(
  sourceFolder: string,
  targetFolder: string
): Promise<void> {
  await fs.cp(sourceFolder, targetFolder, {
    recursive: true,
    filter: (src: string) => copyFilterFn(src),
  });
}

function getNameWithSuffix(name: string, suffix: number): string {
  return suffix ? `${name}${suffix}` : name;
}

/* eslint-disable @typescript-eslint/no-explicit-any */
function ensureFunctionNameIsNotExist(jsonObj: any[], key: string, functionName: string): string {
  let suffix = 0;
  let conflict: boolean;
  do {
    conflict = false;
    for (const obj of jsonObj) {
      if (obj?.[key] === getNameWithSuffix(functionName, suffix)) {
        suffix++;
        conflict = true;
        break;
      }
    }
  } while (conflict);
  return getNameWithSuffix(functionName, suffix);
}

function ensureFileNameIsNotExist(dirPath: string, filename: string, ext: string): string {
  let suffix = 0;
  while (fsSync.existsSync(path.join(dirPath, `${getNameWithSuffix(filename, suffix)}${ext}`))) {
    suffix++;
  }
  return `${getNameWithSuffix(filename, suffix)}${ext}`;
}

async function readJson(filePath: string): Promise<any> {
  const content = await fs.readFile(filePath, "utf-8");
  return JSON.parse(content);
}

async function writeJson(filePath: string, data: any): Promise<void> {
  await fs.writeFile(filePath, JSON.stringify(data, null, 2), "utf-8");
}

/**
 * Unify the project ID: update manifest.json with a new UUID and
 * sync it into the .env.dev file.
 */
export async function unifyProjectID(projectFolder: string): Promise<void> {
  const manifestPath = path.join(projectFolder, APP_PACKAGE_FOLDER, MANIFEST_FILE);
  const envFilePath = path.join(projectFolder, ENV_FOLDER, ENV_FILE_NAME);

  const manifest = await readJson(manifestPath);
  const newUUID = crypto.randomUUID();
  manifest.id = newUUID;

  try {
    await fs.access(envFilePath);
    const envContent = await fs.readFile(envFilePath, "utf-8");
    const updated = envContent.replace(/^TEAMS_APP_ID=.*$/m, `TEAMS_APP_ID=${newUUID}`);
    if (updated === envContent) {
      // key didn't exist, append
      await fs.appendFile(envFilePath, `\nTEAMS_APP_ID=${newUUID}\n`);
    } else {
      await fs.writeFile(envFilePath, updated, "utf-8");
    }
  } catch {
    // env file doesn't exist — create it
    await fs.mkdir(path.join(projectFolder, ENV_FOLDER), { recursive: true });
    await fs.writeFile(envFilePath, `TEAMS_APP_ID=${newUUID}\n`, "utf-8");
  }

  await writeJson(manifestPath, manifest);
}

async function modifyManifest(
  projectFolder: string,
  DAFilename: string
): Promise<{ w: string; x: string; p: string }> {
  let commandNameW = DEFAULT_CMD_NAME_W;
  let commandNameX = DEFAULT_CMD_NAME_X;
  let commandNameP = DEFAULT_CMD_NAME_P;

  const manifestPath = path.join(projectFolder, APP_PACKAGE_FOLDER, MANIFEST_FILE);
  const manifest = await readJson(manifestPath);

  manifest.id = DEFAULT_MANIFEST_ID;

  manifest.copilotAgents = {
    declarativeAgents: [{ id: DEFAULT_DA_ID, file: DAFilename }],
  };

  const runtimes = manifest.extensions?.[0]?.runtimes;
  if (!runtimes) throw new Error("No runtimes found in manifest.extensions!");

  let added = false;
  for (const runtime of runtimes) {
    if (runtime?.code?.script?.includes(DEFAULT_CMD_FILE_NAME)) {
      const actions = runtime.actions ?? [];
      commandNameW = ensureFunctionNameIsNotExist(actions, "id", commandNameW);
      commandNameX = ensureFunctionNameIsNotExist(actions, "id", commandNameX);
      commandNameP = ensureFunctionNameIsNotExist(actions, "id", commandNameP);

      const newActions = [
        { id: commandNameW, type: "executeDataFunction" },
        { id: commandNameX, type: "executeDataFunction" },
        { id: commandNameP, type: "executeDataFunction" },
      ];
      runtime.actions = [...actions, ...newActions];
      added = true;
      break;
    }
  }
  if (!added) throw new Error("No command's runtime found in manifest.extensions!");

  await writeJson(manifestPath, manifest);
  return { w: commandNameW, x: commandNameX, p: commandNameP };
}

async function generateDAFile(
  projectFolder: string,
  DAFilename: string,
  ActionFilename: string,
  appName: string
): Promise<void> {
  const fileJson = {
    $schema:
      "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.4/schema.json",
    version: "v1.4",
    name: `Add-in Skill + Agent for ${appName}`,
    description:
      "You are an agent for working with add-in. You can work with any cells, not only well formatted table.",
    instructions:
      "You are an agent for working with add-in. You can work with any cells, not only well formatted table.",
    conversation_starters: [
      {
        title: "Change cell color (for excel)",
        text: "Change the cell below A2 to the color of grass. Tell me how long it took in seconds.",
      },
      {
        title: "Add footer (for word)",
        text: "Add a footer with message 'Hello Agent!'. Tell me how long it took in seconds.",
      },
      {
        title: "Add text to slide (for powerpoint)",
        text: "Please add text 'Hello PPT!' to the slide. Tell me how long it took in seconds.",
      },
    ],
    actions: [{ id: "alchemyPlugin", file: ActionFilename }],
  };

  await writeJson(path.join(projectFolder, APP_PACKAGE_FOLDER, DAFilename), fileJson);
}

async function generateActionFile(
  projectFolder: string,
  ActionFilename: string,
  appName: string,
  commandName: { w: string; x: string; p: string }
): Promise<void> {
  const fileJson = {
    $schema: "https://developer.microsoft.com/json-schemas/copilot/plugin/v2.3/schema.json",
    schema_version: "v2.3",
    name_for_human: `Add-in Skill + Agent for ${appName}`,
    description_for_human: "Get answer for user's question related to Microsoft 365 products",
    namespace: "AddInFunctions",
    functions: [
      {
        name: commandName.w,
        description:
          "Action addfooter: take in arg a JSON object, with a footer message in the field 'Footer'.",
        parameters: {
          type: "object",
          properties: {
            Footer: {
              type: "string",
              description: "example message to be added to footer",
              default: "Declarative Agent Footer",
            },
          },
          required: ["Footer"],
        },
      },
      {
        name: commandName.x,
        description:
          "Action fillcolor: take in arg a JSON object, a cell location and a color in hex. Cell location is a single cell.",
        parameters: {
          type: "object",
          properties: {
            Cell: { type: "string", description: "example cell location", default: "B7" },
            Color: { type: "string", description: "example color in hex", default: "#30d5c8" },
          },
          required: ["Cell", "Color"],
        },
      },
      {
        name: commandName.p,
        description:
          "Action addtexttoslide: take in arg a JSON object, a text to be added to a slide.",
        parameters: {
          type: "object",
          properties: {
            Text: {
              type: "string",
              description: "example text to be added to a slide",
              default: "hello declarative agent",
            },
          },
          required: ["Text"],
        },
      },
    ],
    runtimes: [
      {
        type: "LocalPlugin",
        spec: { local_endpoint: "Microsoft.Office.Addin" },
        run_for_functions: [commandName.w, commandName.x, commandName.p],
      },
    ],
  };

  await writeJson(path.join(projectFolder, APP_PACKAGE_FOLDER, ActionFilename), fileJson);
}

async function addCodeToCommands(
  projectFolder: string,
  commandName: { w: string; x: string; p: string }
): Promise<void> {
  const cmdsPath = path.join(projectFolder, "src", "commands", "commands.ts");
  if (!fsSync.existsSync(cmdsPath)) throw new Error("command.ts file doesn't exist!");

  const code = `
/* global Office */
/* global Word, Excel, PowerPoint, performance, console */

async function addFooter(message) {
  await Word.run(async (context) => {
    context.document.sections
      .getFirst()
      .getFooter(Word.HeaderFooterType.primary)
      .insertParagraph(\`From Agent: \${message}\`, "End");
    await context.sync();
  });
}

async function fillColor(cell, color) {
  await Excel.run(async (context) => {
    context.workbook.worksheets.getActiveWorksheet().getRange(cell).format.fill.color = color;
    await context.sync();
  });
}

async function addTextToSlide(text) {
  await PowerPoint.run(async (context) => {
    context.presentation.slides.getItemAt(0).shapes.addTextBox(text, {
      left: Math.random() * 200,
      top: Math.random() * 200,
      height: 150,
      width: 150,
    });
    await context.sync();
  });
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    Office.actions.associate("${commandName.w}", async (message) => {
      const start = performance.now();
      const { Footer: footer } = JSON.parse(message);
      await addFooter(footer);
      const duration = performance.now() - start;
      const result = \`Demo add-in: Footer added! completed in \${duration.toFixed(0)} ms.\`;
      console.log(\`Returning result: "\${result}"\`);
      return result;
    });
  } else if (info.host === Office.HostType.Excel) {
    Office.actions.associate("${commandName.x}", async (message) => {
      const start = performance.now();
      const { Cell: cell, Color: color } = JSON.parse(message);
      await fillColor(cell, color);
      const duration = performance.now() - start;
      const result = \`Demo add-in: Action completed! completed in \${duration.toFixed(0)} ms.\`;
      console.log(\`Returning result: "\${result}"\`);
      return result;
    });
  } else if (info.host === Office.HostType.PowerPoint) {
    Office.actions.associate("${commandName.p}", async (message) => {
      const start = performance.now();
      const { Text: text } = JSON.parse(message);
      await addTextToSlide(text);
      const duration = performance.now() - start;
      const result = \`Demo add-in: text added to slide! completed in \${duration.toFixed(0)} ms.\`;
      console.log(\`Returning result: "\${result}"\`);
      return result;
    });
  }
});
`;

  await fs.appendFile(cmdsPath, code);
}

async function upgradeOfficeAddInDebugging(projectFolder: string): Promise<void> {
  const pkgJsonPath = path.join(projectFolder, "package.json");
  if (!fsSync.existsSync(pkgJsonPath)) throw new Error("package.json file doesn't exist!");

  const pkgJson = await readJson(pkgJsonPath);
  pkgJson.devDependencies = pkgJson.devDependencies ?? {};
  pkgJson.devDependencies["office-addin-debugging"] = "6.0.6";
  await writeJson(pkgJsonPath, pkgJson);
}

/**
 * Extend an existing MetaOS Office add-in project to become a Declarative Agent.
 * Modifies the manifest, generates DA and action files, appends command code,
 * and upgrades the debugging package.
 */
export async function extendToDA(projectFolder: string, appName: string): Promise<void> {
  const DAFilename = ensureFileNameIsNotExist(
    path.join(projectFolder, APP_PACKAGE_FOLDER),
    DEFAULT_DA_FILENAME,
    FILE_EXTENSION
  );
  const ActionFilename = ensureFileNameIsNotExist(
    path.join(projectFolder, APP_PACKAGE_FOLDER),
    DEFAULT_ACTION_FILENAME,
    FILE_EXTENSION
  );

  const commandNames = await modifyManifest(projectFolder, DAFilename);
  await generateDAFile(projectFolder, DAFilename, ActionFilename, appName);
  await generateActionFile(projectFolder, ActionFilename, appName, commandNames);
  await addCodeToCommands(projectFolder, commandNames);
  await upgradeOfficeAddInDebugging(projectFolder);
}
