// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import metaOsAssets from "./assets/metaOs.json";
import { renderFragment } from "../renderFragment";
import { randomUUID } from "crypto";
import { FxError, SystemError, TeamsManifestWrapper, UserError } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import * as path from "path";
import { Result, err, ok } from "neverthrow";
import { capabilityDeclarations } from "../../capabilities/declarations";
import { RegisteredStep, StepContext } from "../../pipeline/runScaffoldPipeline";
import { defineStep } from "../../pipeline/defineStep";
import { stringParam } from "../../pipeline/stepParams";

/** MetaOS post-render steps. */

const SOURCE = "Scaffold";

/** Engine step name `metaos/unify-project-id`. */
export const STEP_UNIFY_PROJECT_ID = capabilityDeclarations.step.unifyProjectId.id;
/** Engine step name `metaos/upgrade-existing-project`. */
export const STEP_UPGRADE_EXISTING_PROJECT = capabilityDeclarations.step.upgradeExistingProject.id;

const APP_PACKAGE_FOLDER = "appPackage";
const MANIFEST_PATH = "appPackage/manifest.json";
const ENV_PATH = "env/.env.dev";
const PACKAGE_JSON_PATH = "package.json";
const COMMANDS_PATH = "src/commands/commands.ts";
const DEFAULT_MANIFEST_ID = "${{TEAMS_APP_ID}}";
const DEFAULT_DA_ID = "declarativeAgentAlc";
const DEFAULT_COMMAND_FILE_NAME = "commands.js";
const OFFICE_ADDIN_DEBUGGING_VERSION = "6.0.6";

const EXCLUDED_FILES = new Set([
  "README.md",
  "teamsapp.yml",
  "m365agents.yml",
  "package-lock.json",
  "pnpm-lock.yaml",
  "yarn.lock",
]);
const EXCLUDED_FOLDERS = new Set(["node_modules", "env"]);

function systemError(name: string, message: string): SystemError {
  return new SystemError({ source: SOURCE, name, message });
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function nestedRecord(
  record: Record<string, unknown>,
  key: string
): Record<string, unknown> | undefined {
  const value = record[key];
  return isRecord(value) ? value : undefined;
}

function parseJsonObject(
  json: Buffer,
  errorName: string,
  filePath: string
): Result<Record<string, unknown>, FxError> {
  let parsed: unknown;
  try {
    parsed = JSON.parse(json.toString("utf8"));
  } catch {
    return err(systemError(errorName, `'${filePath}' is not valid JSON.`));
  }
  if (!isRecord(parsed)) {
    return err(systemError(errorName, `'${filePath}' must be a JSON object.`));
  }
  return ok(parsed);
}

function readRequired(
  ctx: StepContext,
  filePath: string,
  errorName: string
): Result<Buffer, FxError> {
  const current = ctx.read(filePath);
  if (current === undefined) {
    return err(systemError(errorName, `Cannot read '${filePath}'.`));
  }
  return ok(current);
}

function userError(name: string, message: string): UserError {
  return new UserError({ source: SOURCE, name, message });
}

function withTeamsAppId(envText: string, appId: string): string {
  const lines = envText.split(/\r?\n/);
  let found = false;
  const updated = lines.map((line) => {
    if (line.startsWith("TEAMS_APP_ID=")) {
      found = true;
      return "TEAMS_APP_ID=" + appId;
    }
    return line;
  });
  if (!found) {
    updated.push("TEAMS_APP_ID=" + appId);
  }
  return updated.join("\n").replace(/\n*$/, "\n");
}

function writeJson(ctx: StepContext, filePath: string, value: unknown): void {
  ctx.write(filePath, Buffer.from(JSON.stringify(value, null, 2) + "\n", "utf8"));
}

function readRequiredJsonObject(
  ctx: StepContext,
  filePath: string,
  missingErrorName: string,
  invalidErrorName: string
): Result<Record<string, unknown>, FxError> {
  const raw = readRequired(ctx, filePath, missingErrorName);
  if (raw.isErr()) {
    return err(raw.error);
  }
  return parseJsonObject(raw.value, invalidErrorName, filePath);
}

function unifyProjectId(
  ctx: StepContext,
  manifestPath: string,
  envPath: string
): Result<void, FxError> {
  const manifestContents = readRequired(ctx, manifestPath, "MetaOsManifestMissing");
  if (manifestContents.isErr()) {
    return err(manifestContents.error);
  }
  let manifest: TeamsManifestWrapper;
  try {
    manifest = TeamsManifestWrapper.fromJSON(manifestContents.value.toString("utf8"));
  } catch {
    return err(systemError("MetaOsManifestInvalid", `'${manifestPath}' is not valid JSON.`));
  }
  const appId = randomUUID();
  manifest.setId(appId);
  ctx.write(manifestPath, Buffer.from(manifest.toJSON(), "utf8"));
  const envRaw = ctx.read(envPath);
  const envText = envRaw === undefined ? "" : envRaw.toString("utf8");
  ctx.write(envPath, Buffer.from(withTeamsAppId(envText, appId), "utf8"));
  return ok(undefined);
}

function isExcludedRelativePath(relativePath: string): boolean {
  const normalized = relativePath.replace(/\\/g, "/");
  const segments = normalized.split("/");
  if (segments.some((segment) => EXCLUDED_FOLDERS.has(segment))) {
    return true;
  }
  return EXCLUDED_FILES.has(segments[segments.length - 1] ?? normalized);
}

function copySourceFiles(sourceFolder: string, ctx: StepContext): Result<void, FxError> {
  const root = path.resolve(sourceFolder);
  const walk = (dir: string): void => {
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
      if (entry.isSymbolicLink()) {
        continue;
      }
      const fullPath = path.join(dir, entry.name);
      const relativePath = path.relative(root, fullPath).replace(/\\/g, "/");
      if (isExcludedRelativePath(relativePath)) {
        continue;
      }
      if (entry.isDirectory()) {
        walk(fullPath);
        continue;
      }
      ctx.write(relativePath, fs.readFileSync(fullPath));
    }
  };

  try {
    walk(root);
    return ok(undefined);
  } catch {
    return err(
      userError("MetaOsSourceProjectInvalid", "The Office Add-in project folder could not be read.")
    );
  }
}

function uniqueFileName(ctx: StepContext, baseName: string, extension: string): string {
  let suffix = 0;
  while (true) {
    const candidate = suffix === 0 ? `${baseName}${extension}` : `${baseName}${suffix}${extension}`;
    if (ctx.read(`${APP_PACKAGE_FOLDER}/${candidate}`) === undefined) {
      return candidate;
    }
    suffix++;
  }
}

interface CommandNames {
  word: string;
  excel: string;
  powerpoint: string;
}

function updateManifestForDa(ctx: StepContext, daFilename: string): Result<CommandNames, FxError> {
  const manifestContents = readRequired(ctx, MANIFEST_PATH, "MetaOsManifestMissing");
  if (manifestContents.isErr()) {
    return err(manifestContents.error);
  }
  let manifest: TeamsManifestWrapper;
  try {
    manifest = TeamsManifestWrapper.fromJSON(manifestContents.value.toString("utf8"));
  } catch {
    return err(systemError("MetaOsManifestInvalid", `'${MANIFEST_PATH}' is not valid JSON.`));
  }
  manifest
    .setId(DEFAULT_MANIFEST_ID)
    .removeDeclarativeAgent(DEFAULT_DA_ID)
    .addDeclarativeAgent(DEFAULT_DA_ID, daFilename);
  if (!manifest.hasExtensions()) {
    return err(systemError("MetaOsManifestShape", "No runtimes found in manifest.extensions."));
  }
  const commandIds = manifest.addExtensionRuntimeActions(DEFAULT_COMMAND_FILE_NAME, [
    { baseId: "addfooter", type: "executeDataFunction" },
    { baseId: "fillcolor", type: "executeDataFunction" },
    { baseId: "addtexttoslide", type: "executeDataFunction" },
  ]);
  if (
    commandIds === undefined ||
    commandIds[0] === undefined ||
    commandIds[1] === undefined ||
    commandIds[2] === undefined
  ) {
    return err(
      systemError(
        "MetaOsCommandsRuntimeMissing",
        "No command runtime found in manifest.extensions."
      )
    );
  }
  ctx.write(MANIFEST_PATH, Buffer.from(manifest.toJSON(), "utf8"));
  return ok({ word: commandIds[0], excel: commandIds[1], powerpoint: commandIds[2] });
}

function daManifest(appName: string, actionFilename: string): Record<string, unknown> {
  const agent = structuredClone(metaOsAssets.agent);
  agent.name = renderFragment([agent.name], { appName });
  for (const action of agent.actions) action.file = actionFilename;
  return agent;
}

function actionManifest(appName: string, commandNames: CommandNames): Record<string, unknown> {
  const plugin = structuredClone(metaOsAssets.plugin);
  plugin.name_for_human = renderFragment([plugin.name_for_human], { appName });
  const names = [commandNames.word, commandNames.excel, commandNames.powerpoint];
  plugin.functions.forEach((definition, index) => {
    definition.name = names[index];
  });
  for (const runtime of plugin.runtimes) runtime.run_for_functions = names;
  return plugin;
}

function commandHandlerCode(commandNames: CommandNames): string {
  return renderFragment(metaOsAssets.commandHandlerCode, commandNames);
}

function appendCommandHandlers(
  ctx: StepContext,
  commandNames: CommandNames
): Result<void, FxError> {
  const current = readRequired(ctx, COMMANDS_PATH, "MetaOsCommandsMissing");
  if (current.isErr()) {
    return err(current.error);
  }
  ctx.write(
    COMMANDS_PATH,
    Buffer.from(current.value.toString("utf8") + commandHandlerCode(commandNames), "utf8")
  );
  return ok(undefined);
}

function upgradeOfficeAddinDebugging(ctx: StepContext): Result<void, FxError> {
  const packageJson = readRequiredJsonObject(
    ctx,
    PACKAGE_JSON_PATH,
    "MetaOsPackageJsonMissing",
    "MetaOsPackageJsonInvalid"
  );
  if (packageJson.isErr()) {
    return err(packageJson.error);
  }
  let devDependencies = nestedRecord(packageJson.value, "devDependencies");
  if (devDependencies === undefined) {
    devDependencies = {};
    packageJson.value.devDependencies = devDependencies;
  }
  devDependencies["office-addin-debugging"] = OFFICE_ADDIN_DEBUGGING_VERSION;
  writeJson(ctx, PACKAGE_JSON_PATH, packageJson.value);
  return ok(undefined);
}

function extendToDeclarativeAgent(ctx: StepContext, appName: string): Result<void, FxError> {
  const daFilename = uniqueFileName(ctx, "declarativeAgent", ".json");
  const actionFilename = uniqueFileName(ctx, "alchemy-plugin", ".json");
  const commandNames = updateManifestForDa(ctx, daFilename);
  if (commandNames.isErr()) {
    return err(commandNames.error);
  }

  writeJson(ctx, `${APP_PACKAGE_FOLDER}/${daFilename}`, daManifest(appName, actionFilename));
  writeJson(
    ctx,
    `${APP_PACKAGE_FOLDER}/${actionFilename}`,
    actionManifest(appName, commandNames.value)
  );

  const appended = appendCommandHandlers(ctx, commandNames.value);
  if (appended.isErr()) {
    return err(appended.error);
  }
  return upgradeOfficeAddinDebugging(ctx);
}

/** Registered step for mirroring v3 MetaOSHelper.unifyProjectID. */
export const metaOsUnifyProjectId: RegisteredStep = defineStep({
  parse(resolved): Result<{ manifestPath: string; envPath: string }, string> {
    const manifestPath = stringParam(resolved, "manifestPath");
    if (manifestPath === undefined) {
      return err("missing string parameter 'manifestPath'");
    }
    const envPath = stringParam(resolved, "envPath");
    if (envPath === undefined) {
      return err("missing string parameter 'envPath'");
    }
    return ok({ manifestPath, envPath });
  },
  invalidParams: () => systemError("MetaOsUnifyParams", "resolved parameters are not all strings"),
  apply({ manifestPath, envPath }, ctx): Result<void, FxError> {
    return unifyProjectId(ctx, manifestPath, envPath);
  },
});

/** Registered step for mirroring v3 MetaOSHelper copy + extend + unify for upgrade. */
export const metaOsUpgradeExistingProject: RegisteredStep = defineStep({
  parse(resolved): Result<{ sourceFolder: string; appName: string }, string> {
    const sourceFolder = stringParam(resolved, "sourceFolder");
    if (sourceFolder === undefined) {
      return err("missing string parameter 'sourceFolder'");
    }
    const appName = stringParam(resolved, "appName");
    if (appName === undefined) {
      return err("missing string parameter 'appName'");
    }
    return ok({ sourceFolder, appName });
  },
  invalidParams: () =>
    systemError("MetaOsUpgradeParams", "resolved parameters are not all strings"),
  apply({ sourceFolder, appName }, ctx): Result<void, FxError> {
    const copied = copySourceFiles(sourceFolder, ctx);
    if (copied.isErr()) {
      return err(copied.error);
    }
    const extended = extendToDeclarativeAgent(ctx, appName);
    if (extended.isErr()) {
      return err(extended.error);
    }
    return unifyProjectId(ctx, MANIFEST_PATH, ENV_PATH);
  },
});
