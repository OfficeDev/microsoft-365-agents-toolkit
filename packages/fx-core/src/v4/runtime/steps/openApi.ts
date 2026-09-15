// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  AdaptiveCardUpdateStrategy,
  ConstantString,
  GenerateResult,
  ListAPIInfo,
  ProjectType,
  SpecParser,
  SpecParserError,
  Utils,
  ValidationStatus,
} from "@microsoft/m365-spec-parser";
import {
  AppPackageFolderName,
  DefaultApiSpecJsonFileName,
  DefaultApiSpecFolderName,
  DefaultApiSpecYamlFileName,
  DefaultPluginManifestFileName,
  DeclarativeAgentManifestWrapper,
  FxError,
  ManifestTemplateFileName,
  SystemError,
  UserError,
} from "@microsoft/teamsfx-api";
import axios from "axios";
import * as fs from "fs-extra";
import * as path from "path";
import { Result, err, ok } from "neverthrow";
import { SpecParserSource } from "../../../common/constants";
import { generatePlugin, validateOpenAPISpec } from "../../../common/daSpecParser";
import { getLocalizedString } from "../../../common/localizeUtils";
import { getParserOptions } from "../../../common/openApiParserOptions";
import { isValidHttpUrl } from "../../../common/stringUtils";
import { isJsonSpecFile } from "../../../common/utils";
import { ProgrammingLanguage } from "../../../question/constants";
import { capabilityDeclarations } from "../../capabilities/declarations";
import { RegisteredStep, StepContext, StepParams } from "../../pipeline/runScaffoldPipeline";
import { defineStep } from "../../pipeline/defineStep";
import { stringParam, stringArrayParam } from "../../pipeline/stepParams";
import { withTempDirectory } from "../withTempDirectory";
import { generateTeamsAiCustomApiFiles } from "./openApiCustomApi";
import { AuthRegistration, injectOpenApiAuthActions } from "./openApiAuth";

/** Generate API plugin files through spec-parser, then copy artifacts back via `ctx.write`. */

const SOURCE = "Scaffold";

export const STEP_GENERATE_OPENAPI_PLUGIN_FILES =
  capabilityDeclarations.step.generateOpenApiPluginFiles.id;
export const STEP_GENERATE_TEAMS_AI_CUSTOM_API_FILES =
  capabilityDeclarations.step.generateTeamsAiCustomApiFiles.id;

const MANIFEST_PATH = `${AppPackageFolderName}/${ManifestTemplateFileName}`;
const AGENT_PATH = `${AppPackageFolderName}/declarativeAgent.json`;
const PLUGIN_PATH = `${AppPackageFolderName}/${DefaultPluginManifestFileName}`;
const API_SPEC_PATH = `${AppPackageFolderName}/${DefaultApiSpecFolderName}/${DefaultApiSpecYamlFileName}`;
const ORIGINAL_API_SPEC_PATH = `${API_SPEC_PATH}.original`;
const DEFAULT_ACTION_ID = "action_1";
/** Mirrors `invalidApiSpecErrorName` in the legacy openApiSpec helper, which v4 must not import. */
const INVALID_API_SPEC_ERROR_NAME = "invalid-api-spec";
const M365_AGENTS_YML = "m365agents.yml";
const M365_AGENTS_LOCAL_YML = "m365agents.local.yml";

interface TeamsAiLanguageFiles {
  appPath: string;
  handlerPath: string;
}

function systemError(name: string, message: string): SystemError {
  return new SystemError({ source: SOURCE, name, message });
}

function readRequired(ctx: StepContext, filePath: string): Result<Buffer, FxError> {
  const current = ctx.read(filePath);
  if (current === undefined) {
    return err(
      systemError(
        "OpenApiGeneratedBaseFileMissing",
        `Cannot generate OpenAPI plugin files because '${filePath}' was not produced by the render phase.`
      )
    );
  }
  return ok(current);
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function registrationIdEnvName(authName: string): string {
  return Utils.getSafeRegistrationIdEnvName(`${authName}_${ConstantString.RegistrationIdPostfix}`);
}

function authType(operation: ListAPIInfo): "apiKey" | "oauth2" | undefined {
  if (!operation.auth) {
    return undefined;
  }
  if (
    Utils.isBearerTokenAuth(operation.auth.authScheme) ||
    Utils.isAPIKeyAuthButNotInCookie(operation.auth.authScheme)
  ) {
    return "apiKey";
  }
  if (Utils.isOAuthWithAuthCodeFlow(operation.auth.authScheme)) {
    return "oauth2";
  }
  return undefined;
}

function selectedAuthRegistrations(
  selectedOperations: ListAPIInfo[]
): Result<AuthRegistration[], FxError> {
  const serverUrls = new Set<string>();
  const authNames = new Set<string>();
  const registrations: AuthRegistration[] = [];
  for (const operation of selectedOperations) {
    const type = authType(operation);
    const authName = operation.auth?.name;
    if (type === undefined || !authName) {
      continue;
    }
    if (operation.server) {
      serverUrls.add(operation.server);
    }
    if (authNames.has(authName)) {
      continue;
    }
    authNames.add(authName);
    registrations.push({
      authName,
      authType: type,
      registrationIdEnvName: registrationIdEnvName(authName),
    });
  }
  if (serverUrls.size > 1) {
    return err(
      new UserError({
        source: SOURCE,
        name: "OpenApiMultipleAuthServers",
        message: `Selected authenticated operations span multiple servers: ${Array.from(serverUrls).join(", ")}.`,
      })
    );
  }
  return ok(registrations);
}

function conversationStarterText(operation: ListAPIInfo): string | undefined {
  const text = operation.summary?.trim() || operation.description?.trim();
  return text || undefined;
}

function updateAuthYml(
  ctx: StepContext,
  filePath: string,
  registrations: AuthRegistration[]
): Result<void, FxError> {
  const current = ctx.read(filePath);
  if (current === undefined) {
    return ok(undefined);
  }
  const updated = injectOpenApiAuthActions(
    current.toString("utf8"),
    registrations,
    `./${API_SPEC_PATH}`
  );
  if (updated.isErr()) return err(updated.error);
  ctx.write(filePath, Buffer.from(updated.value, "utf8"));
  return ok(undefined);
}

function languageParam(params: StepParams): ProgrammingLanguage | undefined {
  const language = stringParam(params, "language");
  switch (language) {
    case ProgrammingLanguage.TS:
      return ProgrammingLanguage.TS;
    case ProgrammingLanguage.JS:
      return ProgrammingLanguage.JS;
    case ProgrammingLanguage.PY:
      return ProgrammingLanguage.PY;
    default:
      return undefined;
  }
}

function teamsAiLanguageFiles(language: ProgrammingLanguage): TeamsAiLanguageFiles {
  if (language === ProgrammingLanguage.TS) {
    return { appPath: "src/app/app.ts", handlerPath: "src/app/handlers.ts" };
  }
  if (language === ProgrammingLanguage.JS) {
    return { appPath: "src/app/app.js", handlerPath: "src/app/handlers.js" };
  }
  return { appPath: "src/app.py", handlerPath: "src/handlers.py" };
}

async function writeTempFile(root: string, relativePath: string, data: Buffer): Promise<void> {
  const destination = path.join(root, relativePath);
  await fs.ensureDir(path.dirname(destination));
  await fs.writeFile(destination, data);
}

async function writeTeamsAiBaseFiles(
  ctx: StepContext,
  root: string,
  languageFiles: TeamsAiLanguageFiles
): Promise<Result<void, FxError>> {
  for (const filePath of [MANIFEST_PATH, languageFiles.appPath, languageFiles.handlerPath]) {
    const current = readRequired(ctx, filePath);
    if (current.isErr()) {
      return err(current.error);
    }
    await writeTempFile(root, filePath, current.value);
  }
  return ok(undefined);
}

async function writeTempTreeToContext(root: string, ctx: StepContext): Promise<void> {
  const walk = async (dir: string): Promise<void> => {
    for (const entry of await fs.readdir(dir, { withFileTypes: true })) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        await walk(fullPath);
        continue;
      }
      const relativePath = path.relative(root, fullPath).replace(/\\/g, "/");
      ctx.write(relativePath, await fs.readFile(fullPath));
    }
  };
  await walk(root);
}

async function writeTempBaseFiles(
  root: string,
  manifest: Buffer,
  agent: Buffer
): Promise<{ manifestPath: string; pluginPath: string; apiSpecPath: string }> {
  const manifestPath = path.join(root, MANIFEST_PATH);
  const agentPath = path.join(root, AGENT_PATH);
  const pluginPath = path.join(root, PLUGIN_PATH);
  const apiSpecPath = path.join(root, API_SPEC_PATH);
  await fs.ensureDir(path.dirname(manifestPath));
  await fs.ensureDir(path.dirname(apiSpecPath));
  await fs.writeFile(manifestPath, manifest);
  await fs.writeFile(agentPath, agent);
  return { manifestPath, pluginPath, apiSpecPath };
}

async function readOriginalOpenApiSpec(apiSpecLocation: string): Promise<Buffer> {
  if (isValidHttpUrl(apiSpecLocation)) {
    const response = await axios.get<ArrayBuffer>(apiSpecLocation, { responseType: "arraybuffer" });
    return Buffer.from(response.data);
  }
  return await fs.readFile(apiSpecLocation);
}

function parseOpenApiParams(
  resolved: StepParams
): Result<{ apiSpecLocation: string; apiOperations: string[] }, string> {
  const apiSpecLocation = stringParam(resolved, "apiSpecLocation");
  if (apiSpecLocation === undefined) return err("missing string parameter 'apiSpecLocation'");
  const apiOperations = stringArrayParam(resolved, "apiOperations");
  if (apiOperations === undefined) return err("missing string[] parameter 'apiOperations'");
  return ok({ apiSpecLocation, apiOperations });
}

export const openApiGeneratePluginFiles: RegisteredStep = defineStep({
  parse: parseOpenApiParams,
  invalidParams: () =>
    systemError("OpenApiGenerateParams", "resolved parameters are not all valid"),
  async apply({ apiSpecLocation, apiOperations }, ctx): Promise<Result<void, FxError>> {
    const manifest = readRequired(ctx, MANIFEST_PATH);
    if (manifest.isErr()) {
      return err(manifest.error);
    }
    const agent = readRequired(ctx, AGENT_PATH);
    if (agent.isErr()) {
      return err(agent.error);
    }

    // Same shared gate the legacy path runs, so an incompatible document fails before generation.
    const validation = await validateOpenAPISpec(apiSpecLocation);
    if (validation.status === ValidationStatus.Error) {
      const message = getLocalizedString(
        "core.createProjectQuestion.apiSpec.multipleValidationErrors.message"
      );
      return err(new UserError(SOURCE, INVALID_API_SPEC_ERROR_NAME, message, message));
    }

    return withTempDirectory(
      "m365atk-openapi-",
      (phase, error) =>
        systemError(
          "OpenApiGenerateFailed",
          `Failed to generate OpenAPI plugin files during temporary file ${phase}: ${errorMessage(error)}`
        ),
      async (tempRoot) => {
        const temp = await writeTempBaseFiles(tempRoot, manifest.value, agent.value);
        const parser = new SpecParser(apiSpecLocation, getParserOptions(ProjectType.Copilot, true));
        const listed = await parser.list();
        const selectedOperations = listed.APIs.filter((operation) =>
          apiOperations.includes(operation.api)
        );

        // Share the v3 declarative-agent generator so the Kiota-integration flag selects the same branch.
        let generated: GenerateResult;
        try {
          generated = await generatePlugin(
            apiSpecLocation,
            temp.manifestPath,
            temp.apiSpecPath,
            temp.pluginPath,
            apiOperations,
            AdaptiveCardUpdateStrategy.CreateNew
          );
        } catch (error) {
          if (!(error instanceof SpecParserError)) {
            throw error;
          }
          return err(
            new SystemError(
              SpecParserSource,
              error.errorType.toString(),
              error.message,
              error.message
            )
          );
        }
        for (const warning of generated.warnings) {
          ctx.warn?.({ type: String(warning.type), content: warning.content });
        }

        const agentManifestPath = path.join(tempRoot, AGENT_PATH);
        let agentManifest: DeclarativeAgentManifestWrapper;
        try {
          agentManifest = DeclarativeAgentManifestWrapper.fromJSON(
            await fs.readFile(agentManifestPath, "utf8")
          );
        } catch (error) {
          return err(
            systemError(
              "JSONSyntaxError",
              `The declarative agent manifest is not valid JSON: ${errorMessage(error)}`
            )
          );
        }
        agentManifest.upsertAction(DEFAULT_ACTION_ID, DefaultPluginManifestFileName);
        const remainingStarterSlots = Math.max(0, 6 - agentManifest.conversationStarters.length);
        const seenStarterTexts = new Set(
          agentManifest.conversationStarters.map((starter) => starter.text)
        );
        const starterTexts = selectedOperations
          .map(conversationStarterText)
          .filter((text): text is string => text !== undefined)
          .filter((text) => {
            if (seenStarterTexts.has(text)) {
              return false;
            }
            seenStarterTexts.add(text);
            return true;
          })
          .slice(0, remainingStarterSlots);
        for (const text of starterTexts) {
          agentManifest.addConversationStarter(text);
        }
        await fs.writeFile(agentManifestPath, agentManifest.toJSON(), "utf8");

        const registrations = selectedAuthRegistrations(selectedOperations);
        if (registrations.isErr()) {
          return err(registrations.error);
        }

        await writeTempTreeToContext(tempRoot, ctx);
        if (ctx.read(ORIGINAL_API_SPEC_PATH) === undefined) {
          // The Kiota branch already emits the unfiltered document beside the generated spec.
          ctx.write(ORIGINAL_API_SPEC_PATH, await readOriginalOpenApiSpec(apiSpecLocation));
        }
        for (const filePath of [M365_AGENTS_YML, M365_AGENTS_LOCAL_YML]) {
          const updated = updateAuthYml(ctx, filePath, registrations.value);
          if (updated.isErr()) return err(updated.error);
        }
        return ok(undefined);
      }
    );
  },
});

export const openApiGenerateTeamsAiCustomApiFiles: RegisteredStep = defineStep({
  parse(resolved) {
    const params = parseOpenApiParams(resolved);
    if (params.isErr()) return err(params.error);
    const language = languageParam(resolved);
    if (language === undefined) return err("missing supported language parameter 'language'");
    return ok({ ...params.value, language });
  },
  invalidParams: () => systemError("OpenApiTeamsAiParams", "resolved parameters are not all valid"),
  async apply({ apiSpecLocation, apiOperations, language }, ctx): Promise<Result<void, FxError>> {
    return withTempDirectory(
      "m365atk-openapi-teams-ai-",
      (phase, error) =>
        systemError(
          "OpenApiTeamsAiGenerateFailed",
          `Failed to generate Teams AI custom API files during temporary file ${phase}: ${errorMessage(
            error
          )}`
        ),
      async (tempRoot) => {
        const languageFiles = teamsAiLanguageFiles(language);
        const baseFiles = await writeTeamsAiBaseFiles(ctx, tempRoot, languageFiles);
        if (baseFiles.isErr()) {
          return err(baseFiles.error);
        }

        const openapiSpecFileName = (await isJsonSpecFile(apiSpecLocation))
          ? DefaultApiSpecJsonFileName
          : DefaultApiSpecYamlFileName;
        const apiSpecPath = `${AppPackageFolderName}/${DefaultApiSpecFolderName}/${openapiSpecFileName}`;
        const tempApiSpecPath = path.join(tempRoot, apiSpecPath);
        await fs.ensureDir(path.dirname(tempApiSpecPath));

        const parser = new SpecParser(apiSpecLocation, getParserOptions(ProjectType.TeamsAi));
        const validation = await parser.validate();
        if (validation.status === ValidationStatus.Error) {
          return err(
            new UserError({
              source: SOURCE,
              name: "OpenApiSpecInvalid",
              message:
                "The OpenAPI description document is invalid or contains no supported operations.",
            })
          );
        }

        await parser.generate(
          path.join(tempRoot, MANIFEST_PATH),
          apiOperations,
          tempApiSpecPath,
          undefined
        );
        const specs = await parser.getFilteredSpecs(apiOperations);
        const filteredSpec = specs[1];
        if (filteredSpec === undefined) {
          return err(
            systemError(
              "OpenApiTeamsAiFilteredSpecMissing",
              "Failed to generate the filtered OpenAPI document for the selected operations."
            )
          );
        }
        const warnings = await generateTeamsAiCustomApiFiles(
          filteredSpec,
          language,
          tempRoot,
          openapiSpecFileName
        );
        for (const warning of warnings) {
          ctx.warn?.({ type: String(warning.type), content: warning.content });
        }

        await writeTempTreeToContext(tempRoot, ctx);
        return ok(undefined);
      }
    );
  },
});
