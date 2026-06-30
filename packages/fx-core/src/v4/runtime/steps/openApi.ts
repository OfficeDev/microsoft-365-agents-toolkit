// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ConstantString,
  ListAPIInfo,
  ProjectType,
  SpecParser,
  Utils,
} from "@microsoft/m365-spec-parser";
import {
  AppPackageFolderName,
  DefaultApiSpecFolderName,
  DefaultApiSpecYamlFileName,
  DefaultPluginManifestFileName,
  FxError,
  ManifestTemplateFileName,
  SystemError,
  UserError,
} from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import * as os from "os";
import * as path from "path";
import { Result, err, ok } from "neverthrow";
import { RegisteredStep, StepContext, StepParams } from "../../pipeline/runScaffoldPipeline";

/** Generate API plugin files through spec-parser, then copy artifacts back via `ctx.write`. */

const SOURCE = "Scaffold";

export const STEP_GENERATE_OPENAPI_PLUGIN_FILES = "openapi/generate-plugin-files";

const MANIFEST_PATH = `${AppPackageFolderName}/${ManifestTemplateFileName}`;
const AGENT_PATH = `${AppPackageFolderName}/declarativeAgent.json`;
const PLUGIN_PATH = `${AppPackageFolderName}/${DefaultPluginManifestFileName}`;
const API_SPEC_PATH = `${AppPackageFolderName}/${DefaultApiSpecFolderName}/${DefaultApiSpecYamlFileName}`;
const DEFAULT_ACTION_ID = "action_1";
const M365_AGENTS_YML = "m365agents.yml";
const M365_AGENTS_LOCAL_YML = "m365agents.local.yml";

interface AuthRegistration {
  authName: string;
  authType: "apiKey" | "oauth2";
  registrationIdEnvName: string;
}

interface ConversationStarter {
  text: string;
}

function systemError(name: string, message: string): SystemError {
  return new SystemError({ source: SOURCE, name, message });
}

function stringParam(params: StepParams, key: string): string | undefined {
  const value = params[key];
  return typeof value === "string" ? value : undefined;
}

function stringArrayParam(params: StepParams, key: string): string[] | undefined {
  const value = params[key];
  if (!Array.isArray(value) || !value.every((item) => typeof item === "string")) {
    return undefined;
  }
  return value;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
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

function parseJsonObject(
  json: string,
  filePath: string,
  errorName: string
): Result<Record<string, unknown>, FxError> {
  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch (error) {
    return err(
      systemError(`${errorName}Parse`, `Cannot parse '${filePath}' as JSON: ${errorMessage(error)}`)
    );
  }
  if (!isRecord(parsed)) {
    return err(systemError(`${errorName}Shape`, `'${filePath}' is not a JSON object`));
  }
  return ok(parsed);
}

function isConversationStarter(value: unknown): value is ConversationStarter {
  return isRecord(value) && typeof value.text === "string";
}

function conversationStartersFromUnknown(value: unknown): ConversationStarter[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value.filter(isConversationStarter).map((starter) => ({ text: starter.text }));
}

function pluginConversationStarters(plugin: Record<string, unknown>): ConversationStarter[] {
  const capabilities = plugin.capabilities;
  if (!isRecord(capabilities)) {
    return [];
  }
  return conversationStartersFromUnknown(capabilities.conversation_starters);
}

function pluginOperationIds(plugin: Record<string, unknown>): Set<string> {
  const operationIds = new Set<string>();
  if (!Array.isArray(plugin.functions)) {
    return operationIds;
  }
  for (const func of plugin.functions) {
    if (!isRecord(func) || typeof func.name !== "string") {
      continue;
    }
    operationIds.add(func.name);
  }
  return operationIds;
}

async function specConversationStarters(
  apiSpecPath: string,
  operationIds: Set<string>
): Promise<ConversationStarter[]> {
  if (operationIds.size === 0) {
    return [];
  }
  const parser = new SpecParser(apiSpecPath, openApiParseOptions());
  const listed = await parser.list();
  const starters: ConversationStarter[] = [];
  for (const operation of listed.APIs) {
    if (!operation.isValid || !operation.operationId || !operationIds.has(operation.operationId)) {
      continue;
    }
    const text = operation.summary || operation.description;
    if (text) {
      starters.push({ text });
    }
  }
  return starters;
}

function appendConversationStarters(
  agent: Record<string, unknown>,
  starters: ConversationStarter[]
): void {
  if (starters.length === 0) {
    return;
  }

  const nextStarters = Array.isArray(agent.conversation_starters)
    ? [...agent.conversation_starters]
    : [];
  for (const starter of starters) {
    if (nextStarters.length >= 6) {
      break;
    }
    const alreadyExists = nextStarters.some(
      (existingStarter) =>
        isConversationStarter(existingStarter) && existingStarter.text === starter.text
    );
    if (!alreadyExists) {
      nextStarters.push(starter);
    }
  }
  agent.conversation_starters = nextStarters;
}

async function addActionAndConversationStarters(
  agentJson: string,
  pluginJson: string,
  apiSpecPath: string
): Promise<Result<string, FxError>> {
  const agent = parseJsonObject(agentJson, AGENT_PATH, "OpenApiDeclarativeAgent");
  if (agent.isErr()) {
    return err(agent.error);
  }
  const plugin = parseJsonObject(pluginJson, PLUGIN_PATH, "OpenApiPluginManifest");
  if (plugin.isErr()) {
    return err(plugin.error);
  }

  const actions = agent.value.actions;
  if (actions !== undefined && !Array.isArray(actions)) {
    return err(
      systemError("OpenApiDeclarativeAgentActions", `'${AGENT_PATH}' actions is not an array`)
    );
  }

  const nextActions = Array.isArray(actions) ? [...actions] : [];
  nextActions.push({ id: DEFAULT_ACTION_ID, file: DefaultPluginManifestFileName });
  agent.value.actions = nextActions;

  let starters = pluginConversationStarters(plugin.value);
  if (starters.length === 0) {
    starters = await specConversationStarters(apiSpecPath, pluginOperationIds(plugin.value));
  }
  appendConversationStarters(agent.value, starters);

  return ok(JSON.stringify(agent.value, null, 4) + "\n");
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

async function selectedAuthRegistrations(
  apiSpecLocation: string,
  apiOperations: string[]
): Promise<Result<AuthRegistration[], FxError>> {
  const parser = new SpecParser(apiSpecLocation, openApiParseOptions());
  const listed = await parser.list();
  const selected = listed.APIs.filter((operation) => apiOperations.includes(operation.api));
  const serverUrls = new Set<string>();
  const authNames = new Set<string>();
  const registrations: AuthRegistration[] = [];
  for (const operation of selected) {
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

function authActionBlock(registration: AuthRegistration): string {
  if (registration.authType === "apiKey") {
    return [
      "  # Register API KEY",
      "  - uses: apiKey/register",
      "    with:",
      "      # Name of the API Key",
      `      name: ${registration.authName}`,
      "      # app ID",
      "      appId: ${{TEAMS_APP_ID}}",
      "      # Path to OpenAPI description document",
      `      apiSpecPath: ./${API_SPEC_PATH}`,
      "    # Write the registration information of API Key into environment file for",
      "    # the specified environment variable(s).",
      "    writeToEnvironmentFile:",
      `      registrationId: ${registration.registrationIdEnvName}`,
    ].join("\n");
  }
  return [
    "  - uses: oauth/register",
    "    with:",
    `      name: ${registration.authName}`,
    "      flow: authorizationCode",
    "      # app ID",
    "      appId: ${{TEAMS_APP_ID}}",
    "      # Path to OpenAPI description document",
    `      apiSpecPath: ./${API_SPEC_PATH}`,
    "      # Use below property to change token exchange behaviour, BasicAuthorizationHeader: token exchange is done via HTTP headers. PostRequestBody: token exchange is done via request body",
    "      # tokenExchangeMethodType: BasicAuthorizationHeader",
    "      # Uncomment below property to use proof key for code exchange (PKCE)",
    "      # isPKCEEnabled: true",
    "    writeToEnvironmentFile:",
    `      configurationId: ${registration.registrationIdEnvName}`,
  ].join("\n");
}

function injectAuthActions(yml: string, registrations: AuthRegistration[]): string {
  if (registrations.length === 0) {
    return yml;
  }
  const marker = "  # Build app package with latest env value";
  const block = registrations.map(authActionBlock).join("\n\n") + "\n\n";
  const index = yml.indexOf(marker);
  if (index === -1) {
    return yml + (yml.endsWith("\n") ? "" : "\n") + block;
  }
  return yml.slice(0, index) + block + yml.slice(index);
}

function updateAuthYml(
  ctx: StepContext,
  filePath: string,
  registrations: AuthRegistration[]
): void {
  const current = ctx.read(filePath);
  if (current === undefined) {
    return;
  }
  const updated = injectAuthActions(current.toString("utf8"), registrations);
  ctx.write(filePath, Buffer.from(updated, "utf8"));
}

function openApiParseOptions() {
  return {
    isGptPlugin: true,
    allowAPIKeyAuth: true,
    allowBearerTokenAuth: true,
    allowMultipleParameters: true,
    allowOauth2: true,
    projectType: ProjectType.Copilot,
    allowMissingId: true,
    allowSwagger: true,
    allowMethods: ["get", "post", "put", "delete", "patch", "head", "connect", "options", "trace"],
    allowResponseSemantics: true,
    allowConversationStarters: false,
    allowConfirmation: false,
  };
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

export const openApiGeneratePluginFiles: RegisteredStep = {
  validateParams(resolved: StepParams): string | undefined {
    if (stringParam(resolved, "apiSpecLocation") === undefined) {
      return "missing string parameter 'apiSpecLocation'";
    }
    if (stringArrayParam(resolved, "apiOperations") === undefined) {
      return "missing string[] parameter 'apiOperations'";
    }
    return undefined;
  },

  async apply(resolved: StepParams, ctx: StepContext): Promise<Result<void, FxError>> {
    const apiSpecLocation = stringParam(resolved, "apiSpecLocation");
    const apiOperations = stringArrayParam(resolved, "apiOperations");
    if (apiSpecLocation === undefined || apiOperations === undefined) {
      return err(systemError("OpenApiGenerateParams", "resolved parameters are not all valid"));
    }

    const manifest = readRequired(ctx, MANIFEST_PATH);
    if (manifest.isErr()) {
      return err(manifest.error);
    }
    const agent = readRequired(ctx, AGENT_PATH);
    if (agent.isErr()) {
      return err(agent.error);
    }

    const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "m365atk-openapi-"));
    try {
      const temp = await writeTempBaseFiles(tempRoot, manifest.value, agent.value);
      const parser = new SpecParser(apiSpecLocation, openApiParseOptions());

      await parser.generateForCopilot(
        temp.manifestPath,
        apiOperations,
        temp.apiSpecPath,
        temp.pluginPath
      );

      const updatedAgent = await addActionAndConversationStarters(
        (await fs.readFile(path.join(tempRoot, AGENT_PATH))).toString("utf8"),
        (await fs.readFile(temp.pluginPath)).toString("utf8"),
        temp.apiSpecPath
      );
      if (updatedAgent.isErr()) {
        return err(updatedAgent.error);
      }

      ctx.write(MANIFEST_PATH, await fs.readFile(temp.manifestPath));
      ctx.write(AGENT_PATH, Buffer.from(updatedAgent.value, "utf8"));
      ctx.write(PLUGIN_PATH, await fs.readFile(temp.pluginPath));
      ctx.write(API_SPEC_PATH, await fs.readFile(temp.apiSpecPath));
      const registrations = await selectedAuthRegistrations(apiSpecLocation, apiOperations);
      if (registrations.isErr()) {
        return err(registrations.error);
      }
      updateAuthYml(ctx, M365_AGENTS_YML, registrations.value);
      updateAuthYml(ctx, M365_AGENTS_LOCAL_YML, registrations.value);
      return ok(undefined);
    } catch (error) {
      return err(
        systemError(
          "OpenApiGenerateFailed",
          `Failed to generate OpenAPI plugin files: ${errorMessage(error)}`
        )
      );
    } finally {
      await fs.remove(tempRoot);
    }
  },
};
