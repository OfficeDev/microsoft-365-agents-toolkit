// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FxError,
  SystemError,
  UserInteraction,
} from "@microsoft/teamsfx-api";
import {
  ListAPIInfo,
  ParseOptions,
  ProjectType,
  SpecParser,
  Utils,
  ValidationStatus,
} from "@microsoft/m365-spec-parser";
import { Result, err } from "neverthrow";
import {
  CollectInputsPort,
  OptionsProvider,
  OptionsSchema,
  Validator,
  collectInputs,
} from "../collectInputs/collectInputs";
import { openCreateQuestions } from "../distribution/createQuestions";
import { openDeclarativePackage } from "../distribution/declarativePackage";
import { evaluateExpression } from "../expression/evaluateExpression";
import { Answers, DeclarativeLocator } from "../model/dataModel";
import { parseDeclaredKeys } from "../runtime/packageParse";
import { createExpressionPort } from "../runtime/whitelist";
import { createUiPromptUI } from "./uiPromptUI";

/**
 * The live surface wiring of `collect-inputs` for the create path: run one
 * engine-decided template's Q2 (`questions.json`) over the host
 * `UserInteraction`, producing the v4 `Answers`. This is the half of the
 * front-loaded create funnel that comes after `resolve-build-target` /
 * `route-declarative-via-selector` pick the `templateId` + the v4 engine — the
 * v4 template asks its own follow-up questions through the v4 engine, never the
 * v3 question tree.
 *
 * Spec: docs/03-specs/operations/scaffolding/collect-create-inputs.md
 * Scenario: docs/03-specs/scenarios/da/create-mcp-server.md
 *
 * v4-owned (INV-7): imports no v3 symbol. `UserInteraction` is
 * `@microsoft/teamsfx-api` (upstream of both worlds). The feature-flag reader is
 * injected (INV-4) — v4 imports no `featureFlagManager`; the floor read is
 * injected by the caller (INV-5), so the operation is CI-testable from an
 * in-memory floor with no built artifact.
 */

/**
 * The default `mcp.serverTypes` provider: the **remote** server type only. Live
 * local-server detection (the `odr` external dependency) rides an injected
 * provider in a later increment, so the default stays CI-testable with no
 * external process (spec Boundary).
 */
const remoteOnlyServerTypes: OptionsProvider = {
  fetch() {
    return { options: [{ id: "remote", label: "Remote" }] };
  },
};

const openApiMethods = [
  "get",
  "post",
  "put",
  "delete",
  "patch",
  "head",
  "connect",
  "options",
  "trace",
];

function openApiParseOptions(): ParseOptions {
  return {
    isGptPlugin: true,
    allowAPIKeyAuth: true,
    allowBearerTokenAuth: true,
    allowMultipleParameters: true,
    allowOauth2: true,
    projectType: ProjectType.Copilot,
    allowMissingId: true,
    allowSwagger: true,
    allowMethods: openApiMethods,
    allowResponseSemantics: true,
    allowConversationStarters: false,
    allowConfirmation: false,
  };
}

function operationDetail(operation: ListAPIInfo): string {
  if (!operation.auth) {
    return "No authentication";
  }
  if (Utils.isBearerTokenAuth(operation.auth.authScheme)) {
    return "API key";
  }
  if (Utils.isOAuthWithAuthCodeFlow(operation.auth.authScheme)) {
    return "OAuth";
  }
  if (Utils.isAPIKeyAuthButNotInCookie(operation.auth.authScheme)) {
    return "API key with header or query parameter";
  }
  return "Unsupported authentication";
}

function sortOperations(operations: ListAPIInfo[]): ListAPIInfo[] {
  return [...operations].sort((left, right) => {
    const leftParts = left.api.toLowerCase().split(" ");
    const rightParts = right.api.toLowerCase().split(" ");
    if (leftParts[0] < rightParts[0]) {
      return -1;
    }
    if (leftParts[0] > rightParts[0]) {
      return 1;
    }
    return (leftParts[1] ?? "").localeCompare(rightParts[1] ?? "");
  });
}

const openApiOperationsProvider: OptionsProvider = {
  async fetch(params) {
    const apiSpecLocation = params.apiSpecLocation?.trim();
    if (!apiSpecLocation) {
      throw new SystemError({
        source: "Scaffold",
        name: "OpenApiMissingSpecLocation",
        message: "OpenAPI operations cannot be listed without an API spec location.",
      });
    }
    const parser = new SpecParser(apiSpecLocation, openApiParseOptions());
    const validation = await parser.validate();
    if (validation.status === ValidationStatus.Error) {
      throw new SystemError({
        source: "Scaffold",
        name: "OpenApiSpecInvalid",
        message: "The OpenAPI description document is invalid or contains no supported operations.",
      });
    }
    const listed = await parser.list();
    return {
      options: sortOperations(listed.APIs)
        .filter((operation) => operation.isValid)
        .map((operation) => ({
          id: operation.api,
          label: operation.api,
          groupName: operation.api.toUpperCase().split(" ")[0],
          detail: operationDetail(operation),
        })),
    };
  },
};

/** The default `optionsFrom` provider registry (spec Boundary). */
const defaultProviders: Record<string, OptionsProvider> = {
  "mcp.serverTypes": remoteOnlyServerTypes,
  "openapi.operations": openApiOperationsProvider,
};

/** The `"uri"` validator: a value that does not parse as a URL is user-fixable. */
const uriValidator: Validator = (value: string): string | undefined => {
  try {
    new URL(value);
    return undefined;
  } catch {
    return "must be a valid URI";
  }
};

/** The engine-registered validator registry (spec §6.4). */
const validators: Record<string, Validator> = {
  uri: uriValidator,
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** The descriptor's language axis (decision 5), falling back to `["common"]`. */
function descriptorLanguages(descriptor: unknown): string[] {
  if (isRecord(descriptor) && Array.isArray(descriptor.languages)) {
    const languages = descriptor.languages.filter(
      (language): language is string => typeof language === "string"
    );
    if (languages.length > 0) {
      return languages;
    }
  }
  return ["common"];
}

/**
 * The CLI-only .NET gate — a v4-local copy of v3's `FeatureFlags.CLIDotNet` name
 * (INV-7: v4 imports no `featureFlagManager`; the flag reader is injected). v3
 * surfaces the `csharp` language only when this flag is on (the CLI
 * `listTemplates` / `create` gate it on the same flag) and never in the VS Code
 * extension (whose template metadata carries no `csharp`).
 */
const CLI_DOTNET_FLAG = "TEAMSFX_CLI_DOTNET";

/** The C# language id (v3 `ProgrammingLanguage.CSharp`) — surface-gated below. */
const CSHARP_LANGUAGE = "csharp";

/** The default env-backed feature-flag reader (a flag is on iff its env var is exactly `"true"`). */
function envFlagReader(name: string): boolean {
  return process.env[name] === "true";
}

/**
 * Gate the engine-owned language axis (decision 5) by surface + the .NET flag,
 * mirroring v3's platform gating: `csharp` is dropped for the VS Code extension
 * (`surface === "vscode"`), and kept for the CLI / VS surfaces only when
 * `TEAMSFX_CLI_DOTNET` is on. Every other language passes through unchanged, order
 * preserved. When the gate removes the only language the axis is simply not asked
 * — the surface offers no .NET variant of that template.
 */
export function gateLanguagesBySurface(
  languages: string[],
  surface: string,
  flagReader: (name: string) => boolean
): string[] {
  const allowCsharp = surface !== "vscode" && flagReader(CLI_DOTNET_FLAG);
  return allowCsharp ? languages : languages.filter((language) => language !== CSHARP_LANGUAGE);
}

/** The Q2 options schema: the declared identifier domain (`optionsSchema.properties` ids). */
function declaredOptionsSchema(descriptor: unknown): OptionsSchema {
  const properties: Record<string, unknown> = {};
  for (const key of parseDeclaredKeys(descriptor)) {
    properties[key] = {};
  }
  return { properties };
}

/** Injected overrides: a provider registry (live `mcp.*`) and the feature-flag reader. */
export interface CreateInputsDeps {
  /** Override `optionsFrom` providers (e.g. a live `mcp.serverTypes`); merged over the defaults. */
  optionsProvider?: Record<string, OptionsProvider>;
  /** The feature-flag reader behind `featureFlag('…')` (default: env-backed). */
  flagReader?: (name: string) => boolean;
  /** The host surface (`vscode` / `cli` / `vs`) — gates the `csharp` language axis (default `vscode`). */
  surface?: string;
}

/**
 * Run one create template's Q2 over the host surface.
 *
 * @param floorBytes  the bundled-floor channel zip (injected; CI-testable)
 * @param locator     the engine-decided create target (`{ kind, templateId }`)
 * @param entryParams pre-filled answers used as-is, never prompted (INPUT-12)
 * @param ui          the host surface (`@microsoft/teamsfx-api`)
 * @param deps        injected provider / feature-flag overrides
 * @returns `ok(Answers)`, or a `UserError` (validation failure) / `SystemError`
 *          (a missing `questions.json` / `descriptor.json`, an unknown
 *          provider / validator).
 */
export async function runCreateInputs(
  floorBytes: Buffer,
  locator: DeclarativeLocator,
  entryParams: Answers,
  ui: UserInteraction,
  deps: CreateInputsDeps = {}
): Promise<Result<Answers, FxError>> {
  const questions = openCreateQuestions(floorBytes, locator);
  if (questions.isErr()) {
    return err(questions.error);
  }

  const opened = openDeclarativePackage(floorBytes, locator);
  if (opened.isErr()) {
    return err(opened.error);
  }
  const descriptor = opened.value.descriptor;
  const languages = gateLanguagesBySurface(
    descriptorLanguages(descriptor),
    deps.surface ?? "vscode",
    deps.flagReader ?? envFlagReader
  );

  const providers = { ...defaultProviders, ...(deps.optionsProvider ?? {}) };
  const expressionPort = createExpressionPort(deps.flagReader);
  const port: CollectInputsPort = {
    ui: createUiPromptUI(ui),
    optionsProvider: (providerId) => providers[providerId],
    validator: (name) => validators[name],
    evaluate: (node, scope) => evaluateExpression(node, scope, expressionPort),
  };

  return collectInputs(
    questions.value,
    declaredOptionsSchema(descriptor),
    entryParams,
    languages,
    port
  );
}
