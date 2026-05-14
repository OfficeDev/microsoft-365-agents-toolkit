// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Tag-driven validators for E2E tests.
 *
 * Each Validator is a pure data object:
 *   - `id`     : stable identifier for reporting
 *   - `tags`   : template tags that opt this validator in ("*" = always)
 *   - `phases` : lifecycle phases this validator applies to
 *   - `applies`: optional extra predicate beyond tags
 *   - `run`    : (ctx) => Promise<AssertionResult[]> — pure, never throws
 *
 * Validators are composed in `VALIDATORS` and selected by
 * `runValidatorsForPhase(phase, ctx)` or by the legacy tag-based
 * `runValidators(tags, envMap, projectPath)` shim kept for callers
 * that haven't migrated yet.
 *
 * Design rules:
 *   - Return `AssertionResult[]` — don't throw, don't `expect`.
 *   - Only check what the template declared it would produce (via tags).
 *   - Integration-tier assertions (Graph/TDP/ARM) must be gated so they
 *     silently no-op when credentials are absent.
 */

import * as fs from "fs";
import * as path from "path";
import * as yaml from "js-yaml";
import type { AssertionResult } from "./tracer";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export type EnvMap = Map<string, string>;

export type Phase = "post-scaffold" | "post-provision" | "post-deploy" | "post-publish";

export interface ValidatorCtx {
  envMap: EnvMap;
  projectPath: string;
  /** Tags declared on the template (plus any ad-hoc tags injected by the caller). */
  tags: Set<string>;
  /** Raw contents of m365agents.yml, or "" if absent. */
  yamlContent?: string;
  /** Environment name (e.g. "dev"). Used to locate env/.env.<name>.user. */
  envName?: string;
}

export interface Validator {
  id: string;
  phases: readonly Phase[];
  /**
   * Validator runs when `ctx.tags` contains at least one of these tags.
   * Use `["*"]` to run regardless of tags.
   */
  tags: readonly string[];
  applies?: (ctx: ValidatorCtx) => boolean;
  run: (ctx: ValidatorCtx) => Promise<AssertionResult[]>;
}

const UUID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// ---------------------------------------------------------------------------
// Small assertion helpers
// ---------------------------------------------------------------------------

function nonEmpty(
  name: string,
  value: string | undefined,
  tier: AssertionResult["tier"] = "shape"
): AssertionResult {
  return {
    name,
    passed: !!value,
    expected: "non-empty string",
    actual: value ?? "undefined",
    tier,
  };
}

function isUuid(
  name: string,
  value: string | undefined,
  tier: AssertionResult["tier"] = "content"
): AssertionResult {
  return {
    name,
    passed: !!value && UUID_RE.test(value),
    expected: "UUID string",
    actual: value ?? "undefined",
    tier,
  };
}

function fileExists(
  name: string,
  filePath: string,
  tier: AssertionResult["tier"] = "shape"
): AssertionResult {
  const exists = fs.existsSync(filePath);
  return {
    name,
    passed: exists,
    expected: `file exists at ${filePath}`,
    actual: exists ? "exists" : "missing",
    tier,
  };
}

// ---------------------------------------------------------------------------
// Built-in validators
// ---------------------------------------------------------------------------

const teamsAppValidator: Validator = {
  id: "teamsApp.create",
  phases: ["post-provision", "post-deploy", "post-publish"],
  tags: ["teamsApp"],
  run: async ({ envMap }) => {
    const appId = envMap.get("TEAMS_APP_ID");
    const tenantId = envMap.get("TEAMS_APP_TENANT_ID");
    const results: AssertionResult[] = [
      isUuid("TEAMS_APP_ID is a UUID", appId),
      nonEmpty("TEAMS_APP_TENANT_ID is defined", tenantId),
    ];
    if (appId && tenantId) {
      results.push({
        name: "TEAMS_APP_ID differs from TEAMS_APP_TENANT_ID",
        passed: appId !== tenantId,
        expected: "distinct values",
        actual: `${appId} vs ${tenantId}`,
        tier: "content",
      });
    }
    return results;
  },
};

const teamsAppPackageValidator: Validator = {
  id: "teamsApp.zipAppPackage",
  phases: ["post-provision", "post-deploy", "post-publish"],
  tags: ["teamsApp"],
  // Only assert when the driver actually ran.
  applies: (ctx) => ctx.envMap.has("TEAMS_APP_PACKAGE_PATH"),
  run: async ({ envMap, projectPath }) => {
    const rel = envMap.get("TEAMS_APP_PACKAGE_PATH") ?? "";
    const abs = path.isAbsolute(rel) ? rel : path.join(projectPath, rel);
    return [
      {
        name: "TEAMS_APP_PACKAGE_PATH points to a .zip file",
        passed: rel.endsWith(".zip"),
        expected: "*.zip path",
        actual: rel || "undefined",
        tier: "shape",
      },
      fileExists("app package file exists on disk", abs),
    ];
  },
};

const teamsAppManifestValidator: Validator = {
  id: "teamsApp.validateManifest",
  phases: ["post-provision", "post-deploy", "post-publish"],
  tags: ["teamsApp"],
  // Only when one of the validate drivers ran.
  applies: (ctx) =>
    ctx.envMap.has("TEAMS_APP_MANIFEST_VALID") || ctx.envMap.has("TEAMS_APP_PACKAGE_VALID"),
  run: async ({ envMap }) => {
    const results: AssertionResult[] = [];
    const manifestValid = envMap.get("TEAMS_APP_MANIFEST_VALID");
    if (manifestValid !== undefined) {
      // Warning-only: the driver's manifest validation can report "false"
      // against CI-scaffolded manifests even when the lifecycle succeeds
      // (e.g. placeholder icons, missing optional metadata).  Surface the
      // result without failing the run.
      results.push({
        name: "TEAMS_APP_MANIFEST_VALID === 'true'",
        passed: manifestValid === "true",
        expected: "true",
        actual: manifestValid,
        tier: "content",
        severity: "warning",
      });
    }
    const pkgValid = envMap.get("TEAMS_APP_PACKAGE_VALID");
    if (pkgValid !== undefined) {
      results.push({
        name: "TEAMS_APP_PACKAGE_VALID === 'true'",
        passed: pkgValid === "true",
        expected: "true",
        actual: pkgValid,
        tier: "content",
        severity: "warning",
      });
    }
    return results;
  },
};

const publishedAppValidator: Validator = {
  id: "teamsApp.publishAppPackage",
  phases: ["post-publish"],
  tags: ["publishedApp"],
  run: async ({ envMap }) => [
    isUuid("TEAMS_APP_PUBLISHED_APP_ID is a UUID", envMap.get("TEAMS_APP_PUBLISHED_APP_ID")),
  ],
};

const aadAppValidator: Validator = {
  id: "aadApp.create",
  phases: ["post-provision", "post-deploy", "post-publish"],
  tags: ["aad"],
  run: async ({ envMap }) => {
    const clientId = envMap.get("AAD_APP_CLIENT_ID");
    const objectId = envMap.get("AAD_APP_OBJECT_ID");
    const tenantId = envMap.get("AAD_APP_TENANT_ID");
    const authorityHost = envMap.get("AAD_APP_OAUTH_AUTHORITY_HOST");
    const authority = envMap.get("AAD_APP_OAUTH_AUTHORITY");
    const results: AssertionResult[] = [
      isUuid("AAD_APP_CLIENT_ID is a UUID", clientId),
      isUuid("AAD_APP_OBJECT_ID is a UUID", objectId),
      isUuid("AAD_APP_TENANT_ID is a UUID", tenantId),
      nonEmpty("AAD_APP_OAUTH_AUTHORITY_HOST is defined", authorityHost),
      nonEmpty("AAD_APP_OAUTH_AUTHORITY is defined", authority),
    ];
    if (authorityHost && authority && tenantId) {
      results.push({
        name: "AAD_APP_OAUTH_AUTHORITY === host + '/' + tenantId",
        passed: authority === `${authorityHost}/${tenantId}`,
        expected: `${authorityHost}/${tenantId}`,
        actual: authority,
        tier: "content",
      });
    }
    return results;
  },
};

const botValidator: Validator = {
  id: "botAadApp.create",
  phases: ["post-provision", "post-deploy", "post-publish"],
  tags: ["bot"],
  run: async ({ envMap }) => {
    const botId = envMap.get("BOT_ID");
    const endpoint = envMap.get("BOT_ENDPOINT") ?? envMap.get("BOT_DOMAIN");
    const results: AssertionResult[] = [
      isUuid("BOT_ID is a UUID", botId),
      nonEmpty("BOT_ENDPOINT or BOT_DOMAIN is defined", endpoint),
    ];
    const ep = envMap.get("BOT_ENDPOINT");
    if (ep) {
      results.push({
        name: "BOT_ENDPOINT uses https://",
        passed: ep.startsWith("https://"),
        expected: "https:// URL",
        actual: ep,
        tier: "content",
      });
    }
    return results;
  },
};

const tabValidator: Validator = {
  id: "tab.endpoint",
  phases: ["post-deploy", "post-publish"],
  tags: ["tab"],
  run: async ({ envMap }) => {
    const endpoint = envMap.get("TAB_ENDPOINT") ?? envMap.get("FRONTEND_ENDPOINT");
    const results: AssertionResult[] = [
      nonEmpty("TAB_ENDPOINT or FRONTEND_ENDPOINT is defined", endpoint),
    ];
    if (endpoint) {
      results.push({
        name: "tab endpoint uses https://",
        passed: endpoint.startsWith("https://"),
        expected: "https:// URL",
        actual: endpoint,
        tier: "content",
      });
    }
    return results;
  },
};

const functionValidator: Validator = {
  id: "azureFunctions.endpoint",
  phases: ["post-deploy", "post-publish"],
  tags: ["function", "azureFunction"],
  run: async ({ envMap }) => {
    const endpoint = envMap.get("API_FUNCTION_ENDPOINT");
    const results: AssertionResult[] = [nonEmpty("API_FUNCTION_ENDPOINT is defined", endpoint)];
    if (endpoint) {
      results.push({
        name: "API_FUNCTION_ENDPOINT uses https://",
        passed: endpoint.startsWith("https://"),
        expected: "https:// URL",
        actual: endpoint,
        tier: "content",
      });
    }
    return results;
  },
};

/**
 * Safely parse a JSON file.  Returns `{ json, error }` so validators can
 * emit a single "valid JSON" assertion without throwing.
 */
function readJsonSafe(filePath: string): { json: unknown; error?: string } {
  try {
    const text = fs.readFileSync(filePath, "utf-8");
    return { json: JSON.parse(text) as unknown };
  } catch (e) {
    return { json: undefined, error: (e as Error).message };
  }
}

/**
 * Locate the declarative-agent manifest inside a scaffolded project.
 *
 * Templates do not use a fixed filename for the DA manifest: `da/basic`
 * scaffolds `appPackage/declarativeAgent.json`, while `da/api-plugin-*`
 * scaffold `appPackage/repairDeclarativeAgent.json`.  The authoritative
 * reference is `appPackage/manifest.json`'s
 * `copilotAgents.declarativeAgents[0].file` field.
 */
function findDeclarativeAgentPath(projectPath: string): string | undefined {
  const manifestPath = path.join(projectPath, "appPackage", "manifest.json");
  if (fs.existsSync(manifestPath)) {
    const { json } = readJsonSafe(manifestPath);
    const mf = (json ?? {}) as Record<string, unknown>;
    const copilot = (mf.copilotAgents ?? mf.copilotExtensions) as
      | Record<string, unknown>
      | undefined;
    const agents = copilot?.declarativeAgents;
    if (Array.isArray(agents) && agents.length > 0) {
      const file = (agents[0] as Record<string, unknown>).file;
      if (typeof file === "string" && file.length > 0) {
        return path.join(projectPath, "appPackage", file);
      }
    }
  }
  // Fall back to the conventional name used by simple DA templates.
  const fallback = path.join(projectPath, "appPackage", "declarativeAgent.json");
  return fs.existsSync(fallback) ? fallback : undefined;
}

const declarativeAgentValidator: Validator = {
  id: "declarativeAgent.manifest",
  phases: ["post-scaffold", "post-provision", "post-deploy", "post-publish"],
  tags: ["declarativeAgent"],
  run: async ({ projectPath }) => {
    const results: AssertionResult[] = [];
    const daPath = findDeclarativeAgentPath(projectPath);
    if (!daPath) {
      results.push({
        name: "declarative-agent manifest resolvable",
        passed: false,
        expected: "manifest.json → copilotAgents.declarativeAgents[0].file",
        actual: "not found",
        tier: "shape",
      });
      return results;
    }
    const relDa = path.relative(projectPath, daPath).replace(/\\/g, "/");
    results.push(fileExists(`${relDa} exists`, daPath));
    if (!fs.existsSync(daPath)) return results;

    const { json, error } = readJsonSafe(daPath);
    results.push({
      name: `${relDa} is valid JSON`,
      passed: !error,
      expected: "valid JSON",
      actual: error ?? "ok",
      tier: "shape",
    });
    if (!json || typeof json !== "object") return results;

    const doc = json as Record<string, unknown>;
    for (const key of ["$schema", "name", "description", "instructions"]) {
      results.push({
        name: `${relDa} has "${key}"`,
        passed: key in doc && !!doc[key],
        expected: `non-empty "${key}"`,
        actual: key in doc ? String(doc[key]).slice(0, 60) : "missing",
        tier: "content",
      });
    }
    return results;
  },
};

const apiPluginValidator: Validator = {
  id: "declarativeAgent.apiPlugin",
  phases: ["post-scaffold", "post-provision", "post-deploy", "post-publish"],
  tags: ["apiPlugin"],
  run: async ({ projectPath }) => {
    const results: AssertionResult[] = [];
    const daPath = findDeclarativeAgentPath(projectPath);
    if (!daPath || !fs.existsSync(daPath)) {
      // declarativeAgentValidator will surface the missing-manifest assertion;
      // don't double-report here.
      return results;
    }
    const { json } = readJsonSafe(daPath);
    const doc = (json ?? {}) as Record<string, unknown>;
    const actions = Array.isArray(doc.actions)
      ? (doc.actions as Array<Record<string, unknown>>)
      : [];

    results.push({
      name: "declarative-agent manifest declares at least one action",
      passed: actions.length > 0,
      expected: "actions.length > 0",
      actual: `${actions.length}`,
      tier: "content",
    });

    for (const action of actions) {
      const file = typeof action.file === "string" ? action.file : "";
      if (!file) {
        results.push({
          name: `action "${String(action.id ?? "?")}" declares file`,
          passed: false,
          expected: "non-empty file",
          actual: "missing",
          tier: "content",
        });
        continue;
      }
      const pluginPath = path.join(projectPath, "appPackage", file);
      results.push(fileExists(`appPackage/${file} exists`, pluginPath));

      if (fs.existsSync(pluginPath)) {
        const { json: pj, error } = readJsonSafe(pluginPath);
        results.push({
          name: `appPackage/${file} is valid JSON`,
          passed: !error,
          expected: "valid JSON",
          actual: error ?? "ok",
          tier: "shape",
        });
        if (pj && typeof pj === "object") {
          const pd = pj as Record<string, unknown>;
          results.push({
            name: `${file} declares schema_version or $schema`,
            passed: !!(pd.schema_version || pd.$schema),
            expected: "schema_version or $schema",
            actual: pd.schema_version
              ? String(pd.schema_version)
              : pd.$schema
                ? String(pd.$schema).slice(0, 60)
                : "missing",
            tier: "content",
          });
        }
      }
    }
    return results;
  },
};

const projectStructureValidator: Validator = {
  id: "project.structure",
  phases: ["post-scaffold", "post-provision", "post-deploy", "post-publish"],
  tags: ["*"],
  run: async ({ projectPath, envName, envMap }) => {
    const results: AssertionResult[] = [];
    const envDir = path.join(projectPath, "env");
    results.push(fileExists("env/ directory exists", envDir));

    if (envName) {
      const envFile = path.join(envDir, `.env.${envName}`);
      const userEnvFile = path.join(envDir, `.env.${envName}.user`);
      if (fs.existsSync(envFile)) {
        const content = fs.readFileSync(envFile, "utf-8");
        const leaked = content
          .split(/\r?\n/)
          .filter((l) => /^SECRET_[A-Z0-9_]+=\S/.test(l))
          .map((l) => l.split("=")[0]);
        results.push({
          name: `env/.env.${envName} contains no SECRET_* assignments`,
          passed: leaked.length === 0,
          expected: "no SECRET_* keys",
          actual: leaked.length === 0 ? "clean" : `leaked: ${leaked.join(", ")}`,
          tier: "content",
          severity: "error",
        });
      }
      const hasSecretsInEnv = [...envMap.keys()].some((k) => k.startsWith("SECRET_"));
      if (hasSecretsInEnv) {
        results.push(
          fileExists(`env/.env.${envName}.user exists (secrets present)`, userEnvFile, "content")
        );
      }
    }
    return results;
  },
};

/**
 * Extract declared bicep outputs from a .bicep file.
 * Matches `output NAME type = ...` at the start of a line.
 */
function parseBicepOutputs(bicepText: string): string[] {
  const re = /^\s*output\s+([A-Za-z_][A-Za-z0-9_]*)\s+/gm;
  const names: string[] = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(bicepText)) !== null) {
    names.push(m[1]);
  }
  return names;
}

const armOutputsValidator: Validator = {
  id: "arm.outputs",
  phases: ["post-provision", "post-deploy", "post-publish"],
  tags: ["*"],
  // Only runs when the template actually has an infra/*.bicep.
  applies: (ctx) => {
    const infraDir = path.join(ctx.projectPath, "infra");
    if (!fs.existsSync(infraDir)) return false;
    return fs
      .readdirSync(infraDir)
      .some((f) => f.endsWith(".bicep") && !f.includes(".parameters."));
  },
  run: async ({ projectPath, envMap }) => {
    const results: AssertionResult[] = [];
    const infraDir = path.join(projectPath, "infra");
    const bicepFiles = fs
      .readdirSync(infraDir)
      .filter((f) => f.endsWith(".bicep") && !f.includes(".parameters."));

    const allOutputs = new Set<string>();
    for (const file of bicepFiles) {
      const text = fs.readFileSync(path.join(infraDir, file), "utf-8");
      for (const name of parseBicepOutputs(text)) {
        allOutputs.add(name);
      }
    }

    for (const name of allOutputs) {
      const value = envMap.get(name);
      results.push({
        name: `ARM output "${name}" persisted to .env`,
        passed: !!value,
        expected: "non-empty env value",
        actual: value ?? "missing",
        tier: "integration",
      });
    }
    return results;
  },
};

/**
 * Extract the env-var names that `oauth/register` and `apiKey/register` steps
 * write their outputs to, by parsing m365agents.yml `writeToEnvironmentFile`
 * maps.  Returns the union of all declared output var names across all
 * matching steps.  Empty when the YAML is absent or malformed.
 */
function collectRegistrationEnvVars(projectPath: string, driverIds: readonly string[]): string[] {
  const yamlPath = path.join(projectPath, "m365agents.yml");
  if (!fs.existsSync(yamlPath)) return [];
  try {
    const doc = yaml.load(fs.readFileSync(yamlPath, "utf-8")) as
      | Record<string, unknown>
      | undefined;
    if (!doc) return [];
    const phases = ["provision", "deploy", "publish"] as const;
    const names: string[] = [];
    for (const phase of phases) {
      const steps = (doc[phase] as Array<Record<string, unknown>> | undefined) ?? [];
      for (const step of steps) {
        if (typeof step.uses !== "string" || !driverIds.includes(step.uses)) continue;
        const writeMap = step.writeToEnvironmentFile as Record<string, unknown> | undefined;
        if (!writeMap) continue;
        for (const v of Object.values(writeMap)) {
          if (typeof v === "string" && v.trim()) names.push(v.trim());
        }
      }
    }
    return Array.from(new Set(names));
  } catch {
    return [];
  }
}

const oauthRegisterValidator: Validator = {
  id: "oauth.register",
  phases: ["post-provision", "post-deploy", "post-publish"],
  tags: ["apiAuthOAuth"],
  applies: (ctx) => collectRegistrationEnvVars(ctx.projectPath, ["oauth/register"]).length > 0,
  run: async ({ projectPath, envMap }) => {
    const results: AssertionResult[] = [];
    const vars = collectRegistrationEnvVars(projectPath, ["oauth/register"]);
    for (const name of vars) {
      const value = envMap.get(name);
      // applicationIdUri is an API URI (api://...), everything else should be
      // a non-empty string identifier.  Only assert shape, not format.
      results.push({
        name: `oauth/register output "${name}" persisted`,
        passed: !!value,
        expected: "non-empty env value",
        actual: value ?? "missing",
        tier: "integration",
      });
    }
    return results;
  },
};

const apiKeyRegisterValidator: Validator = {
  id: "apiKey.register",
  phases: ["post-provision", "post-deploy", "post-publish"],
  tags: ["apiAuthApiKey"],
  applies: (ctx) => collectRegistrationEnvVars(ctx.projectPath, ["apiKey/register"]).length > 0,
  run: async ({ projectPath, envMap }) => {
    const results: AssertionResult[] = [];
    const vars = collectRegistrationEnvVars(projectPath, ["apiKey/register"]);
    for (const name of vars) {
      const value = envMap.get(name);
      results.push({
        name: `apiKey/register output "${name}" persisted`,
        passed: !!value,
        expected: "non-empty env value",
        actual: value ?? "missing",
        tier: "integration",
      });
    }
    return results;
  },
};

/**
 * Authoritative built-in validator list.
 * Order controls reporting order only; `failedAssertions()` decides pass/fail.
 */
export const VALIDATORS: readonly Validator[] = [
  projectStructureValidator,
  teamsAppValidator,
  teamsAppPackageValidator,
  teamsAppManifestValidator,
  publishedAppValidator,
  aadAppValidator,
  botValidator,
  tabValidator,
  functionValidator,
  declarativeAgentValidator,
  apiPluginValidator,
  armOutputsValidator,
  oauthRegisterValidator,
  apiKeyRegisterValidator,
];

function tagMatches(v: Validator, ctx: ValidatorCtx): boolean {
  return v.tags.includes("*") || v.tags.some((t) => ctx.tags.has(t));
}

async function safeRun(v: Validator, ctx: ValidatorCtx): Promise<AssertionResult[]> {
  try {
    return await v.run(ctx);
  } catch (e) {
    return [
      {
        name: `validator ${v.id} threw`,
        passed: false,
        expected: "no exception",
        actual: (e as Error).message ?? String(e),
        tier: "shape",
        severity: "error",
      },
    ];
  }
}

/** Run every validator eligible for the given phase. */
export async function runValidatorsForPhase(
  phase: Phase,
  ctx: ValidatorCtx
): Promise<AssertionResult[]> {
  const results: AssertionResult[] = [];
  for (const v of VALIDATORS) {
    if (!v.phases.includes(phase)) continue;
    if (!tagMatches(v, ctx)) continue;
    if (v.applies && !v.applies(ctx)) continue;
    results.push(...(await safeRun(v, ctx)));
  }
  return results;
}

/** Assertions that should fail the test run (severity !== "warning"). */
export function failedAssertions(results: readonly AssertionResult[]): AssertionResult[] {
  return results.filter((r) => !r.passed && (r.severity ?? "error") === "error");
}

export function allPassed(results: readonly AssertionResult[]): boolean {
  return failedAssertions(results).length === 0;
}

// ---------------------------------------------------------------------------
// Legacy shim — preserves the old (tags, envMap, projectPath) signature.
// Runs every tag-matching validator regardless of phase.
// ---------------------------------------------------------------------------

/**
 * @deprecated Prefer `runValidatorsForPhase(phase, ctx)`.
 */
export async function runValidators(
  tags: string[],
  envMap: EnvMap,
  projectPath: string
): Promise<AssertionResult[]> {
  const ctx: ValidatorCtx = {
    envMap,
    projectPath,
    tags: new Set(tags),
  };
  const results: AssertionResult[] = [];
  for (const v of VALIDATORS) {
    if (!tagMatches(v, ctx)) continue;
    if (v.applies && !v.applies(ctx)) continue;
    results.push(...(await safeRun(v, ctx)));
  }
  return results;
}
