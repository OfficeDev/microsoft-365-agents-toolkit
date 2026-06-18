// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, UserError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";

/**
 * The v4 template-package validation gate — is one `templates-v4@<version>`
 * package well-formed, and may it run on **this** engine?
 *
 * Spec: docs/03-specs/operations/scaffolding/validate-template-package.md
 * Decision: docs/02-architecture/adr/ADR-0015-templates-version-artifact-shape.md
 *           (placeholder closure shares ADR-0016 invariant 5; routing consistency
 *           shares ADR-0014 §5.3 descriptor-derived routing)
 *
 * This is part of the v4 world; it imports no v3 symbol (spec INV-7) and is a
 * pure function of its inputs (spec INV-8 / AC-21): no fs / http / clock. The
 * same checks run at build CI (a violation fails the build) and at engine load
 * (defense-in-depth); only the reverse `minEngineVersion` comparison is
 * `load`-only (spec §"mode", flow diagram).
 */

const SOURCE = "Scaffold";

/** Which per-kind `selector.json` + `templateId` namespace the package belongs to. */
export type PackageKind = "create" | "modify";

/**
 * `build` → a violation fails the build (before ship); `load` → a violation
 * fails the scaffold (defense-in-depth). The *checks* are identical; the one
 * behavioral difference is that the reverse `engineVersion >= minEngineVersion`
 * comparison runs only in `load` (a live engine exists to compare against).
 */
export type ValidateMode = "build" | "load";

/**
 * One content file's path plus the `{{token}}` set extracted from it. The
 * port's `content()` returns these; emptiness of `content/` is expressed by
 * `content()` returning `undefined`, never a marker file (spec INV-1 / AC-06).
 */
export interface ContentFile {
  path: string;
  placeholders: string[];
}

/**
 * A JSON-schema validator face: `undefined` = valid, a string = the failing
 * rule (named back to the author). Keeps ajv behind the port so this operation
 * stays pure (interface-segregation; the full runtime composes the real ajv).
 */
export type SchemaValidator = (data: unknown) => string | undefined;

/**
 * The narrow port `validate-template-package` actually uses (spec §"Inputs":
 * "declares the narrow `TemplatePackagePort` it actually uses"). The first
 * seven faces are the spec's per-package faces; `callerFloor` and
 * `presentTemplateIds` are the two artifact/engine-context faces the
 * placeholder-closure (AC-11 "caller-injected identifier") and routing-
 * consistency (AC-13/14, INV-4) criteria provably require — the full runtime
 * composes them, an in-memory fake supplies them in tests. Surfacing them as
 * port faces keeps this pure v4 operation free of hardcoded surface knowledge
 * (the caller floor's surface-gated members are not yet finalized).
 */
export interface TemplatePackagePort {
  /** The package's parsed `descriptor.json`, or `undefined` when absent. */
  descriptor(): unknown | undefined;
  /** The package's parsed `questions.json`, or `undefined` when absent. */
  questions(): unknown | undefined;
  /** The package's parsed `pipeline.json`, or `undefined` when absent. */
  pipeline(): unknown | undefined;
  /** Each content file's path + `{{token}}` set, or `undefined` when `content/` is absent. */
  content(): ContentFile[] | undefined;
  /** The per-kind `selector.json` (parsed). */
  selector(kind: PackageKind): unknown;
  /** The JSON-schema validators under `templates/v4/schema/`. */
  schemas: { descriptor: SchemaValidator; question: SchemaValidator; selector: SchemaValidator };
  /** The consuming engine's SemVer (the `load`-mode reverse gate). */
  engineVersion(): string;
  /** The closed caller-injected identifier names (`appName`, the `language` axis, …). */
  callerFloor(): string[];
  /** The `templateId`s whose `descriptor.json` is present in the artifact, per kind. */
  presentTemplateIds(kind: PackageKind): string[];
}

/** The validated package outcome (spec §"Outputs"). */
export interface ValidatedPackage {
  /** The parsed, schema-valid descriptor. */
  descriptor: Record<string, unknown>;
  /** The resolved reverse-gate floor (recorded on outcome / telemetry). */
  minEngineVersion: string;
  /** The validated content-file list (empty when `content/` is absent). */
  contentFiles: ContentFile[];
}

/** `UserError` name: a required package file (`descriptor`/`questions`/`pipeline`) is absent. */
export const VALIDATE_REQUIRED_FILE = "TemplatePackageRequiredFile";
/** `UserError` name: a package file failed its JSON schema. */
export const VALIDATE_SCHEMA = "TemplatePackageSchema";
/** `UserError` name: a content token has no producer, or a required var has no consumer. */
export const VALIDATE_PLACEHOLDER_DRIFT = "TemplatePackagePlaceholderDrift";
/** `UserError` name: a v4 selector route names a `templateId` with no present descriptor. */
export const VALIDATE_DANGLING_ROUTE = "TemplatePackageDanglingRoute";
/** `UserError` name: the same `templateId` is routed in both the create and modify selectors. */
export const VALIDATE_KIND_OVERLAP = "TemplatePackageKindOverlap";
/** `UserError` name: `descriptor.minEngineVersion` is missing (it is mandatory). */
export const VALIDATE_MIN_ENGINE_MISSING = "TemplatePackageMinEngineVersionMissing";
/** `UserError` name: `engineVersion < minEngineVersion` — the engine is too old. */
export const VALIDATE_ENGINE_TOO_OLD = "TemplatePackageEngineTooOld";

function userError(name: string, message: string): UserError {
  return new UserError({ source: SOURCE, name, message });
}

function isRecord(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

function getString(rec: Record<string, unknown>, key: string): string | undefined {
  const v = rec[key];
  return typeof v === "string" ? v : undefined;
}

function getArray(rec: Record<string, unknown>, key: string): unknown[] | undefined {
  const v = rec[key];
  return Array.isArray(v) ? v : undefined;
}

function getRecord(rec: Record<string, unknown>, key: string): Record<string, unknown> | undefined {
  const v = rec[key];
  return isRecord(v) ? v : undefined;
}

/** The v4 `templateId`s a selector routes (engine `v4` routes only). */
function v4RouteIds(selectorData: unknown): string[] {
  const ids: string[] = [];
  if (!isRecord(selectorData)) {
    return ids;
  }
  const routes = getArray(selectorData, "routes") ?? [];
  for (const routeRaw of routes) {
    if (!isRecord(routeRaw) || getString(routeRaw, "engine") !== "v4") {
      continue;
    }
    const tid = getString(routeRaw, "templateId");
    if (tid !== undefined) {
      ids.push(tid);
    }
  }
  return ids;
}

/** Compare `major.minor.patch` numerically: <0 / 0 / >0. */
function compareSemver(a: string, b: string): number {
  const pa = parseSemver(a);
  const pb = parseSemver(b);
  for (let i = 0; i < 3; i++) {
    if (pa[i] !== pb[i]) {
      return pa[i] < pb[i] ? -1 : 1;
    }
  }
  return 0;
}

function parseSemver(v: string): number[] {
  const parts = v.split(".");
  const nums: number[] = [];
  for (let i = 0; i < 3; i++) {
    const n = Number.parseInt(parts[i] ?? "0", 10);
    nums.push(Number.isNaN(n) ? 0 : n);
  }
  return nums;
}

/**
 * Validate that one `<kind>/<id>` template package conforms to the ADR-0015
 * artifact shape and may run on this engine, before any content is rendered.
 *
 * @param kind   selects the per-kind `selector.json` + `templateId` namespace
 * @param id     which `<kind>/<id>/` package to validate (named in diagnostics)
 * @param mode   `build` (fail the build) or `load` (fail the scaffold); only the
 *               reverse `minEngineVersion` comparison is `load`-only
 * @param port   the narrow injected `TemplatePackagePort`
 * @returns `ok(ValidatedPackage)`, or a `UserError` naming the file + rule (or
 *          the required version) so the fix is unambiguous.
 */
export function validateTemplatePackage(
  kind: PackageKind,
  id: string,
  mode: ValidateMode,
  port: TemplatePackagePort
): Result<ValidatedPackage, FxError> {
  const pkg = `${kind}/${id}`;

  // 1. Four-file isomorphism (INV-1): descriptor / questions / pipeline are
  // always required, even when empty; `content/` is optional.
  const descriptor = port.descriptor();
  const questions = port.questions();
  const pipeline = port.pipeline();
  if (descriptor === undefined) {
    return err(userError(VALIDATE_REQUIRED_FILE, `${pkg}: descriptor.json is required`));
  }
  if (questions === undefined) {
    return err(userError(VALIDATE_REQUIRED_FILE, `${pkg}: questions.json is required`));
  }
  if (pipeline === undefined) {
    return err(userError(VALIDATE_REQUIRED_FILE, `${pkg}: pipeline.json is required`));
  }
  if (!isRecord(descriptor)) {
    return err(userError(VALIDATE_SCHEMA, `${pkg}: descriptor.json must be a JSON object`));
  }

  // 2. Schema validity: descriptor / questions / selector against their schemas.
  const selectorData = port.selector(kind);
  const dSchemaErr = port.schemas.descriptor(descriptor);
  if (dSchemaErr !== undefined) {
    return err(
      userError(VALIDATE_SCHEMA, `${pkg}: descriptor.json failed schema validation: ${dSchemaErr}`)
    );
  }
  const qSchemaErr = port.schemas.question(questions);
  if (qSchemaErr !== undefined) {
    return err(
      userError(VALIDATE_SCHEMA, `${pkg}: questions.json failed schema validation: ${qSchemaErr}`)
    );
  }
  const sSchemaErr = port.schemas.selector(selectorData);
  if (sSchemaErr !== undefined) {
    return err(
      userError(VALIDATE_SCHEMA, `${pkg}: selector.json failed schema validation: ${sSchemaErr}`)
    );
  }

  // 3. Placeholder closure (INV-3, both directions). Producers a content token
  // may reference: replaceMap vars ∪ answer vars (optionsSchema.properties ∪
  // question names) ∪ caller-injected floor. Vars that must be consumed: the
  // *required* (unconditionally-emitted {const}/{from}/{expr}) replaceMap vars;
  // a {when,value} var is optional (RCTX-04 / §3.4).
  const replaceMapVars: string[] = [];
  const requiredVars: string[] = [];
  for (const entryRaw of getArray(descriptor, "replaceMap") ?? []) {
    if (!isRecord(entryRaw)) {
      continue;
    }
    const v = getString(entryRaw, "var");
    if (v === undefined) {
      continue;
    }
    replaceMapVars.push(v);
    if (!("when" in entryRaw)) {
      requiredVars.push(v);
    }
  }

  const answerVars = new Set<string>();
  const optionsSchema = getRecord(descriptor, "optionsSchema");
  const props = optionsSchema === undefined ? undefined : getRecord(optionsSchema, "properties");
  if (props !== undefined) {
    for (const k of Object.keys(props)) {
      answerVars.add(k);
    }
  }
  if (isRecord(questions)) {
    for (const qRaw of getArray(questions, "questions") ?? []) {
      if (isRecord(qRaw)) {
        const n = getString(qRaw, "name");
        if (n !== undefined) {
          answerVars.add(n);
        }
      }
    }
  }

  const mayReference = new Set<string>([...replaceMapVars, ...answerVars, ...port.callerFloor()]);

  const contentFiles = port.content() ?? [];
  const contentTokens = new Set<string>();
  for (const file of contentFiles) {
    for (const token of file.placeholders) {
      contentTokens.add(token);
      if (!mayReference.has(token)) {
        return err(
          userError(
            VALIDATE_PLACEHOLDER_DRIFT,
            `${pkg}: content file '${file.path}' references '{{${token}}}', which no replaceMap entry, question, or caller-injected identifier produces`
          )
        );
      }
    }
  }
  for (const v of requiredVars) {
    if (!contentTokens.has(v)) {
      return err(
        userError(
          VALIDATE_PLACEHOLDER_DRIFT,
          `${pkg}: emits required render var '${v}' that no content file consumes`
        )
      );
    }
  }

  // 4. Routing derived from descriptors (INV-4 / ADR-0014 §5.3): every v4 route
  // resolves to a present descriptor, and the two kinds are disjoint.
  const thisV4Ids = v4RouteIds(selectorData);
  const present = new Set(port.presentTemplateIds(kind));
  for (const tid of thisV4Ids) {
    if (!present.has(tid)) {
      return err(
        userError(
          VALIDATE_DANGLING_ROUTE,
          `${pkg}: selector.json routes v4 templateId '${tid}', but no descriptor for it is present in the artifact`
        )
      );
    }
  }
  const otherKind: PackageKind = kind === "create" ? "modify" : "create";
  const otherV4Ids = new Set(v4RouteIds(port.selector(otherKind)));
  for (const tid of thisV4Ids) {
    if (otherV4Ids.has(tid)) {
      return err(
        userError(
          VALIDATE_KIND_OVERLAP,
          `templateId '${tid}' is routed in both the create and modify selectors; the two kinds own disjoint templateId namespaces`
        )
      );
    }
  }

  // 5. The reverse gate is explicit (INV-5): minEngineVersion is mandatory.
  const minEngineVersion = getString(descriptor, "minEngineVersion");
  if (minEngineVersion === undefined) {
    return err(
      userError(
        VALIDATE_MIN_ENGINE_MISSING,
        `${pkg}: descriptor.json must declare minEngineVersion (the reverse compatibility signal)`
      )
    );
  }

  // 6. Reverse compatibility, per package (INV-6) — only in `load` mode.
  if (mode === "load" && compareSemver(port.engineVersion(), minEngineVersion) < 0) {
    return err(
      userError(
        VALIDATE_ENGINE_TOO_OLD,
        `${pkg}: requires engine ${minEngineVersion}, but this engine is ${port.engineVersion()}; upgrade the engine (no silent fallback)`
      )
    );
  }

  return ok({ descriptor, minEngineVersion, contentFiles });
}
