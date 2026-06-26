// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { randomUUID } from "crypto";
import { FxError, SystemError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { RegisteredStep, StepContext, StepParams } from "../../pipeline/runScaffoldPipeline";

/** MetaOS post-render steps. */

const SOURCE = "Scaffold";

/** Engine step name `metaos/unify-project-id`. */
export const STEP_UNIFY_PROJECT_ID = "metaos/unify-project-id";

function systemError(name: string, message: string): SystemError {
  return new SystemError({ source: SOURCE, name, message });
}

function stringParam(params: StepParams, key: string): string | undefined {
  const value = params[key];
  return typeof value === "string" ? value : undefined;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
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

/** Registered step for mirroring v3 MetaOSHelper.unifyProjectID. */
export const metaOsUnifyProjectId: RegisteredStep = {
  validateParams(resolved: StepParams): string | undefined {
    if (stringParam(resolved, "manifestPath") === undefined) {
      return "missing string parameter 'manifestPath'";
    }
    if (stringParam(resolved, "envPath") === undefined) {
      return "missing string parameter 'envPath'";
    }
    return undefined;
  },
  apply(resolved: StepParams, ctx: StepContext): Result<void, FxError> {
    const manifestPath = stringParam(resolved, "manifestPath");
    const envPath = stringParam(resolved, "envPath");
    if (manifestPath === undefined || envPath === undefined) {
      return err(systemError("MetaOsUnifyParams", "resolved parameters are not all strings"));
    }

    const manifestRaw = readRequired(ctx, manifestPath, "MetaOsManifestMissing");
    if (manifestRaw.isErr()) {
      return err(manifestRaw.error);
    }
    const manifest = parseJsonObject(manifestRaw.value, "MetaOsManifestInvalid", manifestPath);
    if (manifest.isErr()) {
      return err(manifest.error);
    }
    const envRaw = readRequired(ctx, envPath, "MetaOsEnvMissing");
    if (envRaw.isErr()) {
      return err(envRaw.error);
    }

    const appId = randomUUID();
    manifest.value.id = appId;
    ctx.write(manifestPath, Buffer.from(JSON.stringify(manifest.value, null, 2) + "\n", "utf8"));
    ctx.write(envPath, Buffer.from(withTeamsAppId(envRaw.value.toString("utf8"), appId), "utf8"));
    return ok(undefined);
  },
};
