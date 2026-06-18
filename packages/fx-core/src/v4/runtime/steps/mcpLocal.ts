// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { RegisteredStep, StepContext, StepParams } from "../../pipeline/runScaffoldPipeline";

/**
 * The `mcp-local/materialize-servers` pipeline step (ADR-0017 whitelist;
 * domain-typed name, INV-6) for the `da/mcp-server` create flow's **local**
 * branch. A v4 render var is a scalar string or a `string[]` (collect-inputs
 * INV-7), so it cannot carry the array-of-objects the local `.vscode/mcp.json`
 * needs; that file is therefore written programmatically by this step, which
 * overwrites the render-phase remote stub (read-modify-write,
 * run-scaffold-pipeline AC-21).
 *
 * It reads two `with` values: `selected` — the `string[]` of chosen server ids
 * (the multiSelect selection passed structurally, AC-22) — and `catalog` — a
 * provider-populated JSON object string mapping each id to its `{command, args}`
 * launch spec. Each selected id becomes a `stdio` server entry, faithful to the
 * v3 local output contract (`{type:"stdio", command, args}`).
 *
 * Spec: docs/03-specs/scenarios/da/create-mcp-server.md
 *       (SCN-CREATE-MCP-11..14)
 *
 * v4-owned (INV-7): imports no v3 symbol. The live catalog provider (odr.exe
 * enumeration) is the deferred external dependency; this step consumes the
 * already-resolved `catalog` answer, so it is fully CI-testable in isolation.
 */

const SOURCE = "Scaffold";

/** Engine step name `mcp-local/materialize-servers`. */
export const STEP_MATERIALIZE_LOCAL_SERVERS = "mcp-local/materialize-servers";

function systemError(name: string, message: string): SystemError {
  return new SystemError({ source: SOURCE, name, message });
}

/** Read a `with` value as a string, or `undefined` if it is absent / non-string. */
function stringParam(params: StepParams, key: string): string | undefined {
  const value = params[key];
  return typeof value === "string" ? value : undefined;
}

/** Read a `with` value as the multiSelect `string[]` (AC-22), or `undefined`. */
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

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

/** One materialized stdio server entry (the v3 local `.vscode/mcp.json` shape). */
interface LocalServer {
  type: "stdio";
  command: string;
  args: string[];
}

/**
 * Build one stdio server from its catalog record, validating the boundary: a
 * selected id absent from the catalog, or an entry without a string `command`,
 * is an engine-side break a build/collect gate should have caught (SystemError).
 */
function toLocalServer(id: string, entry: unknown): Result<LocalServer, FxError> {
  if (!isRecord(entry)) {
    return err(
      systemError("McpLocalCatalogEntry", `local server catalog has no entry for '${id}'`)
    );
  }
  const command = entry.command;
  if (typeof command !== "string") {
    return err(
      systemError("McpLocalCatalogCommand", `local server '${id}' has no string 'command'`)
    );
  }
  const rawArgs = entry.args;
  if (
    rawArgs !== undefined &&
    (!Array.isArray(rawArgs) || !rawArgs.every((a) => typeof a === "string"))
  ) {
    return err(
      systemError("McpLocalCatalogArgs", `local server '${id}' has a non-string-array 'args'`)
    );
  }
  return ok({ type: "stdio", command, args: rawArgs === undefined ? [] : rawArgs });
}

/**
 * `mcp-local/materialize-servers` — write `.vscode/mcp.json` for the local
 * branch (one `stdio` server per selected id), overwriting the render-phase
 * remote stub.
 */
export const mcpLocalMaterializeServers: RegisteredStep = {
  validateParams(resolved: StepParams): string | undefined {
    if (stringParam(resolved, "target") === undefined) {
      return "missing string parameter 'target'";
    }
    if (stringArrayParam(resolved, "selected") === undefined) {
      return "missing string[] parameter 'selected'";
    }
    if (stringParam(resolved, "catalog") === undefined) {
      return "missing string parameter 'catalog'";
    }
    return undefined;
  },
  apply(resolved: StepParams, ctx: StepContext): Result<void, FxError> {
    const target = stringParam(resolved, "target");
    const selected = stringArrayParam(resolved, "selected");
    const catalogRaw = stringParam(resolved, "catalog");
    if (target === undefined || selected === undefined || catalogRaw === undefined) {
      return err(
        systemError("McpLocalParams", "resolved parameters are not all of the expected type")
      );
    }
    let parsed: unknown;
    try {
      parsed = JSON.parse(catalogRaw);
    } catch (error) {
      return err(
        systemError(
          "McpLocalCatalogParse",
          `local server catalog is not valid JSON: ${errorMessage(error)}`
        )
      );
    }
    if (!isRecord(parsed)) {
      return err(systemError("McpLocalCatalogShape", "local server catalog is not a JSON object"));
    }
    const servers: Record<string, LocalServer> = {};
    for (const id of selected) {
      const built = toLocalServer(id, parsed[id]);
      if (built.isErr()) {
        return err(built.error);
      }
      servers[id] = built.value;
    }
    ctx.write(target, Buffer.from(JSON.stringify({ servers }, null, 2) + "\n", "utf8"));
    return ok(undefined);
  },
};
