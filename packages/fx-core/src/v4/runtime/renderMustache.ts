// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import Mustache from "mustache";
import { Result, err, ok } from "neverthrow";
import { RenderVars } from "../model/dataModel";

/**
 * The single Mustache surface the engine uses for both `.tpl` paths/bodies and
 * step `with` values (run-scaffold-pipeline INV-4 — "two surfaces, no third").
 *
 * Spec: docs/03-specs/operations/scaffolding/run-scaffold-pipeline.md (AC-18)
 *
 * It renders every token with a producer in `renderVars`, evaluates
 * `{{#…}}` / `{{^…}}` sections, performs no HTML escaping, and — crucially —
 * leaves a token with **no** producer literal (e.g. the env ref
 * `${{APP_NAME_SUFFIX}}`, which provision resolves later, never the engine). It
 * reimplements the v3 "escape empty variable" behavior in the v4 world (INV-7),
 * without reusing v3 code and without the v3 token-tuple cast.
 */

const SOURCE = "Scaffold";

/** `SystemError` name — a malformed template body the render surface cannot parse. */
export const RENDER_PARSE_ERROR = "RenderParseError";

const DELIMITERS: [string, string] = ["{{", "}}"];

// Private-use sentinels that wrap a producer-less token's name so the Mustache
// pass treats it as inert text; restored verbatim afterwards. They cannot occur
// in a real template, and never reach the returned string.
const SENTINEL_OPEN = "\uE000";
const SENTINEL_CLOSE = "\uE001";

// A bare value token `{{ name }}` (dotted names allowed). Never matches a
// section / inverted / comment / partial / unescaped tag — each of those
// carries a leading sigil (`#` `^` `/` `!` `>` `&`) the character class rejects.
const VALUE_TOKEN = /\{\{\s*([A-Za-z_][A-Za-z0-9_.]*)\s*\}\}/g;
const SENTINEL_TOKEN = /\uE000([^\uE001]+)\uE001/g;

export function renderMustache(template: string, renderVars: RenderVars): Result<string, FxError> {
  // Protect every producer-less value token so the Mustache pass cannot empty
  // it; a producer-backed token is left for Mustache to substitute.
  const guarded = template.replace(VALUE_TOKEN, (match: string, name: string) =>
    name in renderVars ? match : `${SENTINEL_OPEN}${name}${SENTINEL_CLOSE}`
  );

  const previousEscape = Mustache.escape;
  Mustache.escape = (value) => value; // no HTML escaping — URLs and ${{…}} refs pass through
  let rendered: string;
  try {
    rendered = Mustache.render(guarded, renderVars, {}, DELIMITERS);
  } catch (error) {
    return err(
      new SystemError({
        source: SOURCE,
        name: RENDER_PARSE_ERROR,
        message: `Failed to render a template body: ${
          error instanceof Error ? error.message : String(error)
        }`,
      })
    );
  } finally {
    Mustache.escape = previousEscape;
  }

  // Restore the protected producer-less tokens verbatim (e.g. `${{APP_NAME_SUFFIX}}`).
  const result = rendered.replace(SENTINEL_TOKEN, (_match: string, name: string) => `{{${name}}}`);
  return ok(result);
}
