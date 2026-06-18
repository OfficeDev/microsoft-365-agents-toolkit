// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * The shared data model of the v4 scaffolding world — the data nouns that flow
 * **between** operations (the wires), as opposed to any single operation's
 * private behavior types or its dependency-inverted `*Port`.
 *
 * This module is a pure leaf: it references only the `Record` / `Buffer`
 * globals and imports no other v4 symbol, so every operation can depend on it
 * without risking an import cycle. Each operation still owns its own behavior
 * types and its own narrow `*Port` (interface segregation, INV-7); only a noun
 * that crosses an operation boundary is hoisted here. This is also where
 * `Answers` lives **once** — it was previously declared identically in both
 * `collect-inputs` (its producer) and `build-render-context` (its consumer).
 *
 * v4 lives isolated from v3 (scaffolding.create.proposal.md §5.1): nothing here
 * is tailored for v3.
 */

// --- collect-inputs → build-render-context ---

/**
 * The resolved answer object: each asked question's value ∪ provider
 * `derived.<id>.<key>`. Produced by `collect-inputs`, consumed by
 * `build-render-context` (which reads it as raw answers ∪ `derived.*`).
 *
 * A value is a scalar `string` for every question kind except `multiSelect`,
 * whose answer is a `string[]` of the selected ids (collect-inputs INV-7). The
 * list is carried verbatim into `build-render-context` `{from}` and step `with`,
 * but never into the scalar expression grammar (`evaluate-expression` stays
 * scalar-valued, ADR-0016 decision 9 unchanged).
 */
export type Answers = Record<string, string | string[]>;

// --- caller → build-render-context → run-scaffold-pipeline ---

/**
 * The caller-injected identifier floor (camelCase: `appName`, the `language`
 * axis, …). Available to `{from}` / `{expr}` as evaluator scope but never
 * emitted into `renderVars` by `build-render-context`, and never shadowed
 * (build-render-context spec INV-2).
 */
export type CallerFloor = Record<string, string>;

/**
 * The render variable map `content/**` and step `with` values interpolate
 * against. A value is a scalar `string`, or a `string[]` for a `multiSelect`
 * answer carried verbatim through `replaceMap` `{from}` (collect-inputs INV-7);
 * a list reaches a step `with` structurally, never a scalar `content/**` token.
 */
export type RenderVars = Record<string, string | string[]>;

// --- (caller / bridge) → open-template-package → run-scaffold-pipeline ---

/**
 * Names which template inside the package to open. The boundary
 * (open → locate → hand back entries) is permanent; this locator shape is
 * transitional — `{ language, scenario }` matches the package's current
 * `<language>/<scenario>/` layout and becomes `{ templateId }` when the
 * proposal §3 authoring layout ships. Only the resolved prefix changes; the
 * open/entry contract and every AC hold (open-template-package INV-1).
 */
export interface TemplateLocator {
  language: string;
  scenario: string;
}

/**
 * Names which authored declarative package inside the channel zip to open: the
 * `v4/<kind>/<templateId>/` subtree (descriptor + pipeline + content). The
 * sibling of {@link TemplateLocator} for the declarative engine — `kind` selects
 * the create / modify namespace, `templateId` the package within it.
 */
export interface DeclarativeLocator {
  /** The package namespace: `create` or `modify`. */
  kind: string;
  /** The package id within the kind, e.g. `da/mcp-server`. */
  templateId: string;
}

/** One file from the located template, rooted at the template's content. */
export interface TemplateFileEntry {
  /** Path relative to the located template's content root, forward-slash normalized. */
  path: string;
  /** The file's raw bytes, verbatim — unrendered, `.tpl` suffix intact. */
  data: Buffer;
}
