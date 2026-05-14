// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Removes unused components, tags, security schemas, and vendor extensions
 * from an OpenAPI document to produce a minimal spec.
 */

import type { OpenAPIV3 } from "openapi-types";

export interface OptimizerOptions {
  removeUnusedComponents: boolean;
  removeUnusedTags: boolean;
  removeUserDefinedRootProperty: boolean;
  removeUnusedSecuritySchemas: boolean;
}

const DEFAULT_OPTIONS: Required<OptimizerOptions> = {
  removeUnusedComponents: true,
  removeUnusedTags: true,
  removeUserDefinedRootProperty: true,
  removeUnusedSecuritySchemas: true,
};

/**
 * Return a deep-cloned, optimized copy of `spec`.
 */
export function optimizeSpec(
  spec: OpenAPIV3.Document,
  options?: Partial<OptimizerOptions>
): OpenAPIV3.Document {
  const opts = { ...DEFAULT_OPTIONS, ...options };
  const newSpec: OpenAPIV3.Document = JSON.parse(JSON.stringify(spec));

  if (opts.removeUserDefinedRootProperty) removeUserDefinedRootProperty(newSpec);
  if (opts.removeUnusedComponents) removeUnusedComponents(newSpec);
  if (opts.removeUnusedTags) removeUnusedTags(newSpec);
  if (opts.removeUnusedSecuritySchemas) removeUnusedSecuritySchemas(newSpec);

  return newSpec;
}

// ---------------------------------------------------------------------------
// Internals
// ---------------------------------------------------------------------------

function removeUserDefinedRootProperty(spec: OpenAPIV3.Document): void {
  for (const key of Object.keys(spec)) {
    if (key.startsWith("x-")) {
      delete (spec as unknown as Record<string, unknown>)[key];
    }
  }
}

function removeUnusedTags(spec: OpenAPIV3.Document): void {
  if (!spec.tags) return;

  const usedTags = new Set<string>();
  for (const pathKey in spec.paths) {
    for (const methodKey in spec.paths[pathKey]) {
      const operation = (spec.paths[pathKey] as unknown as Record<string, unknown>)[
        methodKey
      ] as OpenAPIV3.OperationObject;
      operation.tags?.forEach((tag: string) => usedTags.add(tag));
    }
  }

  spec.tags = spec.tags.filter((t: OpenAPIV3.TagObject) => usedTags.has(t.name));
}

function removeUnusedSecuritySchemas(spec: OpenAPIV3.Document): void {
  if (!spec.components?.securitySchemes) return;

  const used = new Set<string>();

  for (const pathKey in spec.paths) {
    for (const methodKey in spec.paths[pathKey]) {
      const operation = (spec.paths[pathKey] as unknown as Record<string, unknown>)[
        methodKey
      ] as OpenAPIV3.OperationObject;
      operation.security?.forEach((req: OpenAPIV3.SecurityRequirementObject) => {
        for (const key in req) used.add(key);
      });
    }
  }

  spec.security?.forEach((req: OpenAPIV3.SecurityRequirementObject) => {
    for (const key in req) used.add(key);
  });

  for (const key in spec.components.securitySchemes) {
    if (!used.has(key)) delete spec.components.securitySchemes[key];
  }

  if (Object.keys(spec.components.securitySchemes).length === 0) {
    delete spec.components.securitySchemes;
  }
  if (Object.keys(spec.components).length === 0) {
    delete spec.components;
  }
}

function removeUnusedComponents(spec: OpenAPIV3.Document): void {
  const components = spec.components;
  if (!components) return;

  // Temporarily remove components so JSON.stringify only captures usage from paths/etc
  delete spec.components;
  const specString = JSON.stringify(spec);
  const references = getComponentReferences(specString);

  const usedSet = new Set<string>();
  for (const ref of references) {
    addComponentTransitive(ref, usedSet, components);
  }

  // Rebuild components with only used entries
  const newComponents: Record<string, unknown> = {};
  for (const componentPath of usedSet) {
    const parts = componentPath.split("/");
    const component = getComponent(componentPath, components);
    if (!component) continue;

    let current: Record<string, unknown> = newComponents;
    for (let i = 2; i < parts.length; i++) {
      if (i === parts.length - 1) {
        current[parts[i]] = component;
      } else {
        current[parts[i]] = (current[parts[i]] as Record<string, unknown>) ?? {};
      }
      current = current[parts[i]] as Record<string, unknown>;
    }
  }

  // Security schemes referenced by name — keep all, trim with removeUnusedSecuritySchemas
  if (components.securitySchemes) {
    newComponents.securitySchemes = components.securitySchemes;
  }

  if (Object.keys(newComponents).length > 0) {
    spec.components = newComponents as OpenAPIV3.ComponentsObject;
  }
}

// ---------------------------------------------------------------------------
// Component reference traversal
// ---------------------------------------------------------------------------

function getComponentReferences(specString: string): string[] {
  const matches = Array.from(specString.matchAll(/['"](#\/components\/.+?)['"]/g));
  return matches.map((m) => m[1]);
}

function getComponent(
  componentPath: string,
  components: OpenAPIV3.ComponentsObject
): unknown | null {
  const parts = componentPath.split("/");
  let current: unknown = components;
  for (const part of parts) {
    if (part === "#" || part === "components") continue;
    current = (current as Record<string, unknown>)?.[part];
    if (!current) return null;
  }
  return current;
}

function addComponentTransitive(
  componentName: string,
  usedSet: Set<string>,
  components: OpenAPIV3.ComponentsObject
): void {
  if (usedSet.has(componentName)) return;
  usedSet.add(componentName);

  const component = getComponent(componentName, components);
  if (component) {
    const refs = getComponentReferences(JSON.stringify(component));
    for (const ref of refs) {
      addComponentTransitive(ref, usedSet, components);
    }
  }
}
