// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Filter an OpenAPI spec to include only selected operations, then optimize.
 *
 * Expects filter items in the format "METHOD /path" (e.g. "GET /pets").
 */

import type { OpenAPIV3 } from "openapi-types";
import type { ParseOptions } from "./types";
import { convertPathToCamelCase } from "./utils";
import { createValidator } from "./validator";
import { optimizeSpec } from "./optimizer";
import { HTTPMethods as HTTPMethodsConst } from "./constants";

/**
 * Return a new spec containing only the operations listed in `filter`.
 * Operations that fail validation are silently skipped.
 * The result is optimized (unused components, tags, security removed).
 */
export function filterSpec(
  filter: string[],
  unresolvedSpec: OpenAPIV3.Document,
  resolvedSpec: OpenAPIV3.Document,
  options: ParseOptions
): OpenAPIV3.Document {
  const newSpec = { ...unresolvedSpec };
  const newPaths: OpenAPIV3.PathsObject = {};

  const validator = createValidator(resolvedSpec, options);

  for (const filterItem of filter) {
    const [method, path] = filterItem.split(" ");
    if (!path || !path.startsWith("/")) continue;
    const methodName = method.toLowerCase();

    const pathObj = resolvedSpec.paths?.[path] as Record<string, unknown> | undefined;
    if (
      !HTTPMethodsConst.AllOperationMethods.includes(
        methodName as (typeof HTTPMethodsConst.AllOperationMethods)[number]
      ) ||
      !pathObj ||
      !pathObj[methodName]
    )
      continue;

    const validateResult = validator.validateAPI(methodName, path);
    if (!validateResult.isValid) continue;

    if (!newPaths[path]) {
      // Copy path-level properties (servers, parameters, etc.) but not operations
      newPaths[path] = { ...unresolvedSpec.paths![path] };
      for (const m of HTTPMethodsConst.AllOperationMethods) {
        delete (newPaths[path] as Record<string, unknown>)[m];
      }
    }

    (newPaths[path] as Record<string, unknown>)[methodName] = (
      unresolvedSpec.paths![path] as Record<string, unknown>
    )[methodName];

    // Auto-generate operationId if missing
    if (!(newPaths[path] as Record<string, unknown>)[methodName]) continue;
    const op = (newPaths[path] as Record<string, unknown>)[methodName] as OpenAPIV3.OperationObject;
    if (!op.operationId) {
      op.operationId = `${methodName}${convertPathToCamelCase(path)}`;
    }
  }

  newSpec.paths = newPaths;
  return optimizeSpec(newSpec);
}
