// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { RenderVars } from "../../model/dataModel";
import { Pipeline } from "../../pipeline/runScaffoldPipeline";
import { legacyOpenApiSearchBinding } from "./legacyOpenApiSearch";

const LEGACY_RENDER_BINDINGS = [legacyOpenApiSearchBinding];

export function applyLegacyRenderBindings(
  descriptor: unknown,
  pipeline: Pipeline,
  vars: RenderVars
): RenderVars {
  return LEGACY_RENDER_BINDINGS.reduce(
    (current, migrate) => migrate(descriptor, pipeline, current),
    vars
  );
}
