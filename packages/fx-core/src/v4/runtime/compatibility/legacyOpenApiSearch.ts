// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { lt, valid } from "semver";
import { RenderVars } from "../../model/dataModel";
import { Pipeline } from "../../pipeline/runScaffoldPipeline";

export function legacyOpenApiSearchBinding(
  descriptor: unknown,
  pipeline: Pipeline,
  vars: RenderVars
): RenderVars {
  if (
    typeof descriptor !== "object" ||
    descriptor === null ||
    !("id" in descriptor) ||
    descriptor.id !== "da/api-plugin-from-existing-api" ||
    !("minEngineVersion" in descriptor) ||
    typeof descriptor.minEngineVersion !== "string" ||
    !valid(descriptor.minEngineVersion) ||
    !lt(descriptor.minEngineVersion, "6.12.0") ||
    vars.apiSpecLocation !== undefined
  ) {
    return vars;
  }
  const source = vars["derived.openapi.operations.apiSpecLocation"];
  const hasLegacyBinding = pipeline.steps.some(
    (step) =>
      step.step === "openapi/generate-plugin-files" &&
      step.with?.apiSpecLocation === "{{apiSpecLocation}}"
  );
  return hasLegacyBinding && typeof source === "string"
    ? { ...vars, apiSpecLocation: source }
    : vars;
}
