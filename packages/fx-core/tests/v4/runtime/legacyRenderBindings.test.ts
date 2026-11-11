// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "vitest";
import { applyLegacyRenderBindings } from "../../../src/v4/runtime/compatibility";
import { Pipeline } from "../../../src/v4/pipeline/runScaffoldPipeline";
import { RenderVars } from "../../../src/v4/model/dataModel";

const descriptor = { id: "da/api-plugin-from-existing-api", minEngineVersion: "6.11.0" };
const pipeline: Pipeline = {
  pipeline: "default",
  steps: [
    { step: "openapi/generate-plugin-files", with: { apiSpecLocation: "{{apiSpecLocation}}" } },
  ],
};
const vars = { "derived.openapi.operations.apiSpecLocation": "/spec.yaml" };

describe("legacy render bindings", () => {
  for (const unchanged of [
    undefined,
    {},
    { ...descriptor, id: "other-template" },
    { ...descriptor, minEngineVersion: "6.12.0" },
    { ...descriptor, minEngineVersion: "invalid" },
  ]) {
    it(`CLEAN-06: leaves non-legacy descriptors unchanged: ${JSON.stringify(unchanged)}`, () => {
      assert.strictEqual(applyLegacyRenderBindings(unchanged, pipeline, vars), vars);
    });
  }

  it("CLEAN-06: does not supply missing sources, replace explicit values, or rewrite other bindings", () => {
    const noSource: RenderVars = {};
    const explicit = { ...vars, apiSpecLocation: "/explicit.yaml" };
    assert.strictEqual(applyLegacyRenderBindings(descriptor, pipeline, noSource), noSource);
    assert.strictEqual(applyLegacyRenderBindings(descriptor, pipeline, explicit), explicit);
    assert.strictEqual(
      applyLegacyRenderBindings(descriptor, { pipeline: "default", steps: [] }, vars),
      vars
    );
    assert.strictEqual(
      applyLegacyRenderBindings(
        descriptor,
        {
          pipeline: "default",
          steps: [
            {
              step: "openapi/generate-plugin-files",
              with: { apiSpecLocation: "{{CustomSource}}" },
            },
          ],
        },
        vars
      ),
      vars
    );
  });
});
