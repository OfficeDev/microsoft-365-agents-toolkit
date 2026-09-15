// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { capabilityDeclarations } from "../../capabilities/declarations";
import { RegisteredStep } from "../../pipeline/runScaffoldPipeline";
import { defineStep } from "../../pipeline/defineStep";
import { stringParam } from "../../pipeline/stepParams";
import { DaManifestService, daManifestService } from "../services/daManifestService";

/** Declarative Agent manifest mutation steps for modify flows. */

const SOURCE = "Scaffold";

/** Engine step name `da-action/register-plugin-manifest`. */
export const STEP_REGISTER_PLUGIN_MANIFEST = capabilityDeclarations.step.registerPluginManifest.id;

function systemError(name: string, message: string): SystemError {
  return new SystemError({ source: SOURCE, name, message });
}

export function createDaActionRegisterPluginManifestStep(
  manifests: DaManifestService = daManifestService
): RegisteredStep {
  return defineStep({
    parse(resolved): Result<{ teamsManifestPath: string; pluginManifestPath: string }, string> {
      const teamsManifestPath = stringParam(resolved, "teamsManifestPath");
      if (teamsManifestPath === undefined) {
        return err("missing string parameter 'teamsManifestPath'");
      }
      const pluginManifestPath = stringParam(resolved, "pluginManifestPath");
      if (pluginManifestPath === undefined) {
        return err("missing string parameter 'pluginManifestPath'");
      }
      return ok({ teamsManifestPath, pluginManifestPath });
    },
    invalidParams: () =>
      systemError("DaActionRegisterParams", "resolved parameters are not all strings"),
    apply({ teamsManifestPath, pluginManifestPath }, ctx): Result<void, FxError> {
      return manifests.registerDeclarativeAgentAction(ctx, teamsManifestPath, pluginManifestPath);
    },
  });
}

/** Registered step for adding a rendered API plugin manifest as a DA action. */
export const daActionRegisterPluginManifest = createDaActionRegisterPluginManifestStep();
