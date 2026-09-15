// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { FxError, SystemError } from "@microsoft/teamsfx-api";
import { Result, err, ok } from "neverthrow";
import { capabilityDeclarations } from "../../capabilities/declarations";
import { RegisteredStep, StepParams } from "../../pipeline/runScaffoldPipeline";
import { defineStep } from "../../pipeline/defineStep";
import { DaManifestService, daManifestService } from "../services/daManifestService";

const SOURCE = "Scaffold";

/** Engine step name `da/set-sensitivity-label`. */
export const STEP_SET_SENSITIVITY_LABEL = capabilityDeclarations.step.setSensitivityLabel.id;

function systemError(name: string, message: string): SystemError {
  return new SystemError({ source: SOURCE, name, message });
}

function manifestPath(params: StepParams): string | undefined {
  const value = params.manifestPath;
  return typeof value === "string" && value.trim().length > 0 ? value : undefined;
}

/** Best-effort service used by the sensitivity-label step. */
export interface GeneralSensitivityLabelService {
  resolveId(): Promise<string | undefined>;
}

/** Offline/default service used when no authenticated M365 adapter is registered. */
export const NOOP_GENERAL_SENSITIVITY_LABEL_SERVICE: GeneralSensitivityLabelService = {
  resolveId: (): Promise<undefined> => Promise.resolve(undefined),
};

/** Bind the General-label lookup dependency into its registered step. */
export function createDaSetSensitivityLabelStep(
  generalSensitivityLabel: GeneralSensitivityLabelService,
  manifests: DaManifestService = daManifestService
): RegisteredStep {
  return defineStep({
    parse(resolved): Result<string, string> {
      const path = manifestPath(resolved);
      return path === undefined
        ? err("missing non-empty string parameter 'manifestPath'")
        : ok(path);
    },
    invalidParams: () =>
      systemError("DaSensitivityLabelParams", "resolved manifestPath is not a non-empty string"),
    async apply(path, ctx): Promise<Result<void, FxError>> {
      const id = await generalSensitivityLabel.resolveId();
      if (id === undefined) {
        return ok(undefined);
      }

      if (manifests.setSensitivityLabel === undefined) {
        return err(
          systemError(
            "DaSensitivityLabelWrapperMissing",
            "the Declarative Agent manifest wrapper cannot set a sensitivity label"
          )
        );
      }
      return manifests.setSensitivityLabel(ctx, path, id);
    },
  });
}

/** Default step binding for offline runtimes and compatibility consumers. */
export const daSetSensitivityLabel = createDaSetSensitivityLabelStep(
  NOOP_GENERAL_SENSITIVITY_LABEL_SERVICE
);
