// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result, UserError } from "@microsoft/teamsfx-api";
import {
  ActionStartOptions,
  CapabilityOptions,
  ProjectTypeOptions,
  QuestionNames,
} from "@microsoft/teamsfx-core";
import { ExtensionSource } from "../error/error";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import {
  TelemetryEvent,
  TelemetryProperty,
  TelemetrySuccess,
} from "../telemetry/extTelemetryEvents";
import { localize } from "../utils/localizeUtils";
import { getSystemInputs } from "../utils/systemEnvUtils";
import { getTriggerFromProperty } from "../utils/telemetryUtils";
import { createNewProjectHandler } from "./lifecycleHandlers";

export async function createDeclarativeAgentWithApiSpec(
  args?: any[]
): Promise<Result<any, FxError>> {
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CreateDeclarativeAgentWithApiSpecStart,
    getTriggerFromProperty(args)
  );
  if (!args || args.length !== 1 || !args[0] || typeof args[0] !== "string") {
    const error = new UserError(
      ExtensionSource,
      "invalidParameter",
      localize("teamstoolkit.handler.createDeclarativeAgentWithApiSpec.error.invalidParameter")
    );
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.CreateDeclarativeAgentWithApiSpec, error);
    return err(error);
  }

  const specPath = args[0];

  const inputs = getSystemInputs();
  inputs[QuestionNames.ApiSpecLocation] = specPath;
  inputs[QuestionNames.ActionType] = ActionStartOptions.apiSpec().id;
  inputs.capabilities = CapabilityOptions.declarativeAgent().id;
  inputs[QuestionNames.WithPlugin] = "yes";
  inputs[QuestionNames.ProjectType] = ProjectTypeOptions.copilotAgentOptionId;
  inputs[QuestionNames.OpenAPISpecType] = "enter-url-or-open-local-file";

  const result = await createNewProjectHandler("", inputs);

  if (result.isErr()) {
    ExtTelemetry.sendTelemetryErrorEvent(
      TelemetryEvent.CreateDeclarativeAgentWithApiSpec,
      result.error
    );
    return err(result.error);
  }

  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CreateDeclarativeAgentWithApiSpec, {
    [TelemetryProperty.Success]: TelemetrySuccess.Yes,
  });

  return ok({});
}
