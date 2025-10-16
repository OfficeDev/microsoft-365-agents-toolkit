// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, err, InputsWithProjectPath, ok } from "@microsoft/teamsfx-api";
import {
  CollaborationConstants,
  featureFlagManager,
  FeatureFlags,
  PermissionListInputs,
  PermissionListOptions,
  QuestionNames,
} from "@microsoft/teamsfx-core";
import { getFxCore } from "../../activate";
import { logger } from "../../commonlib/logger";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { ProjectFolderOption } from "../common";
import { agentOwnerOption, azureMessage, spfxMessage } from "./permissionGrant";

export const permissionStatusCommand: CLICommand = {
  name: "status",
  description: commands["collaborator.status"].description,
  options: [
    ...PermissionListOptions,
    {
      name: "all",
      shortName: "a",
      description: commands["collaborator.status"].options["all"],
      type: "boolean",
      required: false,
    },
    ...(featureFlagManager.getBooleanValue(FeatureFlags.ShareEnabled) ? [agentOwnerOption] : []),
    ProjectFolderOption,
  ],
  telemetry: {
    event: TelemetryEvent.CheckPermission,
  },
  reservedOptionNamesInInteractiveMode: ["all"],
  handler: async (ctx) => {
    const inputs = ctx.optionValues as PermissionListInputs & InputsWithProjectPath;
    const listAll = inputs.all || false;
    const core = getFxCore();
    // print necessary messages
    logger.info(azureMessage);
    logger.info(spfxMessage);
    if (ctx.optionValues["agent"]) {
      inputs[QuestionNames.collaborationAppType] = [CollaborationConstants.AgentOptionId];
    }
    const result = listAll
      ? await core.listCollaborator(inputs)
      : await core.checkPermission(inputs);
    if (result.isErr()) {
      return err(result.error);
    }
    return ok(undefined);
  },
};
