// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, InputsWithProjectPath, err, ok } from "@microsoft/teamsfx-api";
import { PermissionGrantInputs, PermissionGrantOptions } from "@microsoft/teamsfx-core";
import { getFxCore } from "../../activate";
import { logger } from "../../commonlib/logger";
import { MissingRequiredOptionError } from "../../error";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { ProjectFolderOption } from "../common";

export const azureMessage =
  "Notice: Azure resources permission needs to be handled by subscription owner since privileged account is " +
  "required to grant permission to Azure resources.\n" +
  "Assign Azure roles using the Azure portal: " +
  "https://docs.microsoft.com/en-us/azure/role-based-access-control/role-assignments-portal?tabs=current";

export const spfxMessage =
  "Notice: SPFX deployment permission needs to be handled manually by SharePoint site administrator.\n" +
  "Manage site admins using SharePoint admin center: " +
  "https://docs.microsoft.com/en-us/sharepoint/manage-site-collection-administrators";

export const permissionGrantCommand: CLICommand = {
  name: "grant",
  description: commands["collaborator.grant"].description,
  options: [...PermissionGrantOptions, ProjectFolderOption],
  telemetry: {
    event: TelemetryEvent.GrantPermission,
  },
  examples: [
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} collaborator grant -i false --teams-manifest-file ./appPackage/manifest.json --env dev --email other@email.com`,
      description: "Grant permission for another Microsoft 365 account to collaborate on the app.",
    },
  ],
  handler: async (ctx) => {
    const inputs = ctx.optionValues as PermissionGrantInputs & InputsWithProjectPath;
    // print necessary messages
    logger.info(azureMessage);
    logger.info(spfxMessage);
    if (!ctx.globalOptionValues.interactive) {
      if (!inputs["manifest-file-path"] && !inputs["manifest-path"]) {
        return err(
          new MissingRequiredOptionError(
            ctx.command.fullName,
            "--manifest-file-path or --manifest-path"
          )
        );
      }
    }
    // setAppTypeInputs(inputs);// app type input is unused in FxCore
    const core = getFxCore();
    const result = await core.grantPermission(inputs);
    if (result.isErr()) {
      return err(result.error);
    }
    return ok(undefined);
  },
};
