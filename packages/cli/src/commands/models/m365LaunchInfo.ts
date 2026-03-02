// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, err, ok } from "@microsoft/teamsfx-api";
import { PackageService } from "@microsoft/teamsfx-core";
import { logger } from "../../commonlib/logger";
import { MissingRequiredOptionError } from "../../error";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { m365utils } from "./m365Sideloading";
import {
  getResourceServiceEndpoint,
  ResourceServiceType,
} from "@microsoft/teamsfx-core/build/common/constants";

export const m365LaunchInfoCommand: CLICommand = {
  name: "launchinfo",
  description: commands.launchinfo.description,
  options: [
    {
      name: "title-id",
      description: commands.launchinfo.options["title-id"],
      type: "string",
    },
    {
      name: "manifest-id",
      description: commands.launchinfo.options["manifest-id"],
      type: "string",
    },
  ],
  examples: [
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} launchinfo --title-id U_xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`,
      description: "Get launch information of the acquired M365 App by Title ID",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} launchinfo --manifest-id xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`,
      description: "Get launch information of the acquired M365 App by Manifest ID",
    },
  ],
  telemetry: {
    event: TelemetryEvent.M365LaunchInfo,
  },
  defaultInteractiveOption: false,
  handler: async (ctx) => {
    const packageService = new PackageService(
      getResourceServiceEndpoint(ResourceServiceType.MOS3),
      logger
    );
    let titleId = ctx.optionValues["title-id"] as string;
    const manifestId = ctx.optionValues["manifest-id"] as string;
    if (titleId === undefined && manifestId === undefined) {
      return err(
        new MissingRequiredOptionError(ctx.command.fullName, `--title-id or --manifest-id`)
      );
    }
    const tokenAndUpn = await m365utils.getTokenAndUpn();
    if (titleId === undefined) {
      titleId = await packageService.retrieveTitleId(tokenAndUpn[0], manifestId);
    }
    await packageService.getLaunchInfoByTitleId(tokenAndUpn[0], titleId);
    return ok(undefined);
  },
};
