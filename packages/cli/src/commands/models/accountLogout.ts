// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, ok } from "@microsoft/teamsfx-api";
import AzureTokenProvider from "../../commonlib/azureLogin";
import { logger } from "../../commonlib/logger";
import M365TokenProvider from "../../commonlib/M365TokenProviderWrapper";
import { commands, strings } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";

export const accountLogoutCommand: CLICommand = {
  name: "logout",
  description: commands["auth.logout"].description,
  arguments: [
    {
      type: "string",
      name: "service",
      description: commands["auth.logout"].arguments.service,
      choices: ["azure", "m365"],
      required: true,
    },
  ],
  telemetry: {
    event: TelemetryEvent.AccountLogout,
  },
  defaultInteractiveOption: false,
  handler: async (ctx) => {
    const service = ctx.argumentValues[0];
    switch (service) {
      case "azure": {
        ctx.telemetryProperties.service = "azure";
        const result = await AzureTokenProvider.signout();
        if (result) {
          logger.info(strings["account.logout.azure"]);
        } else {
          logger.error(strings["account.logout.azure.fail"]);
        }
        break;
      }
      case "m365": {
        ctx.telemetryProperties.service = "m365";
        const result = await M365TokenProvider.signout();
        if (result) {
          logger.info(strings["account.logout.m365"]);
        } else {
          logger.error(strings["account.logout.m365.fail"]);
        }
        break;
      }
    }
    return ok(undefined);
  },
};
