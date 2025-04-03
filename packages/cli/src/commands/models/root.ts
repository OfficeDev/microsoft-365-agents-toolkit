// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, ok } from "@microsoft/teamsfx-api";
import { logger } from "../../commonlib/logger";
import { FooterText } from "../../constants";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { getVersion } from "../../utils";
import { helper } from "../helper";
import { accountCommand } from "./account";
import { addCommand } from "./add";
import { getCreateCommand } from "./create";
import { deployCommand } from "./deploy";
import { entraAppCommand } from "./entraAppUpdate";
import { envCommand } from "./env";
import { listCommand } from "./list";
import { m365LaunchInfoCommand } from "./m365LaunchInfo";
import { m365SideloadingCommand } from "./m365Sideloading";
import { m365UnacquireCommand } from "./m365Unacquire";
import { permissionCommand } from "./permission";
import { previewCommand } from "./preview";
import { provisionCommand } from "./provision";
import { teamsappDoctorCommand } from "./teamsapp/doctor";
import { teamsappPackageCommand } from "./teamsapp/package";
import { teamsappPublishCommand } from "./teamsapp/publish";
import { teamsappUpdateCommand } from "./teamsapp/update";
import { teamsappValidateCommand } from "./teamsapp/validate";
import { upgradeCommand } from "./upgrade";
import { commands } from "../../resource";
import { shareCommand } from "./share";
import { setCommand } from "./set";
import { featureFlagManager, FeatureFlags } from "@microsoft/teamsfx-core";

export const helpCommand: CLICommand = {
  name: "help",
  description: commands.help.description,
  handler: () => {
    const helpText = helper.formatHelp(rootCommand, undefined);
    logger.info(helpText);
    return ok(undefined);
  },
};
export const rootCommand: CLICommand = {
  name: "teamsapp",
  fullName: "teamsapp",
  description: "Microsoft Teams Toolkit CLI.",
  version: getVersion(),
  footer: FooterText,
  commands: [
    accountCommand,
    getCreateCommand(),
    addCommand(),
    provisionCommand,
    deployCommand,
    ...(featureFlagManager.getBooleanValue(FeatureFlags.ShareEnabled) ? [shareCommand] : []),
    previewCommand,
    envCommand,
    permissionCommand,
    upgradeCommand,
    listCommand,
    helpCommand,
    teamsappUpdateCommand,
    teamsappValidateCommand,
    teamsappPackageCommand,
    teamsappPublishCommand,
    teamsappDoctorCommand,
    entraAppCommand,
    m365SideloadingCommand,
    m365UnacquireCommand,
    m365LaunchInfoCommand,
    ...(featureFlagManager.getBooleanValue(FeatureFlags.SensitivityLabelEnabled)
      ? [setCommand()]
      : []),
  ],
  sortCommands: true,
  options: [
    {
      type: "boolean",
      name: "version",
      shortName: "v",
      description: commands.root.options.version,
    },
    {
      type: "boolean",
      name: "help",
      shortName: "h",
      description: commands.root.options.help,
    },
    {
      type: "boolean",
      name: "interactive",
      shortName: "i",
      description: commands.root.options.interactive,
      default: true,
    },
    {
      type: "boolean",
      name: "debug",
      description: commands.root.options.debug,
      default: false,
    },
    {
      type: "boolean",
      name: "verbose",
      description: commands.root.options.verbose,
      default: false,
    },
    {
      type: "boolean",
      name: "telemetry",
      description: commands.root.options.telemetry,
      default: true,
    },
  ],
  telemetry: {
    event: TelemetryEvent.RootCommand,
  },
};
