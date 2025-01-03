// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, UserError, err, ok } from "@microsoft/teamsfx-api";
import AzureTokenProvider from "../../commonlib/azureLogin";
import {
  codeFlowLoginFormat,
  loginComponent,
  servicePrincipalLoginFormat,
  usageError,
} from "../../commonlib/common/constant";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { accountUtils } from "./accountShow";
import { featureFlagManager, FeatureFlags } from "@microsoft/teamsfx-core";

export const accountLoginAzureCommand: CLICommand = {
  name: "azure",
  description: commands["auth.login.azure"].description,
  options: [
    {
      name: "tenant",
      description: commands["auth.login.azure"].options["tenant"],
      type: "string",
      default: "",
    },
    {
      name: "service-principal",
      description: commands["auth.login.azure"].options["service-principal"],
      type: "boolean",
      default: false,
    },
    {
      name: "username",
      shortName: "u",
      description: commands["auth.login.azure"].options.username,
      type: "string",
      default: "",
    },
    {
      name: "password",
      shortName: "p",
      description: commands["auth.login.azure"].options.password,
      type: "string",
      default: "",
    },
  ],
  examples: [
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} auth login azure --interactive false --service-principal -u USERNAME  -p SECRET --tenant TENANT_ID`,
      description: "Log in with a service principal using client secret",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} auth login azure --interactive false --service-principal -u USERNAME  -p "C:/Users/mycertfile.pem" --tenant TENANT_ID`,
      description: "Log in with a service principal using client certificate",
    },
  ],
  telemetry: {
    event: TelemetryEvent.AccountLoginAzure,
  },
  handler: async (ctx) => {
    const args = ctx.optionValues;
    const isSP = args["service-principal"] as boolean;
    if (isSP === true) {
      if (!args.username || !args.password || !args.tenant) {
        return err(new UserError(loginComponent, usageError, servicePrincipalLoginFormat));
      }
    } else {
      if (args.username || args.password) {
        return err(new UserError(loginComponent, usageError, codeFlowLoginFormat));
      }
    }
    await AzureTokenProvider.signout();
    await accountUtils.outputAzureInfo(
      "login",
      args.tenant as string,
      isSP,
      args.username as string,
      args.password as string
    );
    return ok(undefined);
  },
};
