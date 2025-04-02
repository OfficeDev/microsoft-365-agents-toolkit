// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, err, ok } from "@microsoft/teamsfx-api";
import { PackageService, MosServiceEndpoint, MosServiceScope } from "@microsoft/teamsfx-core";
import { logger } from "../../commonlib/logger";
import M365TokenProvider from "../../commonlib/m365Login";
import { ArgumentConflictError, MissingRequiredOptionError } from "../../error";
import { commands } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { AppScope } from "@microsoft/teamsfx-core/build/component/m365/packageService";

export const sideloadingServiceEndpoint =
  process.env.SIDELOADING_SERVICE_ENDPOINT ?? MosServiceEndpoint;
export const sideloadingServiceScope = process.env.SIDELOADING_SERVICE_SCOPE ?? MosServiceScope;

class M365Utils {
  async getTokenAndUpn(): Promise<[string, string]> {
    const tokenRes = await M365TokenProvider.getAccessToken({ scopes: [sideloadingServiceScope] });
    if (tokenRes.isErr()) {
      logger.error(
        `Cannot get token. Use '${process.env.TEAMSFX_CLI_BIN_NAME} account login m365' to log in the correct account.`
      );
      throw tokenRes.error;
    } else {
      let upn = undefined;
      try {
        // workaround to get upn via appstudio scope.
        const accountRes = await M365TokenProvider.getStatus({
          scopes: ["https://dev.teams.microsoft.com/AppDefinitions.ReadWrite"],
        });
        if (accountRes.isOk()) {
          upn = (accountRes.value.accountInfo as any).upn;
        } else {
          throw accountRes.error;
        }
      } catch (error) {
        logger.debug(`Failed to get upn. Error: ${JSON.stringify(error)}`);
      }
      if (upn !== undefined) {
        logger.info(`Using account ${upn}`);
      }
      const token = tokenRes.value;
      return [token, upn];
    }
  }
}

export const m365utils = new M365Utils();

export const m365SideloadingCommand: CLICommand = {
  name: "install",
  aliases: ["sideloading"],
  description: commands.install.description,
  options: [
    {
      name: "file-path",
      description: commands.install.options["file-path"],
      type: "string",
    },
    {
      name: "xml-path",
      description: commands.install.options["xml-path"],
      type: "string",
    },
    {
      name: "scope",
      description: commands.install.options["scope"],
      type: "string",
    },
  ],
  examples: [
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} install --file-path appPackage.zip`,
      description:
        "Sideload the application package with JSON-based manifest to Teams, Outlook, and the Microsoft 365 app.",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} install --file-path appPackage.zip --scope Shared`,
      description:
        "Sideload the application package in Shared scope with JSON-based manifest to Teams, Outlook, and the Microsoft 365 app.",
    },
    {
      command: `${process.env.TEAMSFX_CLI_BIN_NAME} install --xml-path manifest.xml`,
      description:
        "Sideload the Outlook add-in application package with XML-based manifest to Outlook.",
    },
  ],
  telemetry: {
    event: TelemetryEvent.Install,
  },
  defaultInteractiveOption: false,
  handler: async (ctx) => {
    const zipAppPackagePath = ctx.optionValues["file-path"] as string;
    const xmlPath = ctx.optionValues["xml-path"] as string;
    let appScope = parseAppScope(ctx.optionValues["scope"] as string | undefined);
    if (!appScope) {
      appScope = AppScope.Personal;
    }
    if (zipAppPackagePath === undefined && xmlPath === undefined) {
      return err(new MissingRequiredOptionError(ctx.command.fullName, `--file-path or --xml-path`));
    }

    if (zipAppPackagePath !== undefined && xmlPath !== undefined) {
      return err(new ArgumentConflictError(ctx.command.fullName, `--file-path`, `--xml-path`));
    }

    const packageService = new PackageService(sideloadingServiceEndpoint, logger);
    const manifestPath = zipAppPackagePath ?? xmlPath;
    const tokenAndUpn = await m365utils.getTokenAndUpn();
    if (ctx.optionValues["file-path"] !== undefined) {
      await packageService.sideLoading(tokenAndUpn[0], manifestPath, appScope);
    } else {
      await packageService.sideLoadXmlManifest(tokenAndUpn[0], manifestPath);
    }
    return ok(undefined);
  },
};

export function parseAppScope(scope?: string): AppScope {
  if (!scope) {
    return AppScope.Personal;
  }
  switch (scope.toLowerCase()) {
    case AppScope.Shared.toLowerCase():
      return AppScope.Shared;
    case AppScope.Personal.toLowerCase():
      return AppScope.Personal;
    default:
      return AppScope.Personal;
  }
}
