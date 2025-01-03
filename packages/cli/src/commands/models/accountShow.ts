// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { CLICommand, err, ok } from "@microsoft/teamsfx-api";
import {
  AppStudioScopes,
  AzureScopes,
  featureFlagManager,
  FeatureFlags,
} from "@microsoft/teamsfx-core";
import { TextType, colorize } from "../../colorize";
import AzureTokenProvider, { getAzureProvider } from "../../commonlib/azureLogin";
import AzureTokenCIProvider from "../../commonlib/azureLoginCI";
import { checkIsOnline } from "../../commonlib/codeFlowLogin";
import { signedIn } from "../../commonlib/common/constant";
import { logger } from "../../commonlib/logger";
import M365TokenProvider from "../../commonlib/m365Login";
import { commands, strings } from "../../resource";
import { TelemetryEvent } from "../../telemetry/cliTelemetryEvents";
import { listAllTenants } from "@microsoft/teamsfx-core/build/common/tools";

class AccountUtils {
  outputAccountInfoOffline(accountType: string, username: string): boolean {
    logger.outputInfo(
      strings["account.show.info"],
      accountType,
      colorize(username, TextType.Important)
    );
    return true;
  }

  async outputM365Info(commandType: "login" | "show", tid?: string): Promise<boolean> {
    const appStudioTokenJsonRes = await M365TokenProvider.getJsonObject(
      {
        scopes: AppStudioScopes,
      },
      tid
    );
    const result = appStudioTokenJsonRes.isOk() ? appStudioTokenJsonRes.value : undefined;
    if (result) {
      if (tid) {
        await M365TokenProvider.switchTenant(tid);
      }
      const username = (result as any).upn ?? (result as any).unique_name;
      if (commandType === "login") {
        logger.outputSuccess(strings["account.login.m365"]);
      }

      const cachedTenantId = await M365TokenProvider.getTenant();
      if (cachedTenantId) {
        const listTenantToken = await M365TokenProvider.getAccessToken({ scopes: AzureScopes });
        if (listTenantToken.isOk()) {
          const tenants = await listAllTenants(listTenantToken.value);
          const curTenant = tenants.find((tenant) => tenant.tenantId === cachedTenantId);
          logger.outputInfo(
            strings["account.show.m365.tenant"],
            colorize(username, TextType.Important),
            colorize(curTenant?.displayName, TextType.Important)
          );
        }
      } else {
        logger.outputInfo(strings["account.show.m365"], colorize(username, TextType.Important));
      }
      return Promise.resolve(true);
    } else {
      if (commandType === "login") {
        logger.outputError(strings["account.login.m365.fail"]);
      }
    }
    return Promise.resolve(result !== undefined);
  }

  async outputAzureInfo(
    commandType: "login" | "show",
    tenantId = "",
    isServicePrincipal = false,
    userName = "",
    password = ""
  ): Promise<boolean> {
    let azureProvider = getAzureProvider();
    if (isServicePrincipal === true || (await AzureTokenCIProvider.load())) {
      await AzureTokenCIProvider.init(userName, password, tenantId);
      azureProvider = AzureTokenCIProvider;
    }
    const result = await azureProvider.getJsonObject(true, tenantId);
    if (result) {
      if (tenantId) {
        await azureProvider.switchTenant(tenantId);
      }
      const subscriptions = await azureProvider.listSubscriptions();
      const username = (result as any).upn ?? (result as any).unique_name;
      if (commandType === "login") {
        logger.outputSuccess(strings["account.login.azure"]);
      }

      const cachedTenantId = await azureProvider.getTenant();
      if (cachedTenantId) {
        const identityCredential = await azureProvider.getIdentityCredentialAsync(false);
        const listTenantToken = identityCredential
          ? await identityCredential.getToken(AzureScopes)
          : undefined;
        if (listTenantToken && listTenantToken.token) {
          const tenants = await listAllTenants(listTenantToken.token);
          const curTenant = tenants.find((tenant) => tenant.tenantId === cachedTenantId);
          logger.outputInfo(
            strings["account.show.azure.tenant"],
            colorize(username, TextType.Important),
            colorize(curTenant?.displayName, TextType.Important),
            JSON.stringify(subscriptions, null, 2)
          );
        }
      } else {
        logger.outputInfo(
          strings["account.show.azure"],
          colorize(username, TextType.Important),
          JSON.stringify(subscriptions, null, 2)
        );
      }
      return Promise.resolve(true);
    } else {
      if (commandType === "login") {
        logger.outputError(strings["account.login.azure.fail"]);
      }
    }
    return Promise.resolve(result !== undefined);
  }

  async checkIsOnline(): Promise<boolean> {
    return checkIsOnline();
  }
}

export const accountUtils = new AccountUtils();

export const accountShowCommand: CLICommand = {
  name: "list",
  aliases: ["show"],
  description: commands["auth.show"].description,
  telemetry: {
    event: TelemetryEvent.AccountShow,
  },
  handler: async (ctx) => {
    const m365StatusRes = await M365TokenProvider.getStatus({ scopes: AppStudioScopes });
    if (m365StatusRes.isErr()) {
      return err(m365StatusRes.error);
    }
    const m365Status = m365StatusRes.value;
    if (m365Status.status === signedIn) {
      (await accountUtils.checkIsOnline())
        ? await accountUtils.outputM365Info("show")
        : accountUtils.outputAccountInfoOffline(
            "Microsoft 365",
            (m365Status.accountInfo as any).upn
          );
    }

    const azureStatus = await AzureTokenProvider.getStatus();
    if (azureStatus.status === signedIn) {
      (await accountUtils.checkIsOnline())
        ? await accountUtils.outputAzureInfo("show")
        : accountUtils.outputAccountInfoOffline("Azure", (azureStatus.accountInfo as any).upn);
    }

    if (m365Status.status !== signedIn && azureStatus.status !== signedIn) {
      logger.info(
        `Use \`${process.env.TEAMSFX_CLI_BIN_NAME} auth login azure\` or \`${process.env.TEAMSFX_CLI_BIN_NAME} auth login m365\` to log in to Azure or Microsoft 365 account.`
      );
    }
    return ok(undefined);
  },
};
