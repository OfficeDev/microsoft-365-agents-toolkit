// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Command } from "commander";
import { signedIn } from "@microsoft/teamsfx-core-next";
import { wrapHandler } from "../handler";
import { M365Login, AzureAccountManager, AzureLoginCI } from "../auth";
import { checkAzureSPFile } from "../auth/cacheAccess";

/**
 * Account/auth commands: auth login, auth logout, auth show
 */
export function createAccountCommands(program: Command): void {
  const auth = program
    .command("auth")
    .alias("account")
    .description("Manage Microsoft 365 and Azure accounts");

  // atk auth show
  auth
    .command("show")
    .description("Display currently logged-in accounts")
    .action(
      wrapHandler("auth show", async (_opts, _cmd) => {
        const m365 = M365Login.getInstance();
        const m365Status = await m365.getStatus({
          scopes: ["https://dev.teams.microsoft.com/.default"],
        });
        if (m365Status.isOk()) {
          const s = m365Status.value;
          if (s.status === signedIn) {
            const upn = (s.accountInfo as Record<string, string>)?.upn ?? "unknown";
            console.log(`Microsoft 365: Signed in as ${upn}`);
          } else {
            console.log("Microsoft 365: Not signed in");
          }
        }

        const azureProvider = checkAzureSPFile()
          ? AzureLoginCI.getInstance()
          : AzureAccountManager.getInstance();
        const azureStatus = await azureProvider.getStatus();
        if (azureStatus.status === signedIn) {
          const info = azureProvider.getAccountInfo();
          console.log(`Azure: Signed in${info?.username ? ` as ${info.username}` : ""}`);
        } else {
          console.log("Azure: Not signed in");
        }
      })
    );

  // atk auth login
  const login = auth.command("login").description("Log in to Azure or Microsoft 365");

  // atk auth login azure
  login
    .command("azure")
    .description("Log in to Azure")
    .option("--tenant <tenantId>", "Azure AD tenant ID or domain")
    .option("--service-principal", "Log in with a service principal", false)
    .option("--username <username>", "Service principal client ID")
    .option("--password <password>", "Service principal secret or certificate path")
    .action(
      wrapHandler("auth login azure", async (opts, _cmd) => {
        if (opts.servicePrincipal) {
          const clientId = opts.username as string;
          const secret = opts.password as string;
          const tenant = opts.tenant as string;
          if (!clientId || !secret || !tenant) {
            throw new Error(
              "Service principal login requires --username, --password, and --tenant."
            );
          }
          const ci = AzureLoginCI.getInstance();
          await ci.init(clientId, secret, tenant);
          console.log("Azure: Signed in with service principal.");
        } else {
          const mgr = AzureAccountManager.getInstance();
          // Trigger interactive login by requesting a credential
          await mgr.getIdentityCredentialAsync(true);
          const status = await mgr.getStatus();
          if (status.status === signedIn) {
            console.log("Azure: Signed in successfully.");
          }
        }
      })
    );

  // atk auth login m365
  login
    .command("m365")
    .description("Log in to Microsoft 365")
    .option("--tenant <tenantId>", "Microsoft 365 tenant ID or domain")
    .action(
      wrapHandler("auth login m365", async (_opts, _cmd) => {
        const m365 = M365Login.getInstance();
        // Trigger interactive login by requesting a token
        const result = await m365.getAccessToken({
          scopes: ["https://dev.teams.microsoft.com/.default"],
        });
        if (result.isOk()) {
          console.log("Microsoft 365: Signed in successfully.");
        } else {
          throw new Error(`M365 login failed: ${result.error.message}`);
        }
      })
    );

  // atk auth logout
  auth
    .command("logout")
    .description("Log out of Azure and Microsoft 365")
    .action(
      wrapHandler("auth logout", async (_opts, _cmd) => {
        await M365Login.getInstance().signout();
        console.log("Microsoft 365: Signed out.");

        const azureProvider = checkAzureSPFile()
          ? AzureLoginCI.getInstance()
          : AzureAccountManager.getInstance();
        await azureProvider.signout();
        console.log("Azure: Signed out.");
      })
    );
}
