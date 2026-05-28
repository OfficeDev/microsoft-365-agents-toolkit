// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { QuickPickItem, window } from "vscode";
import { FxError, OptionItem, Result, SingleSelectConfig, ok } from "@microsoft/teamsfx-api";
import { Correlator, AppStudioScopes } from "@microsoft/teamsfx-core";
import { ExtTelemetry } from "../../telemetry/extTelemetry";
import { AccountType, TelemetryEvent, TelemetryProperty } from "../../telemetry/extTelemetryEvents";
import { signInAzure, signOutAzure, signInM365, signOutM365 } from "../../utils/accountUtils";
import { localize } from "../../utils/localizeUtils";
import { getTriggerFromProperty } from "../../utils/telemetryUtils";
import azureAccountManager from "../../commonlib/azureLogin";
import M365TokenInstance from "../../commonlib/m365Login";
import { VS_CODE_UI } from "../../qm/vsc_ui";

export interface VscQuickPickItem extends QuickPickItem {
  /**
   * Current id of the option item.
   */
  id: string;
  function: () => Promise<void>;
}

export async function createAccountHandler(args: any[]): Promise<void> {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CreateAccountStart, getTriggerFromProperty(args));
  const m365Option: OptionItem = {
    id: "createAccountM365",
    label: `$(add) ${localize("teamstoolkit.commands.createAccount.m365")}`,
    description: localize("teamstoolkit.commands.createAccount.requireSubscription"),
  };
  const azureOption: OptionItem = {
    id: "createAccountAzure",
    label: `$(add) ${localize("teamstoolkit.commands.createAccount.azure")}`,
    description: localize("teamstoolkit.commands.createAccount.free"),
  };
  const option: SingleSelectConfig = {
    name: "CreateAccounts",
    title: localize("teamstoolkit.commands.createAccount.title"),
    options: [m365Option, azureOption],
  };
  const result = await VS_CODE_UI.selectOption(option);
  if (result.isOk()) {
    if (result.value.result === m365Option.id) {
      await VS_CODE_UI.openUrl("https://developer.microsoft.com/microsoft-365/dev-program");
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CreateAccount, {
        [TelemetryProperty.AccountType]: AccountType.M365,
        ...getTriggerFromProperty(args),
      });
    } else if (result.value.result === azureOption.id) {
      await VS_CODE_UI.openUrl("https://azure.microsoft.com/en-us/free/");
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CreateAccount, {
        [TelemetryProperty.AccountType]: AccountType.Azure,
        ...getTriggerFromProperty(args),
      });
    }
  } else {
    ExtTelemetry.sendTelemetryErrorEvent(TelemetryEvent.CreateAccount, result.error, {
      ...getTriggerFromProperty(args),
    });
  }
  return;
}

export async function cmpAccountsHandler(args: any[]) {
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.ManageAccount, getTriggerFromProperty(args));
  const signInAzureOption: VscQuickPickItem = {
    id: "signInAzure",
    label: localize("teamstoolkit.handlers.signInAzure"),
    function: () => signInAzure(),
  };

  const signOutAzureOption: VscQuickPickItem = {
    id: "signOutAzure",
    label: localize("teamstoolkit.handlers.signOutOfAzure"),
    function: async () =>
      await Correlator.run(async () => {
        await signOutAzure(false);
      }),
  };

  const signInM365Option: VscQuickPickItem = {
    id: "signinM365",
    label: localize("teamstoolkit.handlers.signIn365"),
    function: () => signInM365(),
  };

  const signOutM365Option: VscQuickPickItem = {
    id: "signOutM365",
    label: localize("teamstoolkit.handlers.signOutOfM365"),
    function: async () =>
      await Correlator.run(async () => {
        await signOutM365(false);
      }),
  };

  const createAccountsOption: VscQuickPickItem = {
    id: "createAccounts",
    label: `$(add) ${localize("teamstoolkit.commands.createAccount.title")}`,
    function: async () => {
      await Correlator.run(() => createAccountHandler([]));
    },
  };

  const quickPick = window.createQuickPick();
  const quickItemOptionArray: VscQuickPickItem[] = [];

  const m365AccountRes = await M365TokenInstance.getStatus({ scopes: AppStudioScopes });
  const m365Account = m365AccountRes.isOk() ? m365AccountRes.value : undefined;
  if (m365Account && m365Account.status === "SignedIn") {
    const accountInfo = m365Account.accountInfo;
    const email = (accountInfo as any).upn ?? (accountInfo as any).unique_name ?? undefined;
    if (email !== undefined) {
      signOutM365Option.label = signOutM365Option.label.concat(email);
    }
    quickItemOptionArray.push(signOutM365Option);
  } else {
    quickItemOptionArray.push(signInM365Option);
  }

  const azureAccount = await azureAccountManager.getStatus();
  if (azureAccount.status === "SignedIn") {
    const accountInfo = azureAccount.accountInfo;
    const email = (accountInfo as any).email || (accountInfo as any).upn;
    if (email !== undefined) {
      signOutAzureOption.label = signOutAzureOption.label.concat(email);
    }
    quickItemOptionArray.push(signOutAzureOption);
  } else {
    quickItemOptionArray.push(signInAzureOption);
  }

  quickItemOptionArray.push(createAccountsOption);
  quickPick.items = quickItemOptionArray;
  quickPick.onDidChangeSelection((selection) => {
    if (selection[0]) {
      (selection[0] as VscQuickPickItem).function().catch(console.error);
      quickPick.hide();
    }
  });
  quickPick.onDidHide(() => quickPick.dispose());
  quickPick.show();
}

export async function azureAccountSignOutHelpHandler(
  args?: any[]
): Promise<Result<boolean, FxError>> {
  return Promise.resolve(ok(false));
}
