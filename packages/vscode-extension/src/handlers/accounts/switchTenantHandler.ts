// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { listAllTenants } from "@microsoft/teamsfx-core/build/common/tools";
import { ExtTelemetry } from "../../telemetry/extTelemetry";
import { AccountType, TelemetryEvent, TelemetryProperty } from "../../telemetry/extTelemetryEvents";
import { getTriggerFromProperty } from "../../utils/telemetryUtils";
import M365TokenInstance from "../../commonlib/m365Login";
import azureAccountManager from "../../commonlib/azureLogin";
import { AzureScopes, isUserCancelError } from "@microsoft/teamsfx-core";
import { FxError, SingleSelectConfig, SystemError } from "@microsoft/teamsfx-api";
import { localize } from "../../utils/localizeUtils";
import { VS_CODE_UI } from "../../qm/vsc_ui";
import { ExtensionSource } from "../../error/error";
import { showError } from "../../error/common";

export const switchTenantHandlerDeps = {
  sendTelemetryEvent: (eventName: string, properties?: any) =>
    ExtTelemetry.sendTelemetryEvent(eventName as any, properties),
  sendTelemetryErrorEvent: (eventName: string, error: FxError, properties?: any) =>
    ExtTelemetry.sendTelemetryErrorEvent(eventName as any, error, properties),
  getTriggerFromProperty: (args?: unknown[]) => getTriggerFromProperty(args),
  m365GetAccessToken: (scopes: string[]) => M365TokenInstance.getAccessToken({ scopes }),
  m365SwitchTenant: (tenantId: string) => M365TokenInstance.switchTenant(tenantId),
  azureGetIdentityCredentialAsync: () => azureAccountManager.getIdentityCredentialAsync(false),
  azureSwitchTenant: (tenantId: string) => azureAccountManager.switchTenant(tenantId),
  listAllTenants: (token: string) => listAllTenants(token),
  selectOption: (config: SingleSelectConfig) => VS_CODE_UI.selectOption(config),
  createProgressBar: (title: string, totalSteps: number) =>
    VS_CODE_UI.createProgressBar(title, totalSteps),
  localize: (key: string, ...args: any[]) => localize(key, ...args),
  showError: (error: FxError) => showError(error),
};

export async function onSwitchM365Tenant(...args: unknown[]): Promise<void> {
  switchTenantHandlerDeps.sendTelemetryEvent(TelemetryEvent.SwitchTenantStart, {
    [TelemetryProperty.AccountType]: AccountType.M365,
    ...switchTenantHandlerDeps.getTriggerFromProperty(args),
  });

  let error: FxError | undefined = undefined;
  const tokenRes = await switchTenantHandlerDeps.m365GetAccessToken(AzureScopes());
  if (tokenRes.isOk()) {
    const config: SingleSelectConfig = {
      name: "SwitchTenant",
      title: switchTenantHandlerDeps.localize("teamstoolkit.handlers.switchtenant.quickpick.title"),
      options: async () => {
        const tenants = await switchTenantHandlerDeps.listAllTenants(tokenRes.value);
        return tenants.map((tenant: any) => {
          return {
            id: tenant.tenantId,
            label: tenant.displayName,
            description: tenant.defaultDomain,
          };
        });
      },
    };
    const result = await switchTenantHandlerDeps.selectOption(config);
    if (result.isOk()) {
      const progressHandler = switchTenantHandlerDeps.createProgressBar(
        switchTenantHandlerDeps.localize("teamstoolkit.commands.switchTenant.progressbar.title"),
        1
      );
      await progressHandler.start();
      await progressHandler.next(
        switchTenantHandlerDeps.localize("teamstoolkit.commands.switchTenant.progressbar.detail")
      );
      const switchRes = await switchTenantHandlerDeps.m365SwitchTenant(
        result.value.result as string
      );
      await progressHandler.end(switchRes.isOk());
      if (switchRes.isOk()) {
        switchTenantHandlerDeps.sendTelemetryEvent(TelemetryEvent.SwitchTenant, {
          [TelemetryProperty.AccountType]: AccountType.M365,
          [TelemetryProperty.TenantId]: result.value.result as string,
          ...switchTenantHandlerDeps.getTriggerFromProperty(args),
        });
        return;
      } else {
        error = switchRes.error;
      }
    } else {
      error = result.error;
    }
  } else {
    error = tokenRes.error;
  }

  if (!isUserCancelError(error)) {
    void switchTenantHandlerDeps.showError(error);
  }
  switchTenantHandlerDeps.sendTelemetryErrorEvent(TelemetryEvent.SwitchTenant, error, {
    [TelemetryProperty.AccountType]: AccountType.M365,
    ...switchTenantHandlerDeps.getTriggerFromProperty(args),
  });
}

export async function onSwitchAzureTenant(...args: unknown[]): Promise<void> {
  switchTenantHandlerDeps.sendTelemetryEvent(TelemetryEvent.SwitchTenantStart, {
    [TelemetryProperty.AccountType]: AccountType.Azure,
    ...switchTenantHandlerDeps.getTriggerFromProperty(args),
  });

  const config: SingleSelectConfig = {
    name: "SwitchTenant",
    title: switchTenantHandlerDeps.localize("teamstoolkit.handlers.switchtenant.quickpick.title"),
    options: async () => {
      const tokenCredential = await switchTenantHandlerDeps.azureGetIdentityCredentialAsync();
      const token = tokenCredential ? await tokenCredential.getToken(AzureScopes()) : undefined;
      if (token && token.token) {
        const tenants = await switchTenantHandlerDeps.listAllTenants(token.token);
        return tenants.map((tenant: any) => {
          return {
            id: tenant.tenantId,
            label: tenant.displayName,
            description: tenant.defaultDomain,
          };
        });
      } else {
        throw new SystemError(
          ExtensionSource,
          "SwitchTenantFailed",
          switchTenantHandlerDeps.localize("teamstoolkit.handlers.switchtenant.error")
        );
      }
    },
  };
  const result = await switchTenantHandlerDeps.selectOption(config);
  let error: any;
  if (result.isOk()) {
    const progressHandler = switchTenantHandlerDeps.createProgressBar(
      switchTenantHandlerDeps.localize("teamstoolkit.commands.switchTenant.progressbar.title"),
      1
    );
    await progressHandler.start();
    await progressHandler.next(
      switchTenantHandlerDeps.localize("teamstoolkit.commands.switchTenant.progressbar.detail")
    );
    const switchRes = await switchTenantHandlerDeps.azureSwitchTenant(
      result.value.result as string
    );
    await progressHandler.end(switchRes.isOk());
    if (switchRes.isOk()) {
      switchTenantHandlerDeps.sendTelemetryEvent(TelemetryEvent.SwitchTenant, {
        [TelemetryProperty.AccountType]: AccountType.Azure,
        [TelemetryProperty.TenantId]: result.value.result as string,
        ...switchTenantHandlerDeps.getTriggerFromProperty(args),
      });
      return;
    } else {
      error = switchRes.error;
    }
  } else {
    error = result.error;
  }

  if (!isUserCancelError(error)) {
    void switchTenantHandlerDeps.showError(error);
  }
  switchTenantHandlerDeps.sendTelemetryErrorEvent(TelemetryEvent.SwitchTenant, error, {
    [TelemetryProperty.AccountType]: AccountType.Azure,
    ...switchTenantHandlerDeps.getTriggerFromProperty(args),
  });
}
