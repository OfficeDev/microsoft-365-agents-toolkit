// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { Project } from "../utils/constants";
import { Env } from "../utils/env";
import {
  AppStudioCleanHelper,
  filterResourceGroupByName,
  deleteResourceGroupByName,
  GraphApiCleanHelper,
  SharePointApiCleanHelper,
  DevTunnelCleanHelper,
  M365TitleCleanHelper,
} from "../utils/cleanHelper";
import { ResourceGroupManager } from "../utils/resourceGroupManager";
import { getAppNamePrefix } from "../utils/nameUtil";
import { delay } from "../utils/retryHandler";

// const appStudioAppNamePrefixList: string[] = [Project.namePrefix, "vs"];
const appNamePrefixList: string[] = [Project.namePrefix, "vs", "fx"];
// const aadNamePrefixList: string[] = [Project.namePrefix, "vs"];
const rgNamePrefixList: string[] = [Project.namePrefix, "vs"];
const adminMicrosoftEntraAppName = [
  "delete-client",
  "ATK Test SP",
  "TravelAgent-AADlocal",
  "Test SP",
];
const excludePrefix: string = getAppNamePrefix();

function parseDate(input: unknown): Date | undefined {
  if (!input) {
    return undefined;
  }

  if (input instanceof Date) {
    return Number.isNaN(input.getTime()) ? undefined : input;
  }

  if (typeof input === "string" || typeof input === "number") {
    const parsed = new Date(input);
    return Number.isNaN(parsed.getTime()) ? undefined : parsed;
  }

  return undefined;
}

function getUtcDateKey(date: Date): string {
  return `${date.getUTCFullYear()}-${date.getUTCMonth() + 1}-${date.getUTCDate()}`;
}

function isCreatedToday(createdDate: unknown): boolean {
  const created = parseDate(createdDate);
  if (!created) {
    return false;
  }
  return getUtcDateKey(created) === getUtcDateKey(new Date());
}

function getRgCreatedDate(resourceGroup: unknown): unknown {
  const rg = resourceGroup as {
    systemData?: { createdAt?: string };
    createdTime?: string;
    properties?: { createdTime?: string };
  };

  return (
    rg.systemData?.createdAt ?? rg.createdTime ?? rg.properties?.createdTime
  );
}

function getTunnelCreatedDate(tunnel: unknown): unknown {
  const tunnelInfo = tunnel as {
    created?: string;
    createdAt?: string;
    creationTime?: string;
    createdTime?: string;
    status?: { created?: string; createdAt?: string; createdTime?: string };
  };

  return (
    tunnelInfo.created ??
    tunnelInfo.createdAt ??
    tunnelInfo.creationTime ??
    tunnelInfo.createdTime ??
    tunnelInfo.status?.created ??
    tunnelInfo.status?.createdAt ??
    tunnelInfo.status?.createdTime
  );
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  return String(error);
}

function shouldSkipAadApp(displayName?: string): boolean {
  if (!displayName) {
    return true;
  }

  return (
    adminMicrosoftEntraAppName.some((name) => displayName.startsWith(name)) ||
    displayName.startsWith(excludePrefix)
  );
}

function shouldCleanByPrefix(displayName?: string): boolean {
  if (!displayName) {
    return false;
  }

  return (
    appNamePrefixList.some((name) => displayName.startsWith(name)) &&
    !displayName.startsWith(excludePrefix)
  );
}

async function main() {
  const cleanService = await GraphApiCleanHelper.create(
    Env.cleanTenantId,
    Env.cleanClientId,
    Env.username,
    Env.password,
  );

  try {
    console.log(`clean teams app (exclude ${excludePrefix})`);
    const teamsUserId = await cleanService.getUserIdByName(Env.username);
    const teamsAppList = await cleanService.listTeamsApp(teamsUserId);
    if (teamsAppList) {
      for (const app of teamsAppList) {
        const displayName = app?.teamsAppDefinition?.displayName;
        const createdDate = app?.teamsAppDefinition?.createdDateTime;
        if (isCreatedToday(createdDate)) {
          continue;
        }
        if (shouldCleanByPrefix(displayName)) {
          console.log(displayName);
          try {
            await cleanService.uninstallTeamsApp(teamsUserId, app?.id ?? "");
          } catch {
            console.log(`Failed to uninstall Teams App ${displayName}`);
          }
        }
      }
    }
  } catch {
    console.log(`Failed to clean teams app`);
  }

  try {
    console.log(`audit Entra objects related to test artifacts`);
    const aadList = await cleanService.listAad();
    const enterpriseAppList = await cleanService.listEnterpriseApplications();
    const deletedAadList = await cleanService.listDeletedAad();
    const deletedEnterpriseAppList =
      await cleanService.listDeletedEnterpriseApplications();

    console.log(`active app registrations: ${aadList.length}`);
    console.log(`active enterprise applications: ${enterpriseAppList.length}`);
    console.log(`deleted app registrations: ${deletedAadList.length}`);
    console.log(
      `deleted enterprise applications: ${deletedEnterpriseAppList.length}`,
    );
  } catch (e: unknown) {
    console.log(`Failed to audit Entra objects, ${getErrorMessage(e)}`);
  }

  try {
    console.log(`clean AAD (exclude ${excludePrefix})`);
    const aadList = await cleanService.listAad();
    if (aadList) {
      for (const aad of aadList) {
        if (isCreatedToday(aad?.createdDateTime)) {
          continue;
        }
        if (!shouldSkipAadApp(aad.displayName)) {
          console.log(aad.displayName);
          try {
            await cleanService.deleteAad(aad.id!);
          } catch (e: unknown) {
            console.log(
              `Failed to delete AAD ${aad.displayName} with error: ${getErrorMessage(e)}`,
            );
          }
        }
      }
    }
  } catch {
    console.log(`Failed to clean AAD`);
  }

  try {
    console.log(`clean enterprise applications (exclude ${excludePrefix})`);
    const enterpriseAppList = await cleanService.listEnterpriseApplications();
    if (enterpriseAppList) {
      for (const enterpriseApp of enterpriseAppList) {
        if (isCreatedToday(enterpriseApp?.createdDateTime)) {
          continue;
        }
        if (shouldCleanByPrefix(enterpriseApp.displayName)) {
          console.log(enterpriseApp.displayName);
          try {
            await cleanService.deleteEnterpriseApplication(enterpriseApp.id!);
          } catch (e: unknown) {
            console.log(
              `Failed to delete enterprise application ${enterpriseApp.displayName} with error: ${getErrorMessage(e)}`,
            );
          }
        }
      }
    }
  } catch {
    console.log(`Failed to clean enterprise applications`);
  }

  try {
    console.log(
      `purge deleted AAD app registrations (exclude ${excludePrefix})`,
    );
    const deletedAadList = await cleanService.listDeletedAad();
    if (deletedAadList) {
      for (const aad of deletedAadList) {
        if (isCreatedToday(aad?.createdDateTime)) {
          continue;
        }
        if (!shouldSkipAadApp(aad.displayName)) {
          console.log(aad.displayName);
          try {
            await cleanService.deleteDeletedItem(aad.id!);
          } catch (e: unknown) {
            console.log(
              `Failed to purge deleted AAD ${aad.displayName} with error: ${getErrorMessage(e)}`,
            );
          }
        }
      }
    }
  } catch {
    console.log(`Failed to purge deleted AAD app registrations`);
  }

  try {
    console.log(
      `purge deleted enterprise applications (exclude ${excludePrefix})`,
    );
    const deletedEnterpriseAppList =
      await cleanService.listDeletedEnterpriseApplications();
    if (deletedEnterpriseAppList) {
      for (const enterpriseApp of deletedEnterpriseAppList) {
        if (isCreatedToday(enterpriseApp?.createdDateTime)) {
          continue;
        }
        if (shouldCleanByPrefix(enterpriseApp.displayName)) {
          console.log(enterpriseApp.displayName);
          try {
            await cleanService.deleteDeletedItem(enterpriseApp.id!);
          } catch (e: unknown) {
            console.log(
              `Failed to purge deleted enterprise application ${enterpriseApp.displayName} with error: ${getErrorMessage(e)}`,
            );
          }
        }
      }
    }
  } catch {
    console.log(`Failed to purge deleted enterprise applications`);
  }

  try {
    console.log(`clean app in app studio`);
    const addStudioCleanService = await AppStudioCleanHelper.create(
      Env.cleanTenantId,
      Env.cleanClientId,
      Env.username,
      Env.password,
    );
    const appStudioAppList = await addStudioCleanService.getAppsInAppStudio();
    if (appStudioAppList) {
      for (const app of appStudioAppList) {
        const createdDate =
          app?.createdDateTime ?? app?.createdAt ?? app?.createdTime;
        if (isCreatedToday(createdDate)) {
          continue;
        }
        if (!app?.displayName?.startsWith(excludePrefix)) {
          console.log(app?.displayName);
          try {
            await addStudioCleanService.deleteAppInAppStudio(
              app?.appDefinitionId,
            );
          } catch {
            console.log(
              `Failed to delete Teams App ${app?.displayName} in App Studio`,
            );
          }
        }
      }
    }

    console.log(`clean api key registration`);
    const apiKeyRegistrationList =
      await addStudioCleanService.getApiKeyRegistration();
    if (apiKeyRegistrationList) {
      for (const apiKey of apiKeyRegistrationList) {
        const createdDate =
          apiKey?.createdDateTime ?? apiKey?.createdAt ?? apiKey?.createdTime;
        if (isCreatedToday(createdDate)) {
          continue;
        }
        try {
          await addStudioCleanService.deleteApiKeyRegistration(apiKey?.id);
          console.log(apiKey?.id, " is deleted");
        } catch {
          console.log(`Failed to delete api key ${apiKey?.id}`);
        }
      }
    }
  } catch {
    console.log(`Failed to clean app in app studio`);
  }

  try {
    console.log(
      `clean up the Azure resource group with name start with ${Project.namePrefix} (exclude ${excludePrefix}, skip same-day created)`,
    );
    const rgManager = await ResourceGroupManager.init();
    const rgNameList: string[] = [];
    for (const name of rgNamePrefixList) {
      const group = await filterResourceGroupByName(name);
      group.map((rgName) => rgNameList.push(rgName));
    }
    if (rgNameList.length > 0) {
      for (const rgName of rgNameList) {
        const rg = await rgManager
          .getResourceGroup(rgName)
          .catch(() => undefined);
        if (isCreatedToday(getRgCreatedDate(rg))) {
          continue;
        }
        for (const name of rgNamePrefixList) {
          if (rgName.startsWith(name) && !rgName.startsWith(excludePrefix)) {
            await deleteResourceGroupByName(rgName);
          }
        }
      }
    }
  } catch {
    console.log(
      `Failed to clean up the Azure resource group with name start with ${Project.namePrefix} (exclude ${excludePrefix}, skip same-day created)`,
    );
  }

  try {
    console.log(`clean SharePoint app package files`);
    const sharePointCleanService = await SharePointApiCleanHelper.create(
      Env.cleanTenantId,
      Env.cleanClientId,
      Env.username,
      Env.password,
    );
    const sharePointAppList = await sharePointCleanService.listApp();
    if (sharePointAppList) {
      for (const app of sharePointAppList) {
        const createdDate =
          app?.Created ?? app?.TimeCreated ?? app?.createdDateTime;
        if (isCreatedToday(createdDate)) {
          continue;
        }
        for (const name of appNamePrefixList) {
          if (
            app.Title?.startsWith(name) &&
            !app.Title?.startsWith(excludePrefix)
          ) {
            console.log(app.Title);
            try {
              await sharePointCleanService.deleteApp(app.ID!);
            } catch {
              console.log(`Failed to delete SharePoint app ${app.ID!}`);
            }
          }
        }
      }
    }
  } catch (e: unknown) {
    console.log(
      `Failed to clean up SharePoint app package files, ${getErrorMessage(e)}`,
    );
  }

  try {
    console.log(`clean dev tunnel`);
    const devTunnelCleanHelper = await DevTunnelCleanHelper.create(
      Env.cleanTenantId,
      Env.username,
      Env.password,
    );
    const tunnels = await devTunnelCleanHelper.listTunnels();
    for (const tunnel of tunnels) {
      if (isCreatedToday(getTunnelCreatedDate(tunnel))) {
        continue;
      }

      console.log(`clean dev tunnel ${tunnel?.tunnelId}`);
      await devTunnelCleanHelper.deleteByTunnel(tunnel);
    }
  } catch {
    console.log(`Failed to clean dev tunnel`);
  }

  let retry: boolean;
  let count = 10;
  const total = count + 1;
  do {
    retry = false;
    console.log(`Start to try ${total - count} times`);
    const m365TitleCleanService = await M365TitleCleanHelper.create(
      Env.cleanTenantId,
      "7ea7c24c-b1f6-4a20-9d11-9ae12e9e7ac0",
      Env.username,
      Env.password,
    );
    console.log(`clean M365 Titles (exclude ${excludePrefix})`);
    try {
      const acquisitions = await m365TitleCleanService.listAcquisitions();
      if (acquisitions) {
        for (const acquisition of acquisitions) {
          const createdDate =
            acquisition?.createdDateTime ?? acquisition?.acquiredDateTime;
          if (isCreatedToday(createdDate)) {
            continue;
          }
          if (!acquisition.titleDefinition.name.startsWith(excludePrefix)) {
            console.log(acquisition.titleDefinition.name);
            console.log(acquisition.titleId);
            const result = await m365TitleCleanService.unacquire(
              acquisition.titleId,
              1,
            );
            if (!retry && result) {
              retry = true;
            }
          }
        }
      }
    } catch (e: unknown) {
      console.log(`Get error: ${getErrorMessage(e)}`);
      retry = true;
      if (count > 1) {
        // Retry after a short time if getting "Rate limit is exceeded"
        await delay(30 * 1000);
      }
    }

    count--;
  } while (retry && count > 0);
}

main()
  .then(() => {
    console.log("Clean Job Done.");
  })
  .catch((error) => {
    console.error(error);
    process.exit(-1);
  });
