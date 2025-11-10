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
];
const excludePrefix: string = getAppNamePrefix();

async function main() {
  const cleanService = await GraphApiCleanHelper.create(
    Env.cleanTenantId,
    Env.cleanClientId,
    Env.username,
    Env.password
  );

  console.log(`clean teams app (exclude ${excludePrefix})`);
  const teamsUserId = await cleanService.getUserIdByName(Env.username);
  const teamsAppList = await cleanService.listTeamsApp(teamsUserId);
  if (teamsAppList) {
    for (const app of teamsAppList) {
      if (!app?.teamsAppDefinition?.displayName?.startsWith(excludePrefix)) {
        console.log(app?.teamsAppDefinition?.displayName);
        try {
          await cleanService.uninstallTeamsApp(teamsUserId, app?.id ?? "");
        } catch {
          console.log(
            `Failed to uninstall Teams App ${app?.teamsAppDefinition?.displayName}`
          );
        }
      }
    }
  }

  console.log(`clean AAD (exclude ${excludePrefix})`);
  const aadList = await cleanService.listAad();
  if (aadList) {
    for (const aad of aadList) {
      if (
        !adminMicrosoftEntraAppName.some((name) =>
          aad.displayName?.startsWith(name)
        ) &&
        !aad.displayName?.startsWith(excludePrefix)
      ) {
        console.log(aad.displayName);
        try {
          await cleanService.deleteAad(aad.id!);
        } catch (e: any) {
          console.log(
            `Failed to delete AAD ${aad.displayName} with error: ${e.message}`
          );
        }
      }
    }
  }

  console.log(`clean Enterprise Application (exclude ${excludePrefix})`);
  const spList = await cleanService.listEnterpriseApplications();
  if (spList) {
    for (const sp of spList) {
      if (
        !adminMicrosoftEntraAppName.some((name) =>
          sp.displayName?.startsWith(name)
        ) &&
        !sp.displayName?.startsWith(excludePrefix)
      ) {
        console.log(sp.displayName);
        try {
          await cleanService.deleteEnterpriseApplication(sp.id!);
        } catch (e: any) {
          console.log(
            `Failed to delete Enterprise Application ${sp.displayName} with error: ${e.message}`
          );
        }
      }
    }
  }

  console.log(`clean app in app studio`);
  const addStudioCleanService = await AppStudioCleanHelper.create(
    Env.cleanTenantId,
    Env.cleanClientId,
    Env.username,
    Env.password
  );
  const appStudioAppList = await addStudioCleanService.getAppsInAppStudio();
  if (appStudioAppList) {
    for (const app of appStudioAppList) {
      if (!app?.displayName?.startsWith(excludePrefix)) {
        console.log(app?.displayName);
        try {
          await addStudioCleanService.deleteAppInAppStudio(
            app?.appDefinitionId
          );
        } catch {
          console.log(
            `Failed to delete Teams App ${app?.displayName} in App Studio`
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
      try {
        await addStudioCleanService.deleteApiKeyRegistration(apiKey?.id);
        console.log(apiKey?.id, " is deleted");
      } catch {
        console.log(`Failed to delete api key ${apiKey?.id}`);
      }
    }
  }

  console.log(
    `clean up the Azure resource group with name start with ${Project.namePrefix} (exclude ${excludePrefix})`
  );
  const rgNameList: string[] = [];
  for (const name of rgNamePrefixList) {
    const group = await filterResourceGroupByName(name);
    group.map((rgName) => rgNameList.push(rgName));
  }
  if (rgNameList.length > 0) {
    for (const rgName of rgNameList) {
      for (const name of rgNamePrefixList) {
        if (rgName.startsWith(name) && !rgName.startsWith(excludePrefix)) {
          await deleteResourceGroupByName(rgName);
        }
      }
    }
  }

  console.log(`clean SharePoint app package files`);
  const sharePointCleanService = await SharePointApiCleanHelper.create(
    Env.cleanTenantId,
    Env.cleanClientId,
    Env.username,
    Env.password
  );
  const sharePointAppList = await sharePointCleanService.listApp();
  if (sharePointAppList) {
    for (const app of sharePointAppList) {
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

  console.log(`clean dev tunnel`);
  const devTunnelCleanHelper = await DevTunnelCleanHelper.create(
    Env.cleanTenantId,
    Env.username,
    Env.password
  );
  await devTunnelCleanHelper.deleteAll();

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
      Env.password
    );
    console.log(`clean M365 Titles (exclude ${excludePrefix})`);
    try {
      const acquisitions = await m365TitleCleanService.listAcquisitions();
      if (acquisitions) {
        for (const acquisition of acquisitions) {
          if (!acquisition.titleDefinition.name.startsWith(excludePrefix)) {
            console.log(acquisition.titleDefinition.name);
            console.log(acquisition.titleId);
            const result = await m365TitleCleanService.unacquire(
              acquisition.titleId,
              1
            );
            if (!retry && result) {
              retry = true;
            }
          }
        }
      }
    } catch (e: any) {
      console.log(`Get error: ${e.message}`);
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
  .then((_) => {
    console.log("Clean Job Done.");
  })
  .catch((error) => {
    console.error(error);
    process.exit(-1);
  });
