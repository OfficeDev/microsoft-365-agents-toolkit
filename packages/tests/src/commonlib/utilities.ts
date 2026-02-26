// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import * as uuid from "uuid";
import axios from "axios";
import {
  PluginId,
  ProjectSettingKey,
  provisionParametersKey,
  StateConfigKey,
  TestFilePath,
} from "./constants";
import * as fs from "fs-extra";
import * as path from "path";
import { promisify } from "util";
import { exec } from "child_process";

export async function execAsyncWithRetry(
  command: string,
  options: {
    cwd?: string;
    env?: NodeJS.ProcessEnv;
    timeout?: number;
  },
  retries = 3,
  newCommand?: string
): Promise<{
  stdout: string;
  stderr: string;
}> {
  const sleep = (ms: number) =>
    new Promise((resolve) => setTimeout(resolve, ms));
  while (retries > 0) {
    retries--;
    try {
      const result = await execAsync(command, options);
      return result;
    } catch (e: any) {
      console.log(
        `Run \`${command}\` failed with error msg: ${JSON.stringify(e)}.`
      );
      if (e.killed && e.signal == "SIGTERM") {
        console.log(`Command ${command} killed due to timeout`);
      }
      if (newCommand) {
        command = newCommand;
      }
      await sleep(10000);
    }
  }
  return execAsync(command, options);
}

export function editDotEnvFile(
  filePath: string,
  key: string,
  value: string
): void {
  try {
    const envFileContent: string = fs.readFileSync(filePath, "utf-8");
    const envVars: { [key: string]: string } = envFileContent
      .split("\n")
      .reduce((acc: { [key: string]: string }, line: string) => {
        const [key, value] = line.split("=");
        if (key && value) {
          acc[key.trim()] = value.trim();
        }
        return acc;
      }, {});
    envVars[key] = value;
    const newEnvFileContent: string = Object.entries(envVars)
      .map(([key, value]) => `${key}=${value}`)
      .join("\n");
    fs.writeFileSync(filePath, newEnvFileContent);
  } catch (error) {
    console.log('Failed to edit ".env" file.');
  }
}

const failedToParseResourceIdErrorMessage = (
  name: string,
  resourceId: string
) => `Failed to parse ${name} from resource id ${resourceId}`;

export function getKeyVaultSecretReference(
  vaultName: string,
  secretName: string
): string {
  return `@Microsoft.KeyVault(VaultName=${vaultName};SecretName=${secretName})`;
}

export async function getProvisionParameterValueByKey(
  projectPath: string,
  envName: string,
  key: string
): Promise<string | undefined> {
  const parameters = await fs.readJSON(
    path.join(
      projectPath,
      TestFilePath.configFolder,
      `azure.parameters.${envName}.json`
    )
  );
  if (
    parameters.parameters &&
    parameters.parameters.provisionParameters &&
    parameters.parameters.provisionParameters.value &&
    parameters.parameters.provisionParameters.value[key]
  ) {
    return parameters.parameters.provisionParameters.value[key];
  }
  return undefined;
}

export function getResourceGroupNameFromResourceId(resourceId: string): string {
  const result = parseFromResourceId(
    /\/resourceGroups\/([^\/]*)\//i,
    resourceId
  );
  if (!result) {
    throw new Error(
      failedToParseResourceIdErrorMessage("resource group name", resourceId)
    );
  }
  return result;
}

export function getSubscriptionIdFromResourceId(resourceId: string): string {
  const result = parseFromResourceId(
    /\/subscriptions\/([^\/]*)\//i,
    resourceId
  );
  if (!result) {
    throw new Error(
      failedToParseResourceIdErrorMessage("subscription id", resourceId)
    );
  }
  return result;
}

export function getSiteNameFromResourceId(webAppResourceId: string): string {
  const result = parseFromResourceId(
    /providers\/Microsoft.Web\/sites\/([^\/]*)/i,
    webAppResourceId
  );
  if (!result) {
    throw new Error(
      failedToParseResourceIdErrorMessage("site name", webAppResourceId)
    );
  }
  return result;
}

export function getKeyVaultNameFromResourceId(
  keyVaultResourceId: string
): string {
  const result = parseFromResourceId(
    /providers\/Microsoft.KeyVault\/vaults\/([^\/]*)/i,
    keyVaultResourceId
  );
  if (!result) {
    throw new Error(
      failedToParseResourceIdErrorMessage("key vault name", keyVaultResourceId)
    );
  }
  return result;
}

export function parseFromResourceId(
  pattern: RegExp,
  resourceId: string
): string {
  const result = resourceId.match(pattern);
  return result ? result[1].trim() : "";
}

export async function getWebappSettings(
  subscriptionId: string,
  rg: string,
  name: string,
  token: string
) {
  const baseUrlAppSettings = (
    subscriptionId: string,
    rg: string,
    name: string
  ) =>
    `https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${rg}/providers/Microsoft.Web/sites/${name}/config/appsettings/list?api-version=2019-08-01`;

  try {
    axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    const getResponse = await axios.post(
      baseUrlAppSettings(subscriptionId, rg, name)
    );
    if (getResponse && getResponse.data && getResponse.data.properties) {
      return getResponse.data.properties;
    }
  } catch (error) {
    console.log(error);
  }

  return undefined;
}

export async function getWebappConfigs(
  subscriptionId: string,
  rg: string,
  name: string,
  token: string
) {
  const baseUrlAppConfigs = (
    subscriptionId: string,
    rg: string,
    name: string
  ) =>
    `https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${rg}/providers/Microsoft.Web/sites/${name}/config/web?api-version=2021-02-01`;

  try {
    axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    const getResponse = await axios.get(
      baseUrlAppConfigs(subscriptionId, rg, name)
    );
    if (getResponse && getResponse.data && getResponse.data.properties) {
      return getResponse.data.properties;
    }
  } catch (error) {
    console.log(error);
  }

  return undefined;
}

export async function getBotServiceProperties(
  subscriptionId: string,
  rg: string,
  token: string,
  botClientId: string
): Promise<any> {
  const url = `https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${rg}/providers/Microsoft.BotService/botServices?api-version=2022-09-15`;
  try {
    axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    const response = await runWithRetry(() => axios.get(url));
    if (response?.data?.value) {
      const bot = (response.data.value as any[]).find(
        (b: any) => b.properties?.msaAppId === botClientId
      );
      return bot?.properties;
    }
  } catch (error) {
    console.log(error);
  }
  return undefined;
}

export async function getWebappServicePlan(
  subscriptionId: string,
  rg: string,
  name: string,
  token: string
) {
  const baseUrlPlan = (subscriptionId: string, rg: string, name: string) =>
    `https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${rg}/providers/Microsoft.Web/serverfarms/${name}?api-version=2019-08-01`;

  try {
    axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    const planResponse = await runWithRetry(() =>
      axios.get(baseUrlPlan(subscriptionId, rg, name))
    );
    if (
      planResponse &&
      planResponse.data &&
      planResponse.data.sku &&
      planResponse.data.sku.name
    ) {
      return planResponse.data.sku.name;
    }
  } catch (error) {
    console.log(error);
  }

  return undefined;
}

export async function runWithRetry<T>(fn: () => Promise<T>) {
  const maxTryCount = 3;
  const defaultRetryAfterInSecond = 2;
  const maxRetryAfterInSecond = 3 * 60;
  const secondInMilliseconds = 1000;

  for (let i = 0; i < maxTryCount - 1; i++) {
    try {
      const ret = await fn();
      return ret;
    } catch (e) {
      let retryAfterInSecond = defaultRetryAfterInSecond;
      if (e.response?.status === 429) {
        // See https://docs.microsoft.com/en-us/azure/azure-resource-manager/management/request-limits-and-throttling#error-code.
        const suggestedRetryAfter = e.response?.headers?.["retry-after"];
        // Explicit check, _retryAfter can be 0.
        if (suggestedRetryAfter !== undefined) {
          if (suggestedRetryAfter > maxRetryAfterInSecond) {
            // Don't wait too long.
            throw e;
          } else {
            // Take one more second for time error.
            retryAfterInSecond = suggestedRetryAfter + 1;
          }
        }
      }
      await new Promise((resolve) =>
        setTimeout(resolve, retryAfterInSecond * secondInMilliseconds)
      );
    }
  }

  return fn();
}

export function getUuid(): string {
  return uuid.v4();
}

export function getExpectedM365ApplicationIdUri(
  ctx: any,
  activeResourcePlugins: string[]
): string {
  let expectedM365ApplicationIdUri = "";
  if (activeResourcePlugins.includes(PluginId.FrontendHosting)) {
    const tabDomain = ctx[PluginId.FrontendHosting][StateConfigKey.domain];
    const m365ClientId = ctx[PluginId.Aad][StateConfigKey.clientId];
    expectedM365ApplicationIdUri =
      `api://${tabDomain}/` +
      (activeResourcePlugins.includes(PluginId.Bot)
        ? `botid-${ctx[PluginId.Bot][StateConfigKey.botId]}`
        : `${m365ClientId}`);
  } else if (activeResourcePlugins.includes(PluginId.Bot)) {
    expectedM365ApplicationIdUri = `api://botid-${
      ctx[PluginId.Bot][StateConfigKey.botId]
    }`;
  }
  return expectedM365ApplicationIdUri;
}

export async function getContainerAppProperties(
  subscriptionId: string,
  rg: string,
  containerAppName: string,
  token: string
) {
  const baseUrlAppSettings = (
    subscriptionId: string,
    rg: string,
    containerAppName: string
  ) =>
    `https://management.azure.com/subscriptions/${subscriptionId}/resourceGroups/${rg}/providers/Microsoft.App/containerApps/${containerAppName}?api-version=2023-05-01`;

  try {
    axios.defaults.headers.common["Authorization"] = `Bearer ${token}`;
    const planResponse = await runWithRetry(() =>
      axios.get(baseUrlAppSettings(subscriptionId, rg, containerAppName))
    );
    if (planResponse && planResponse.data && planResponse.data.properties) {
      return planResponse.data.properties;
    }
  } catch (error) {
    console.log(error);
  }

  return undefined;
}

export async function getActivePluginsFromProjectSetting(
  projectPath: string
): Promise<any> {
  const projectSettings = await fs.readJSON(
    path.join(
      projectPath,
      TestFilePath.configFolder,
      TestFilePath.projectSettingsFileName
    )
  );
  return projectSettings[ProjectSettingKey.solutionSettings][
    ProjectSettingKey.activeResourcePlugins
  ];
}

export const execAsync = promisify(exec);
