// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ConsumerOperation,
  GeneratePluginResult,
  KiotaSearchResultItem,
  searchDescription,
  setKiotaConfig,
  getKiotaTree,
  generatePlugin,
  KiotaTreeResult,
  PluginAuthType,
} from "@microsoft/kiota";
import { KiotaGeneratePluginError } from "../error";
import { getLocalizedString } from "./localizeUtils";
import { Utils } from "@microsoft/m365-spec-parser";

const ERROR_LOG_LEVEL = 4;

export async function searchOpenAPISpec(query: string): Promise<SearchOpenAPISpecResult[]> {
  if (process.env.KIOTA_BINARY_PATH) {
    setKiotaConfig({ binaryLocation: process.env.KIOTA_BINARY_PATH });
  }

  const searchResult: Record<string, KiotaSearchResultItem> | undefined = await searchDescription({
    searchTerm: query,
    clearCache: false,
  });

  const result: SearchOpenAPISpecResult[] = [];

  if (searchResult) {
    for (const key in searchResult) {
      const api = searchResult[key];
      if (api && api.DescriptionUrl) {
        result.push({
          key: key,
          url: api.DescriptionUrl,
          description: api.Description,
        });
      }
    }
  }

  return result;
}

export async function listAPITreeInfo(
  specPath: string,
  includeFilters?: string[],
  excludeFilters?: string[]
): Promise<KiotaTreeResult> {
  if (process.env.KIOTA_BINARY_PATH) {
    setKiotaConfig({ binaryLocation: process.env.KIOTA_BINARY_PATH });
  }
  const treeInfo = await getKiotaTree({
    includeFilters: includeFilters,
    descriptionPath: specPath,
    excludeFilters: excludeFilters,
    clearCache: true,
    includeKiotaValidationRules: true,
  });

  if (!treeInfo) {
    throw new Error(getLocalizedString("error.kiotaClient.EmptyResult"));
  }

  const errors = treeInfo.logs
    .filter((log) => log.level >= ERROR_LOG_LEVEL)
    .map((log) => log.message);

  if (errors.length > 0) {
    throw new Error(errors.join("\n"));
  }

  const treeInfoStr = JSON.stringify(treeInfo);
  // kiota treeInfo may contains url encoded characters: $%7B%7BAAD_APP_TENANT_ID%7D%7D,
  const decodedTreeInfoStr = decodeURIComponent(treeInfoStr);
  const resolvedTreeInfo = Utils.resolveEnv(decodedTreeInfoStr);
  return JSON.parse(resolvedTreeInfo) as KiotaTreeResult;
}

export async function kiotageneratePlugin(
  specPath: string,
  outputPath: string,
  pluginName: string,
  workingDirectory: string,
  authType?: PluginAuthType,
  authRefId?: string,
  includePatterns?: string[],
  excludePatterns?: string[],
  noWorkspace?: boolean
): Promise<GeneratePluginResult> {
  if (process.env.KIOTA_BINARY_PATH) {
    setKiotaConfig({ binaryLocation: process.env.KIOTA_BINARY_PATH });
  }

  const config = {
    descriptionPath: specPath,
    outputPath: outputPath,
    includePatterns: includePatterns ?? [],
    excludePatterns: excludePatterns ?? [],
    pluginName: pluginName,
    clearCache: false,
    cleanOutput: false,
    disabledValidationRules: [],
    operation: ConsumerOperation.Edit,
    pluginAuthType: authType ?? null,
    pluginAuthRefid: authRefId ?? undefined,
    workingDirectory: workingDirectory,
    noWorkspace: noWorkspace,
  };

  try {
    const result: GeneratePluginResult | undefined = await generatePlugin(config);
    if (!result) {
      throw new Error("Get empty result from kiota");
    }

    if (!result.isSuccess) {
      const errorMessage = result.logs
        .filter((log) => log.level >= ERROR_LOG_LEVEL)
        .map((log) => log.message)
        .join(";");
      throw new Error(errorMessage);
    }

    return result;
  } catch (error) {
    throw new KiotaGeneratePluginError(error.message);
  }
}

export interface SearchOpenAPISpecResult {
  key: string;
  url: string;
  description: string;
}
