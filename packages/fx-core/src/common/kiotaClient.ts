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
): Promise<KiotaTreeResult | undefined> {
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

  return treeInfo;
}

export async function kiotageneratePlugin(
  specPath: string,
  outputPath: string,
  pluginName: string,
  workingDirectory: string,
  authType?: PluginAuthType,
  authRefId?: string,
  includePatterns?: string[],
  excludePatterns?: string[]
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
  };

  try {
    const result: GeneratePluginResult | undefined = await generatePlugin(config);
    if (!result) {
      throw new Error("Get empty result from kiota");
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
