// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  KiotaSearchResultItem,
  searchDescription,
  setKiotaConfig,
  getKiotaTree,
  KiotaTreeResult,
} from "@microsoft/kiota";

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

export interface SearchOpenAPISpecResult {
  key: string;
  url: string;
  description: string;
}
