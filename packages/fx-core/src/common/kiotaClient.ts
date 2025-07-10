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
import path from "path";
import * as os from "os";

const ERROR_LOG_LEVEL = 4;

function setKiotaBinaryPath() {
  if (process.env.KIOTA_BINARY_PATH) {
    setKiotaConfig({ binaryLocation: process.env.KIOTA_BINARY_PATH });
  } else {
    // If running inside pkg package used by VS, set the binary location to a specific directory to avoid issues.
    const isInsidePkg = typeof (process as any).pkg !== "undefined";

    if (isInsidePkg) {
      const kiotaBinDir = path.join(os.homedir(), "kiota-bin");
      setKiotaConfig({ binaryLocation: kiotaBinDir });
    }
  }
}

export async function searchOpenAPISpec(query: string): Promise<SearchOpenAPISpecResult[]> {
  setKiotaBinaryPath();

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
  setKiotaBinaryPath();
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
  const resolvedTreeInfo = Utils.resolveEnv(treeInfoStr);
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
  setKiotaBinaryPath();

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
