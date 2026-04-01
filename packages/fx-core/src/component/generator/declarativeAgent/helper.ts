// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  AppPackageFolderName,
  Context,
  DefaultApiSpecFolderName,
  DefaultPluginManifestFileName,
  err,
  FxError,
  GeneratorResult,
  Inputs,
  ok,
  PluginManifestSchema,
  Result,
  SystemError,
  UserError,
  Warning,
} from "@microsoft/teamsfx-api";
import { copilotGptManifestUtils } from "../../driver/teamsApp/utils/CopilotGptManifestUtils";
import { pluginManifestUtils } from "../../driver/teamsApp/utils/PluginManifestUtils";
import path from "path";
import fs from "fs-extra";
import { normalizePath } from "../../driver/teamsApp/utils/utils";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";
import { getEnvironmentVariables } from "../../utils/common";
import { sendTelemetryErrorEvent } from "../../../common/telemetry";
import { assembleError } from "../../../error";
import axios, { isAxiosError } from "axios";
import {
  GCScopes,
  getResourceServiceEndpoint,
  ResourceServiceType,
} from "../../../common/constants";
import {
  createGraphClientWithToken,
  encodeSharePointUrl,
  getDriveItemInfo,
  getSharePointSiteByRelativePath,
  ItemMetadata,
} from "./oneDriveSharePointHandler";
import { createContext } from "../../../common/globalVars";
import { QuestionNames } from "../../../question/questionNames";
import {
  fetchMCPTools,
  probeMCPServerAuth,
  readMCPToolsFromFile,
  resolveMCPOAuthMetadata,
} from "../../utils/mcpToolFetcher";
import { ActionInjector } from "../../configManager/actionInjector";
import { pathUtils } from "../../utils/pathUtils";

const logMessageKeys = {
  failValidateOneDriveSharePointItem:
    "core.createProjectQuestion.log.fail.validateOneDriveSharePointItem",
  invalidOneDriveSharePointURL: "core.createProjectQuestion.log.fail.invalidOneDriveSharePointURL",
};
export interface AddExistingPluginResult {
  warnings: Warning[];
  destinationPluginManifestPath: string;
}

const pluginManifestPlaceholderWarning = "add-exsiting-plugin-manifest-placehoder";
const apiSpecPlaceholderWarning = "add-exsiting-plugin-api-spec-placehoder";
const readApiSpecErrorTelemetry = "read-api-spec-error";

export async function addExistingPlugin(
  declarativeCopilotManifestPath: string,
  fromPluginManifestPath: string,
  fromApiSpecPath: string,
  actionId: string,
  context: Context,
  source: string
): Promise<Result<AddExistingPluginResult, FxError>> {
  const pluginManifestRes =
    await pluginManifestUtils.readPluginManifestFile(fromPluginManifestPath);
  if (pluginManifestRes.isErr()) {
    return err(pluginManifestRes.error);
  }
  const pluginManifest = pluginManifestRes.value;

  // prerequiste check
  const checkRes = validateSourcePluginManifest(pluginManifest, source);
  if (checkRes.isErr()) {
    return err(checkRes.error);
  }

  const runtimes = pluginManifest.runtimes!; // have validated that the value exists.
  const destinationApiSpecRelativePath = runtimes.find((runtime) => runtime.type === "OpenApi")!
    .spec.url as string; // have validated that the value exists.

  const outputFolder = path.dirname(declarativeCopilotManifestPath);

  // Copy OpenAPI spec
  const originalDestApiSPecRelativePath = path.resolve(
    outputFolder,
    destinationApiSpecRelativePath
  );
  let destinationApiSpecPath = originalDestApiSPecRelativePath;
  const needUpdatePluginManifest =
    (await fs.pathExists(originalDestApiSPecRelativePath)) ||
    path.relative(outputFolder, originalDestApiSPecRelativePath).startsWith("..");

  if (needUpdatePluginManifest) {
    destinationApiSpecPath = await pluginManifestUtils.getDefaultNextAvailableApiSpecPath(
      fromApiSpecPath,
      path.join(outputFolder, DefaultApiSpecFolderName)
    );
  }
  await fs.ensureFile(destinationApiSpecPath);
  await fs.copyFile(fromApiSpecPath, destinationApiSpecPath);

  // Save plugin manifest
  if (needUpdatePluginManifest) {
    const runtimeSpecUrl = normalizePath(path.relative(outputFolder, destinationApiSpecPath), true);
    for (const runtime of runtimes) {
      if (runtime.type === "OpenApi" && runtime.spec?.url) {
        runtime.spec.url = runtimeSpecUrl;
      }
    }
  }

  const destinationPluginManifestPath =
    await copilotGptManifestUtils.getDefaultNextAvailablePluginManifestPath(outputFolder);
  await fs.ensureFile(destinationPluginManifestPath);
  const pluginManifestContent = JSON.stringify(pluginManifest, undefined, 4);
  await fs.writeFile(destinationPluginManifestPath, pluginManifestContent);

  // Update declarative copilot plugin manifest
  const addActionRes = await copilotGptManifestUtils.addAction(
    declarativeCopilotManifestPath,
    actionId,
    normalizePath(path.relative(outputFolder, destinationPluginManifestPath), true)
  );
  if (addActionRes.isErr()) {
    return err(addActionRes.error);
  }

  const warnings: Warning[] = [];
  const pluginManifestVariables = getEnvironmentVariables(JSON.stringify(pluginManifest));
  if (pluginManifestVariables.length > 0) {
    warnings.push({
      type: pluginManifestPlaceholderWarning,
      content: getLocalizedString(
        "core.addPlugin.warning.manifestVariables",
        pluginManifestVariables.join(", ")
      ),
    });
  }

  try {
    const apiSpecContent = await fs.readFile(destinationApiSpecPath, "utf8");
    const apiSpecVariables = getEnvironmentVariables(apiSpecContent);
    if (apiSpecVariables.length > 0) {
      warnings.push({
        type: apiSpecPlaceholderWarning,
        content: getLocalizedString(
          "core.addPlugin.warning.apiSpecVariables",
          apiSpecVariables.join(", ")
        ),
      });
    }
  } catch (e) {
    sendTelemetryErrorEvent(source, readApiSpecErrorTelemetry, assembleError(e));
  }

  return ok({
    destinationPluginManifestPath,
    warnings,
  });
}

export function validateSourcePluginManifest(
  manifest: PluginManifestSchema,
  source: string
): Result<undefined, UserError> {
  if (!manifest.schema_version) {
    return err(
      new UserError(
        source,
        "MissingSchemaVersion",
        getDefaultString(
          "core.createProjectQuestion.addPlugin.MissingRequiredProperty",
          "schema_version"
        ),
        getLocalizedString(
          "core.createProjectQuestion.addPlugin.MissingRequiredProperty",
          "schema_version"
        )
      )
    );
  }

  if (!manifest.runtimes) {
    return err(
      new UserError(
        source,
        "MissingRuntimes",
        getDefaultString(
          "core.createProjectQuestion.addPlugin.MissingRequiredProperty",
          "runtimes"
        ),
        getLocalizedString(
          "core.createProjectQuestion.addPlugin.MissingRequiredProperty",
          "runtimes"
        )
      )
    );
  }

  const apiSpecPaths = new Set<string>();
  for (const runtime of manifest.runtimes) {
    if (runtime.type === "OpenApi" && runtime.spec?.url) {
      apiSpecPaths.add(runtime.spec.url);
    }
  }

  if (apiSpecPaths.size === 0) {
    return err(
      new UserError(
        source,
        "MissingApiSpec",
        getDefaultString(
          "core.createProjectQuestion.addPlugin.pluginManifestMissingApiSpec",
          "OpenApi"
        ),
        getLocalizedString(
          "core.createProjectQuestion.addPlugin.pluginManifestMissingApiSpec",
          "OpenApi"
        )
      )
    );
  }

  if (apiSpecPaths.size > 1) {
    return err(
      new UserError(
        source,
        "MultipleApiSpecInPluginManifest",
        getDefaultString(
          "core.createProjectQuestion.addPlugin.pluginManifestMultipleApiSpec",
          Array.from(apiSpecPaths).join(", ")
        ),
        getLocalizedString(
          "core.createProjectQuestion.addPlugin.pluginManifestMultipleApiSpec",
          Array.from(apiSpecPaths).join(", ")
        )
      )
    );
  }

  return ok(undefined);
}

export async function getODSPItemInfo(
  context: Context,
  itemUrl: string | undefined
): Promise<Result<ItemMetadata[], UserError>> {
  if (!itemUrl) {
    return err(
      new UserError("validateOneDriveSharePointItem", "InvalidInput", "Item URL is required")
    );
  }

  try {
    const graphClientResult = await createGraphClientWithToken(context);
    if (graphClientResult.isErr()) {
      return err(graphClientResult.error);
    }
    const graphClient = graphClientResult.value;

    const siteResult = await getSharePointSiteByRelativePath(graphClient, itemUrl);
    if (siteResult.isOk()) {
      const site = siteResult.value;
      return ok([
        {
          id: site.id,
          name: site.name,
          webId: site.webId,
          siteId: site.siteId,
        },
      ]);
    }

    const encodedUrl = encodeSharePointUrl(itemUrl);
    const driveItem = await getDriveItemInfo(graphClient, encodedUrl);

    return ok([
      {
        id: driveItem.id,
        name: driveItem.name,
        uniqueId: driveItem.uniqueId,
        listId: driveItem.listId,
        webId: driveItem.webId,
        siteId: driveItem.siteId,
        itemType: driveItem.itemType,
      },
    ]);
  } catch (error) {
    if (isAxiosError(error) && error.response) {
      if (error.response.status >= 400 && error.response.status < 510) {
        context.logProvider?.error(
          getLocalizedString(logMessageKeys.failValidateOneDriveSharePointItem, error.message)
        );
        return err(
          new UserError(
            "ValidateOneDriveSharePointURL",
            "GraphApiError",
            error.message,
            error.message
          )
        );
      }
    }

    const message = JSON.stringify(error);
    context.logProvider?.error(
      getLocalizedString(logMessageKeys.failValidateOneDriveSharePointItem, message)
    );
    return err(new SystemError("ValidateOneDriveSharePointURL", "GraphApiError", message, message));
  }
}

export interface GCItem {
  id: string;
  label: string;
}
export async function getGraphConnectors(): Promise<GCItem[]> {
  const context = createContext();
  const graphTokenRes = await context.tokenProvider!.m365TokenProvider.getAccessToken({
    scopes: GCScopes,
  });
  if (graphTokenRes.isErr()) {
    throw graphTokenRes.error;
  }
  const graphToken = graphTokenRes.value;

  const instance = axios.create({
    baseURL: `${getResourceServiceEndpoint(ResourceServiceType.Graph)}/v1.0`,
    headers: { Authorization: `Bearer ${graphToken}` },
  });

  try {
    const res = await instance.get(`/external/connections`);
    const data = res.data;
    const result = data.value.map((item: { id: string; name: string }) => {
      return { id: item.id, label: item.id };
    });
    return result;
  } catch (error) {
    if (error.response?.status === 403) {
      const err = new UserError(
        "getCopilotConnectors",
        "GraphApiError",
        getDefaultString("core.GCList.insufficientPermission"),
        getDefaultString("core.GCList.insufficientPermission")
      );
      throw err;
    } else {
      const message = `Failed to get Copilot connector item: ${
        error instanceof Error ? error.message : String(error)
      }`;
      const err = new UserError("copilotConnectors", "GraphApiError", message, message);
      throw err;
    }
  }
}

export async function generateForMCPForDA(
  destinationPath: string,
  inputs: Inputs
): Promise<Result<GeneratorResult, FxError>> {
  // 1. Get ai-plugin.json
  const aiPluginFilePath = path.join(
    destinationPath,
    AppPackageFolderName,
    DefaultPluginManifestFileName
  );
  if (!(await fs.pathExists(aiPluginFilePath))) {
    const error = new SystemError(
      "MCPForDAPluginManifestNotFound",
      "PluginManifestNotFound",
      getDefaultString("core.MCPForDA.pluginManifestNotFound", aiPluginFilePath),
      getLocalizedString("core.MCPForDA.pluginManifestNotFound", aiPluginFilePath)
    );
    return err(error);
  }

  const mcpServerUrl = inputs[QuestionNames.MCPForDAServerUrl];
  const serverName = inputs[QuestionNames.MCPForDAServerName];
  const warnings: Warning[] = [];

  // If a tools file is provided (CLI flow), load tools from it
  const toolsFilePath = inputs[QuestionNames.MCPToolsFilePath];
  const existingTools = inputs[QuestionNames.MCPForDAAvailableTools];
  if (toolsFilePath && (!existingTools || existingTools.length === 0)) {
    try {
      const fileTools = await readMCPToolsFromFile(toolsFilePath);
      inputs[QuestionNames.MCPForDAAvailableTools] = fileTools;
      if (!inputs[QuestionNames.MCPForDAPreFetchTools]) {
        inputs[QuestionNames.MCPForDAPreFetchTools] = fileTools.map((t: any) => t.name);
      }
    } catch {
      warnings.push({
        type: "mcpToolsFileReadError",
        content: getLocalizedString("core.MCPForDA.toolsFileReadError", toolsFilePath),
      });
    }
  }

  // Probe auth if tools were loaded from file but auth not yet detected
  if (
    inputs[QuestionNames.MCPForDAAvailableTools]?.length > 0 &&
    !inputs[QuestionNames.MCPForDAAuth] &&
    mcpServerUrl
  ) {
    try {
      const authProbe = await probeMCPServerAuth(mcpServerUrl);
      if (authProbe.requiresAuth) {
        inputs[QuestionNames.MCPForDAAuth] = "OAuthPluginVault";
        if (authProbe.authMetadataUrl) {
          inputs[QuestionNames.MCPForDAAuthMetadataUrl] = authProbe.authMetadataUrl;
        }
      }
    } catch {
      // Auth probe failed — continue without auth
    }
  }

  // Auto-fetch tools from MCP server if no tools loaded yet
  const currentTools = inputs[QuestionNames.MCPForDAAvailableTools];
  if ((!currentTools || currentTools.length === 0) && mcpServerUrl) {
    try {
      const result = await fetchMCPTools(mcpServerUrl);
      if (!result.requiresAuth && result.tools.length > 0) {
        inputs[QuestionNames.MCPForDAAvailableTools] = result.tools;
        if (!inputs[QuestionNames.MCPForDAPreFetchTools]) {
          inputs[QuestionNames.MCPForDAPreFetchTools] = result.tools.map((t) => t.name);
        }
      } else if (result.requiresAuth) {
        // Store auth metadata for later use
        inputs[QuestionNames.MCPForDAAuth] = "OAuthPluginVault";
        if (result.authMetadataUrl) {
          inputs[QuestionNames.MCPForDAAuthMetadataUrl] = result.authMetadataUrl;
        }
        warnings.push({
          type: "mcpAuthRequired",
          content: getLocalizedString("core.MCPForDA.authRequired", mcpServerUrl),
        });
      } else {
        warnings.push({
          type: "mcpNoToolsFetched",
          content: getLocalizedString("core.MCPForDA.noToolsFetched", mcpServerUrl),
        });
      }
    } catch {
      warnings.push({
        type: "mcpFetchError",
        content: getLocalizedString("core.MCPForDA.fetchError", mcpServerUrl),
      });
    }
  }

  // 2. Read ai-plugin.json
  const aiPluginContent = await fs.readJSON(aiPluginFilePath);

  const mcpToolsDetail = inputs[QuestionNames.MCPForDAAvailableTools];
  const mcpToolsSelected = inputs[QuestionNames.MCPForDAPreFetchTools];

  if (
    mcpToolsDetail &&
    mcpToolsDetail.length > 0 &&
    mcpToolsSelected &&
    mcpToolsSelected.length > 0
  ) {
    const mcpAuth = inputs[QuestionNames.MCPForDAAuth];

    // Process tools: if serverName exists (VS Code flow), strip the prefix; otherwise use as-is (CLI flow)
    let processedTools: any[];
    if (serverName) {
      processedTools = mcpToolsDetail
        .filter((tool: any) => tool.name.includes(serverName))
        .map((tool: any) => {
          const index = tool.name.indexOf(serverName);
          const newName = (tool.name as string).substring(
            (index as number) + (serverName.length as number) + 1
          );
          return {
            ...tool,
            name: newName,
          };
        })
        .filter((tool: any) => mcpToolsSelected.includes(tool.name));
    } else {
      processedTools = mcpToolsDetail.filter((tool: any) => mcpToolsSelected.includes(tool.name));
    }

    aiPluginContent.functions = processedTools.map((tool: any) => {
      return {
        name: tool.name,
        description: tool.description || "",
      };
    });

    // Write mcp-tools-1.json with full tool definitions (no field filtering)
    const mcpToolsFileName = "mcp-tools-1.json";
    const mcpToolsOutputPath = path.join(destinationPath, AppPackageFolderName, mcpToolsFileName);
    await fs.writeJSON(mcpToolsOutputPath, { tools: processedTools }, { spaces: 4 });

    const runtimeSpec: any = {
      url: mcpServerUrl,
      mcp_tool_description: {
        file: mcpToolsFileName,
      },
    };
    aiPluginContent.runtimes = [
      {
        type: "RemoteMCPServer",
        spec: runtimeSpec,
        run_for_functions: aiPluginContent.functions.map((func: any) => func.name),
      },
    ];

    // Auth handling
    if (mcpAuth === "OAuthPluginVault") {
      const authType = inputs[QuestionNames.MCPForDAAuthType];
      if (!authType) {
        return err(
          new UserError(
            "MCPForDA",
            "MissingMCPAuthType",
            getDefaultString("core.MCPForDA.missingAuthType"),
            getLocalizedString("core.MCPForDA.missingAuthType")
          )
        );
      }

      const registrationId = "MCP_DA_AUTH_ID_ACTION_1";
      aiPluginContent.runtimes[0].auth = {
        type: "OAuthPluginVault",
        reference_id: `$\{\{${registrationId}\}\}`,
      };

      // Resolve OAuth metadata and inject oauth/register into m365agents.yml
      try {
        let authorizationUrl: string | undefined;
        let tokenUrl: string | undefined;
        let refreshUrl: string | undefined;

        if (authType === "oauth") {
          const metadata = await resolveMCPOAuthMetadata(
            inputs[QuestionNames.MCPForDAAuthMetadataUrl],
            inputs[QuestionNames.MCPForDAAuthWellKnownUrl]
          );
          authorizationUrl = metadata.authorizationUrl;
          tokenUrl = metadata.tokenUrl;
          refreshUrl = metadata.refreshUrl;
        }

        const ymlPath = pathUtils.getYmlFilePath(destinationPath);
        if (ymlPath) {
          await ActionInjector.injectCreateOAuthActionForMCP(
            ymlPath,
            authType,
            "action_1",
            registrationId,
            mcpServerUrl,
            authorizationUrl,
            tokenUrl,
            refreshUrl
          );
        }
      } catch (error: any) {
        warnings.push({
          type: "mcpAuthMetadataError",
          content: getLocalizedString("core.MCPForDA.mcpAuthMetadataMissingError", error.message),
        });
      }
    }
  }
  // If no tools available, leave ai-plugin.json with empty functions/runtimes (template default)

  // 3. Write ai-plugin.json
  await fs.writeJSON(aiPluginFilePath, aiPluginContent, { spaces: 4 });

  return ok({ warnings });
}
