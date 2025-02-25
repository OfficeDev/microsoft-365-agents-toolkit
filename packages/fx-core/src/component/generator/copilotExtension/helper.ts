// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  Context,
  DefaultApiSpecFolderName,
  err,
  FxError,
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
import axios from "axios";
import { createContext } from "vm";
import { GCScopes } from "../../../common/constants";
import { GetGraphTokenFailedError } from "../../driver/deploy/spfx/error/getGraphTokenFailedError";
import { createGraphClientWithToken, encodeSharePointUrl, getDriveItemInfo, ItemMetadata } from "./onedriveSharePointHandler";

const logMessageKeys = {
  failValidateOneDriveSharePointItem: "core.createProjectQuestion.log.fail.validateOneDriveSharePointItem",
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
  const pluginManifestRes = await pluginManifestUtils.readPluginManifestFile(
    fromPluginManifestPath
  );
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
  itemUrl: string | undefined,
  inputs: Inputs,
  shouldLogWarning: boolean,
  existingCorrelationId?: string
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

    const encodedUrl = encodeSharePointUrl(itemUrl);
    const driveItem = await getDriveItemInfo(graphClient, encodedUrl);

    return ok([
      {
        name: driveItem.name,
        uniqueId: driveItem.uniqueId,
        listId: driveItem.listId,
        webId: driveItem.webId,
        siteId: driveItem.siteId,
        itemType: driveItem.itemType,
      },
    ]);
  } catch (error) {
    if (axios.isAxiosError(error) && error.response) {
      if (error.response!.status >= 400 && error.response!.status < 510) {
        context.logProvider?.error(
          getLocalizedString(logMessageKeys.failValidateOneDriveSharePointItem, error.message)
        );
        return err(new UserError(
          "ValidateOneDriveSharePointURL",
          "GraphApiError",
          error.message,
          error.message
        ));
      }
    }

    const message = JSON.stringify(error);
    context.logProvider?.error(
      getLocalizedString(logMessageKeys.failValidateOneDriveSharePointItem, message)
    );
    return err(new SystemError(
      "ValidateOneDriveSharePointURL",
      "GraphApiError",
      message,
      message
    ));
  }
}

export interface GCItem {
  id: string;
  label: string;
}
export async function getGraphConnectors(): Promise<GCItem[]> {
  const context = createContext();
  const graphTokenRes = await context.tokenProvider?.m365TokenProvider.getAccessToken({
    scopes: GCScopes,
  });
  if (!graphTokenRes?.isOk()) {
    throw err(new GetGraphTokenFailedError());
  }
  const graphToken = graphTokenRes.value;

  const instance = axios.create({
    baseURL: "https://graph.microsoft.com/v1.0",
    headers: { Authorization: `Bearer ${graphToken}` },
  });

  try {
    const res = await instance.get(`/external/connections?$select=id,name`);
    const data = res.data;
    return data.value.map((item: any) => {
      return { id: item.id, label: item.name };
    });
  } catch (error) {
    throw err(
      new UserError(
        "getGraphConnectors",
        "GraphApiError",
        `Failed to get Graph Connector item: ${
          error instanceof Error ? error.message : String(error)
        }`,
        error.response?.data?.message || error.message
      )
    );
  }
} 
