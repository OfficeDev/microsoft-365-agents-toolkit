// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import path from "path";
import fs from "fs-extra";
import { MetadataV3 } from "../../common/versionMetadata";
import { ProjectModel } from "../configManager/interface";
import { ProjectTypeProps } from "../../common/telemetry";
import { manifestUtils } from "../driver/teamsApp/utils/ManifestUtils";
import { DeclarativeCopilotManifestSchema, PluginManifestSchema } from "@microsoft/teamsfx-api";

class MetadataDAPropertiesUtil {
  async parseManifest(
    ymlPath: string,
    model: ProjectModel,
    props: { [key: string]: string }
  ): Promise<void> {
    let manifestName = path.join(MetadataV3.teamsManifestFolder, MetadataV3.teamsManifestFileName);
    const action = model.provision?.driverDefs.find((def) => def.uses === "teamsApp/zipAppPackage");
    if (action) {
      const parameters = action.with as { [key: string]: string };
      if (parameters && parameters["manifestPath"]) {
        manifestName = parameters["manifestPath"];
      }
    }

    const projectRoot = path.dirname(ymlPath);
    const manifestPath = path.join(projectRoot, manifestName);
    try {
      const result = await manifestUtils._readAppManifest(manifestPath);
      if (result.isErr()) {
        return;
      }
      const manifest = result.value;
      const declarativeAgentRelativePath = manifest.copilotAgents?.declarativeAgents?.[0].file;

      if (declarativeAgentRelativePath) {
        const manifestFolder = path.dirname(manifestPath);
        const declarativeAgentJsonPath = path.join(manifestFolder, declarativeAgentRelativePath);
        const declarativeAgentJson = (await fs.readJSON(
          declarativeAgentJsonPath
        )) as DeclarativeCopilotManifestSchema;

        const capabilitiesCount = declarativeAgentJson.capabilities?.length ?? 0;
        props[ProjectTypeProps.DeclarativeAgentCapabilitiesCount] = capabilitiesCount.toString();

        if (declarativeAgentJson.capabilities) {
          props[ProjectTypeProps.DeclarativeAgentCapabilities] = declarativeAgentJson.capabilities
            .map((capability: any) => capability.name)
            .join(",");
        } else {
          props[ProjectTypeProps.DeclarativeAgentCapabilities] = "";
        }

        const actionsCount = declarativeAgentJson.actions?.length ?? 0;
        props[ProjectTypeProps.DeclarativeAgentActionsCount] = actionsCount.toString();

        if (declarativeAgentJson.actions) {
          const pluginPaths = declarativeAgentJson.actions.map((action: any) => action.file);

          const declarativeAgentFolder = path.dirname(declarativeAgentJsonPath);
          const authInfo = [];
          for (const pluginPath of pluginPaths ?? []) {
            const pluginJsonPath = path.join(declarativeAgentFolder, pluginPath);
            if (await fs.pathExists(pluginJsonPath)) {
              const pluginJson = (await fs.readJSON(pluginJsonPath)) as PluginManifestSchema;
              const runtimes = pluginJson.runtimes;
              if (runtimes) {
                const authStr = runtimes
                  .filter((runtime) => runtime.type === "OpenApi")
                  .map((runtime: any) => runtime.auth?.type)
                  .filter((type) => type !== undefined)
                  .join(",");
                authInfo.push(authStr);
              }
            }
          }
          props[ProjectTypeProps.DeclarativeAgentPluginAuthTypes] = authInfo.join(";");
        } else {
          props[ProjectTypeProps.DeclarativeAgentPluginAuthTypes] = "";
        }
      }
    } catch (error) {
      return;
    }
  }
}

export const metadataDAPropertiesUtil = new MetadataDAPropertiesUtil();
