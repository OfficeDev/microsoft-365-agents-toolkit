// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import path from "path";
import fs from "fs-extra";
import { MetadataV3 } from "../../common/versionMetadata";
import { ProjectModel } from "../configManager/interface";
import { ProjectTypeProps } from "../../common/telemetry";
import { manifestUtils } from "../driver/teamsApp/utils/ManifestUtils";
import { DeclarativeAgentManifestWrapper, PluginManifestWrapper } from "@microsoft/teamsfx-api";

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
        const declarativeAgentWrapper = await DeclarativeAgentManifestWrapper.read(
          declarativeAgentJsonPath
        );

        // Use wrapper's typed getters for capabilities and actions
        const capabilities = declarativeAgentWrapper.capabilities;
        const actions = declarativeAgentWrapper.actions;

        const capabilitiesCount = capabilities.length;
        props[ProjectTypeProps.DeclarativeAgentCapabilitiesCount] = capabilitiesCount.toString();

        if (capabilitiesCount > 0) {
          props[ProjectTypeProps.DeclarativeAgentCapabilities] = capabilities
            .map((capability) => capability.name)
            .join(",");
        } else {
          props[ProjectTypeProps.DeclarativeAgentCapabilities] = "";
        }

        const actionsCount = actions.length;
        props[ProjectTypeProps.DeclarativeAgentActionsCount] = actionsCount.toString();

        if (actionsCount > 0) {
          const pluginPaths = actions.map((action) => action.file);

          const declarativeAgentFolder = path.dirname(declarativeAgentJsonPath);
          const authInfo = [];
          for (const pluginPath of pluginPaths) {
            const pluginJsonPath = path.join(declarativeAgentFolder, pluginPath);
            if (await fs.pathExists(pluginJsonPath)) {
              const pluginWrapper = await PluginManifestWrapper.read(pluginJsonPath);
              const runtimes = pluginWrapper.runtimes;
              if (runtimes.length > 0) {
                const authStr = runtimes
                  .filter((runtime) => runtime.type === "OpenApi")
                  .map((runtime) => runtime.auth?.type)
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
