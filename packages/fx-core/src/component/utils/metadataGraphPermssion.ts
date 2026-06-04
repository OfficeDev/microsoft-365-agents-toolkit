// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import fs from "fs-extra";
import path from "path";
import isUUID from "validator/lib/isUUID";
import { TelemetryProperty } from "../../common/telemetry";
import { MetadataV3 } from "../../common/versionMetadata";
import { ProjectModel } from "../configManager/interface";
import { AADManifest } from "../driver/aad/interface/AADManifest";
import { getDetailedGraphPermissionMap, graphAppId, graphAppName } from "../driver/aad/permissions";
import { actionName } from "../driver/aad/update";

export const metadataGraphPermissionDeps = {
  getDetailedGraphPermissionMap,
};

interface summary {
  hasGraphPermission: boolean;
  hasRole: boolean;
  hasAdminScope: boolean;
  scopes: string[];
  roles: string[];
}
class MetadataGraphPermissionUtil {
  async parseAadManifest(
    ymlPath: string,
    model: ProjectModel,
    props: { [key: string]: string }
  ): Promise<void> {
    let aadManifestName = MetadataV3.aadManifestFileName;
    const updateAction = model.provision?.driverDefs.find((def) => def.uses === actionName);
    // if aadApp/update action is defined, use the manifest file in the action
    if (updateAction) {
      const parameters = updateAction.with as { [key: string]: string };
      if (parameters && parameters["manifestPath"]) {
        aadManifestName = parameters["manifestPath"];
      }
    }
    const aadManifestPath = path.join(path.dirname(ymlPath), aadManifestName);
    if (!(await fs.pathExists(aadManifestPath))) {
      props[TelemetryProperty.AadManifest] = "false";
      return;
    }
    props[TelemetryProperty.AadManifest] = "true";

    try {
      const manifestString = await fs.readFile(aadManifestPath, "utf8");
      const manifest = JSON.parse(manifestString);
      const graphPermissionSummary = this.summary(manifest);
      if (graphPermissionSummary) {
        props[TelemetryProperty.GraphPermission] = graphPermissionSummary.hasGraphPermission
          ? "true"
          : "false";
        props[TelemetryProperty.GraphPermissionHasRole] = graphPermissionSummary.hasRole
          ? "true"
          : "false";
        props[TelemetryProperty.GraphPermissionHasAdminScope] = graphPermissionSummary.hasAdminScope
          ? "true"
          : "false";
        props[TelemetryProperty.GraphPermissionScopes] = graphPermissionSummary.scopes.join(",");
        props[TelemetryProperty.GraphPermissionRoles] = graphPermissionSummary.roles.join(",");
        model.aadPermission = {
          graphPermission: {
            hasGraphPermission: graphPermissionSummary.hasGraphPermission,
            hasRole: graphPermissionSummary.hasRole,
            hasAdminScope: graphPermissionSummary.hasAdminScope,
            scopes: graphPermissionSummary.scopes,
            roles: graphPermissionSummary.roles,
          },
        };
      }
    } catch (error) {
      return;
    }
  }

  summary(manifest: AADManifest): summary | undefined {
    let hasGraphPermission = false;
    let hasRole = false;
    let hasAdminScope = false;
    const scopes: string[] = [];
    const roles: string[] = [];
    const graphPermissionMap = metadataGraphPermissionDeps.getDetailedGraphPermissionMap();
    if (!graphPermissionMap) {
      return undefined;
    }
    const graphPermission = manifest.requiredResourceAccess?.find(
      (item) => item.resourceAppId === graphAppId || item.resourceAppId === graphAppName
    );
    if (!graphPermission) {
      return {
        hasGraphPermission,
        hasRole,
        hasAdminScope,
        scopes,
        roles,
      };
    }
    hasGraphPermission = true;
    graphPermission.resourceAccess?.forEach((access) => {
      if (access.type === "Role") {
        hasRole = true;
        const id = isUUID(access.id) ? access.id : graphPermissionMap.roles[access.id];
        if (graphPermissionMap.roleIds[id]) {
          roles.push(graphPermissionMap.roleIds[id].value);
        }
      } else {
        const id = isUUID(access.id) ? access.id : graphPermissionMap.scopes[access.id];
        if (graphPermissionMap.scopeIds[id]) {
          scopes.push(graphPermissionMap.scopeIds[id].value);
          if (graphPermissionMap.scopeIds[id].type === "Admin") {
            hasAdminScope = true;
          }
        }
      }
    });
    return {
      hasGraphPermission,
      hasRole,
      hasAdminScope,
      scopes,
      roles,
    };
  }
}

export const metadataGraphPermissionUtil = new MetadataGraphPermissionUtil();
