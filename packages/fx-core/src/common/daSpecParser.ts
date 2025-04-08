// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ListAPIInfo,
  ListAPIResult,
  ParseOptions,
  ProjectType,
  SpecParser,
  AuthInfo,
} from "@microsoft/m365-spec-parser";
import { Platform } from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "./featureFlags";
import { listAPITreeInfo } from "./kiotaClient";
import {
  KiotaOpenApiNode,
  SecurityRequirementObject,
  SecuritySchemeObject,
} from "@microsoft/kiota";

const daProjectConfig: ParseOptions = {
  projectType: ProjectType.Copilot,
  isGptPlugin: true,
  allowMultipleParameters: true,
  allowMissingId: true,
  allowSwagger: true,
  allowAPIKeyAuth: true,
  allowBearerTokenAuth: true,
  allowOauth2: true,
  allowMethods: ["get", "post", "put", "delete", "patch", "head", "connect", "options", "trace"],
  allowResponseSemantics: true,
};

export async function listAPIInfo(specPath: string, platform?: string): Promise<ListAPIResult> {
  if (featureFlagManager.getBooleanValue(FeatureFlags.KiotaNPMIntegration)) {
    const treeInfo = await listAPITreeInfo(specPath);

    if (treeInfo && treeInfo.rootNode) {
      const operations: ListAPIInfo[] = extractOperations(
        treeInfo.rootNode,
        treeInfo.servers,
        treeInfo.security,
        treeInfo.securitySchemes
      );

      return {
        allAPICount: operations.length,
        validAPICount: operations.filter((api) => api.isValid).length,
        APIs: operations,
      };
    }

    return {
      allAPICount: 0,
      validAPICount: 0,
      APIs: [],
    };
  }

  const options: ParseOptions = {
    ...daProjectConfig,
    allowAPIKeyAuth: platform !== Platform.VS,
    allowBearerTokenAuth: platform !== Platform.VS,
    allowOauth2: platform !== Platform.VS,
  };

  const parser = new SpecParser(specPath, options);

  return await parser.list();
}

function extractOperations(
  node: KiotaOpenApiNode,
  parentServer: string[] = [],
  parentSecurity: SecurityRequirementObject[] = [],
  securitySchemes: {
    [key: string]: SecuritySchemeObject;
  } = {}
): ListAPIInfo[] {
  const operations: ListAPIInfo[] = [];

  const server = node.servers && node.servers.length > 0 ? node.servers : parentServer;
  const security = Object.keys(node.security || {}).length > 0 ? node.security : parentSecurity;

  if (node.isOperation) {
    const resourcePath = node.path.split("#")[0].replace(/\\/g, "/");

    let auth: AuthInfo | undefined;
    if (security) {
      const firstRequirementObject = security[0];
      if (firstRequirementObject) {
        const securitySchemeNames = Object.keys(firstRequirementObject);
        if (securitySchemeNames.length > 0) {
          const schemeName = securitySchemeNames[0];

          if (securitySchemeNames.length > 1) {
            auth = {
              name: securitySchemeNames.join(", "),
              authScheme: {
                type: "multipleAuth",
              },
            };
          } else {
            auth = {
              name: schemeName,
              authScheme: securitySchemes[schemeName],
            };
          }
        }
      }
    }

    const apiInfo: ListAPIInfo = {
      api: `${node.segment} ${resourcePath}`,
      server: server[0],
      operationId: node.operationId!,
      isValid: true,
      reason: [],
      auth: auth,
      summary: node.summary ?? "",
      description: node.description ?? "",
    };
    operations.push(apiInfo);
  }

  if (node.children && node.children.length > 0) {
    for (const child of node.children) {
      const childOps: ListAPIInfo[] = extractOperations(child, server, security, securitySchemes);
      operations.push(...childOps);
    }
  }

  return operations;
}
