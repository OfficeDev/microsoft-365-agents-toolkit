// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  ListAPIInfo,
  ListAPIResult,
  ParseOptions,
  ProjectType,
  SpecParser,
  AuthInfo,
  ValidateResult,
  ValidationStatus,
  ErrorType,
  Utils,
  InvalidAPIInfo,
  AdaptiveCardUpdateStrategy,
  GenerateResult,
  ConstantString,
} from "@microsoft/m365-spec-parser";
import {
  Platform,
  PluginManifestSchema,
  RuntimeObjectOpenapi,
  TeamsAppManifest,
} from "@microsoft/teamsfx-api";
import { featureFlagManager, FeatureFlags } from "./featureFlags";
import { kiotageneratePlugin, listAPITreeInfo } from "./kiotaClient";
import {
  KiotaOpenApiNode,
  SecurityRequirementObject,
  SecuritySchemeObject,
} from "@microsoft/kiota";
import * as fs from "fs-extra";
import tmp from "tmp";
import { createHash } from "crypto";
import path from "path";

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

export async function generatePlugin(
  specPath: string,
  teamsManifestPath: string,
  outputAPISpecPath: string,
  outputAIPluginPath: string,
  operations: string[],
  adaptiveCardUpdateStrategy: AdaptiveCardUpdateStrategy,
  platform?: string
): Promise<GenerateResult> {
  const allowAPIKeyAuth = platform !== Platform.VS;
  const allowBearerTokenAuth = platform !== Platform.VS;
  const allowOauth2 = platform !== Platform.VS;

  if (featureFlagManager.getBooleanValue(FeatureFlags.KiotaNPMIntegration)) {
    const tmpWorkingDir = tmp.dirSync({ unsafeCleanup: true });
    const tmpOutputDir = path.join(tmpWorkingDir.name, "plugin");
    const manifest: TeamsAppManifest = await fs.readJSON(teamsManifestPath);
    const namespace = removeEnvsAndSpecialCharaters(manifest.name.short);
    const includePatterns: string[] = [];
    for (const operation of operations) {
      const [method, path] = operation.split(" ");
      includePatterns.push(`${path}#${method}`);
    }

    const kiotaGenerateResult = await kiotageneratePlugin(
      specPath,
      tmpOutputDir,
      namespace,
      tmpWorkingDir.name,
      undefined,
      undefined,
      includePatterns,
      []
    );

    const apiSpecPath = kiotaGenerateResult.openAPISpec;
    const pluginPath = kiotaGenerateResult.aiPlugin;
    await fs.copyFile(apiSpecPath, outputAPISpecPath);
    await fs.copyFile(pluginPath, outputAIPluginPath);

    const relativePath = path.relative(path.dirname(outputAIPluginPath), outputAPISpecPath);
    const normalizedPath = relativePath.replace(/\\/g, "/");
    const pluginManifest = (await fs.readJSON(outputAIPluginPath)) as PluginManifestSchema;
    pluginManifest.runtimes?.forEach((runtime) => {
      if ((runtime as RuntimeObjectOpenapi).spec) {
        (runtime as RuntimeObjectOpenapi).spec.url = normalizedPath;
      }
    });
    await fs.writeJson(outputAIPluginPath, pluginManifest, { spaces: 4 });

    await parseAndUpdatePluginManifestForKiota(outputAIPluginPath, true);
    return {
      allSuccess: true,
      warnings: [],
    };
  }

  const options: ParseOptions = {
    ...daProjectConfig,
    allowAPIKeyAuth,
    allowBearerTokenAuth,
    allowOauth2,
  };

  const parser = new SpecParser(specPath, options);

  const result = await parser.generateForCopilot(
    teamsManifestPath,
    operations,
    outputAPISpecPath,
    outputAIPluginPath,
    undefined,
    undefined,
    adaptiveCardUpdateStrategy
  );
  return result;
}

export async function parseAndUpdatePluginManifestForKiota(
  pluginManifestPath: string,
  updatePlaceholder: boolean
): Promise<
  { authName: string; authType: "apiKey" | "oauth2"; registrationId: string; specPath: string }[]
> {
  const authData: {
    authName: string;
    authType: "apiKey" | "oauth2";
    registrationId: string;
    specPath: string;
  }[] = [];
  const pluginManifest = (await fs.readJSON(pluginManifestPath)) as PluginManifestSchema;
  pluginManifest.runtimes?.forEach((runtime) => {
    if ((runtime as RuntimeObjectOpenapi).auth) {
      const auth = (runtime as RuntimeObjectOpenapi).auth!;
      if (
        auth.reference_id &&
        auth.reference_id.match(/^{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*}/g) &&
        auth.type !== "None"
      ) {
        const registrationId = auth.reference_id.replace(/[{}]/g, "");
        const authName = registrationId.split("_")[0];
        const newReferenceId = authName.toUpperCase() + "_" + ConstantString.RegistrationIdPostfix;
        authData.push({
          authName: authName,
          authType: auth.type === "ApiKeyPluginVault" ? "apiKey" : "oauth2",
          registrationId: newReferenceId,
          specPath: runtime.spec.url as string,
        });
        if (updatePlaceholder) {
          auth.reference_id = `\$\{\{${newReferenceId}\}\}`;
        }
      }
    }
  });

  if (updatePlaceholder && authData.length > 0) {
    await fs.writeJson(pluginManifestPath, pluginManifest, { spaces: 4 });
  }
  return authData;
}

export async function listAPIInfo(specPath: string, platform?: string): Promise<ListAPIResult> {
  const allowAPIKeyAuth = platform !== Platform.VS;
  const allowBearerTokenAuth = platform !== Platform.VS;
  const allowOauth2 = platform !== Platform.VS;

  if (featureFlagManager.getBooleanValue(FeatureFlags.KiotaNPMIntegration)) {
    const treeInfo = await listAPITreeInfo(specPath);

    if (treeInfo && treeInfo.rootNode) {
      const operations: ListAPIInfo[] = extractOperations(
        treeInfo.rootNode,
        treeInfo.servers ?? [],
        treeInfo.security ?? [],
        treeInfo.securitySchemes ?? {}
      );

      for (const operation of operations) {
        if (!operation.server) {
          operation.reason.push(ErrorType.NoServerInformation);
        } else {
          const serverValidateResult = Utils.checkServerUrl([{ url: operation.server }], true);
          operation.reason.push(...serverValidateResult.map((item) => item.type));
        }

        if (operation.auth) {
          if (operation.auth?.authScheme.type === "multipleAuth") {
            operation.reason.push(ErrorType.MultipleAuthNotSupported);
          } else if (
            !(
              (allowAPIKeyAuth && Utils.isAPIKeyAuth(operation.auth.authScheme)) ||
              (allowOauth2 && Utils.isOAuthWithAuthCodeFlow(operation.auth.authScheme)) ||
              (allowBearerTokenAuth && Utils.isBearerTokenAuth(operation.auth.authScheme))
            )
          ) {
            operation.reason.push(ErrorType.AuthTypeIsNotSupported);
          }
        }

        if (operation.reason.length > 0) {
          operation.isValid = false;
        }
      }

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
    allowAPIKeyAuth,
    allowBearerTokenAuth,
    allowOauth2,
  };

  const parser = new SpecParser(specPath, options);

  return await parser.list();
}

export async function validateOpenAPISpec(
  specPath: string,
  platform?: string
): Promise<ValidateResult> {
  if (featureFlagManager.getBooleanValue(FeatureFlags.KiotaNPMIntegration)) {
    let hash = "";
    let apiInfo: ListAPIResult;
    try {
      apiInfo = await listAPIInfo(specPath, platform);
    } catch (e) {
      return {
        status: ValidationStatus.Error,
        warnings: [],
        errors: [{ type: ErrorType.SpecNotValid, content: (e as Error).toString() }],
        specHash: hash,
      };
    }

    if (apiInfo.allAPICount === 0 || apiInfo.validAPICount === 0) {
      const data = [];
      for (const info of apiInfo.APIs) {
        const apiInvalidReason: InvalidAPIInfo = { api: info.api, reason: info.reason };
        data.push(apiInvalidReason);
      }

      return {
        status: ValidationStatus.Error,
        warnings: [],
        errors: [{ type: ErrorType.NoSupportedApi, content: "", data: data }],
        specHash: hash,
      };
    }

    const serverUrl = apiInfo.APIs.find((api) => api.isValid)?.server;
    if (serverUrl) {
      const serverString = JSON.stringify(serverUrl);
      hash = createHash("sha256").update(serverString).digest("hex");
    }

    return {
      status: ValidationStatus.Valid,
      warnings: [],
      errors: [],
      specHash: hash,
    };
  }

  const options: ParseOptions = {
    ...daProjectConfig,
    allowAPIKeyAuth: platform !== Platform.VS,
    allowBearerTokenAuth: platform !== Platform.VS,
    allowOauth2: platform !== Platform.VS,
  };

  const parser = new SpecParser(specPath, options);
  return await parser.validate();
}

function removeEnvsAndSpecialCharaters(str: string): string {
  const placeHolderReg = /\${{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*}}/g;
  const matches = placeHolderReg.exec(str);
  let newStr = str;
  if (matches != null) {
    newStr = newStr.replace(matches[0], "");
  }
  return newStr.toLowerCase().replace(/[^a-z0-9]/g, "");
}

function extractOperations(
  node: KiotaOpenApiNode,
  parentServer: string[],
  parentSecurity: SecurityRequirementObject[],
  securitySchemes: {
    [key: string]: SecuritySchemeObject;
  }
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
      const childOps: ListAPIInfo[] = extractOperations(
        child,
        server,
        security ?? [],
        securitySchemes
      );
      operations.push(...childOps);
    }
  }

  return operations;
}
