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
  WarningType,
  WarningResult,
  AuthType,
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
  KiotaTreeResult,
  OpenApiSpecVersion,
  SecurityRequirementObject,
  SecuritySchemeObject,
} from "@microsoft/kiota";
import * as fs from "fs-extra";
import tmp from "tmp";
import { createHash } from "crypto";
import path from "path";
import { getLocalizedString } from "./localizeUtils";

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
  platform?: string,
  updateExistingPlugin = false
): Promise<GenerateResult> {
  const allowAPIKeyAuth = platform !== Platform.VS;
  const allowBearerTokenAuth = platform !== Platform.VS;
  const allowOauth2 = platform !== Platform.VS;

  if (featureFlagManager.getBooleanValue(FeatureFlags.KiotaNPMIntegration)) {
    const warnings: WarningResult[] = [];
    const tmpWorkingDir = tmp.dirSync({ unsafeCleanup: true });
    const tmpOutputDir = path.join(tmpWorkingDir.name, "plugin");
    const manifest: TeamsAppManifest = await fs.readJSON(teamsManifestPath);

    const namespace = removeEnvsAndSpecialCharaters(manifest.name.short);
    const includePatterns: string[] = [];
    for (const operation of operations) {
      const [method, path] = operation.split(" ");
      includePatterns.push(`${path}#${method}`);
    }

    const treeInfo = await listAPITreeInfo(specPath, includePatterns);

    const operationInfos: ListAPIInfo[] = extractOperations(treeInfo);

    for (let i = 0; i < operationInfos.length; i++) {
      const operationId = operationInfos[i].operationId;

      if (!operationId) {
        warnings.push({
          type: WarningType.OperationIdMissing,
          content: Utils.format(ConstantString.MissingOperationId, operationInfos[i].api),
          data: operationInfos[i].api,
        });
      } else {
        const containsSpecialCharacters = /[^a-zA-Z0-9_]/.test(operationId);
        const safeOperationId = operationId.replace(/[^a-zA-Z0-9]/g, "_");
        if (containsSpecialCharacters) {
          warnings.push({
            type: WarningType.OperationIdContainsSpecialCharacters,
            content: Utils.format(
              ConstantString.OperationIdContainsSpecialCharacters,
              operationId,
              safeOperationId
            ),
            data: operationId,
          });
        }
      }

      const authScheme = operationInfos[i].auth?.authScheme;
      if (authScheme) {
        if (!isAuthTypeSupported(authScheme, allowAPIKeyAuth, allowBearerTokenAuth, allowOauth2)) {
          warnings.push({
            type: WarningType.UnsupportedAuthType,
            content: Utils.format(ConstantString.AuthTypeIsNotSupported, operationId),
            data: operationId,
          });
        }
      }

      // TODO: add logs when kiota update the spec version, wait for kiota api
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

    const extname = path.extname(outputAPISpecPath);

    if (!updateExistingPlugin) {
      const outputSpecWithoutExt = path.join(
        path.dirname(outputAPISpecPath),
        path.basename(outputAPISpecPath, extname)
      );
      outputAPISpecPath = outputSpecWithoutExt + ".yaml";
    }

    await fs.copyFile(apiSpecPath, outputAPISpecPath);

    const adaptiveCardsFolder = path.join(path.dirname(apiSpecPath), "adaptiveCards");
    const destAdaptiveCardsFolder = path.join(path.dirname(outputAIPluginPath), "adaptiveCards");

    if (await fs.pathExists(adaptiveCardsFolder)) {
      await fs.copy(adaptiveCardsFolder, destAdaptiveCardsFolder, {
        overwrite: !updateExistingPlugin,
        errorOnExist: false,
      });
    }

    const relativePath = path.relative(path.dirname(outputAIPluginPath), outputAPISpecPath);
    const normalizedPath = relativePath.replace(/\\/g, "/");
    const generatedPluginManifest = (await fs.readJSON(pluginPath)) as PluginManifestSchema;

    if (!updateExistingPlugin) {
      const originalSpecFolder = path.join(tmpWorkingDir.name, `.kiota/documents/${namespace}/`);
      const files = await fs.readdir(originalSpecFolder);
      const originalSpecFilename = files[0];
      const originalSpecFile = path.join(originalSpecFolder, originalSpecFilename);

      const outputOriginalSpecPath = outputAPISpecPath + ".original";
      await fs.copyFile(originalSpecFile, outputOriginalSpecPath);
      generatedPluginManifest.runtimes?.forEach((runtime) => {
        (runtime as RuntimeObjectOpenapi).spec.url = normalizedPath;
      });
      await fs.writeJson(outputAIPluginPath, generatedPluginManifest, { spaces: 4 });
    } else {
      const existingPluginManifest = (await fs.readJSON(
        outputAIPluginPath
      )) as PluginManifestSchema;

      const functionNamesToRemove = new Set<string>();
      existingPluginManifest.runtimes?.forEach((runtime) => {
        const runtimeObj = runtime as RuntimeObjectOpenapi;
        if (runtimeObj.spec.url === normalizedPath && runtimeObj.run_for_functions) {
          runtimeObj.run_for_functions.forEach((name) => functionNamesToRemove.add(name));
        }
      });

      if (existingPluginManifest.functions) {
        existingPluginManifest.functions = existingPluginManifest.functions
          .filter((f) => !functionNamesToRemove.has(f.name))
          .concat(generatedPluginManifest.functions ?? []);
      } else {
        existingPluginManifest.functions = generatedPluginManifest.functions ?? [];
      }

      const runtimes = existingPluginManifest.runtimes;
      existingPluginManifest.runtimes = runtimes?.filter((r) => r.spec.url !== normalizedPath);

      for (const runtime of generatedPluginManifest.runtimes!) {
        (runtime as RuntimeObjectOpenapi).spec.url = normalizedPath;
        existingPluginManifest.runtimes?.push(runtime);
      }

      await fs.writeJson(outputAIPluginPath, existingPluginManifest, { spaces: 4 });
    }

    await parseAndUpdatePluginManifestForKiota(outputAIPluginPath, true);
    return {
      allSuccess: true,
      warnings: warnings,
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
        const parts = registrationId.split("_");
        const authName = parts.slice(0, -2).join("_");
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

export async function listAPIInfo(
  specPath: string,
  platform?: string
): Promise<
  ListAPIResult & {
    specVersion?: OpenApiSpecVersion;
  }
> {
  const allowAPIKeyAuth = platform !== Platform.VS;
  const allowBearerTokenAuth = platform !== Platform.VS;
  const allowOauth2 = platform !== Platform.VS;

  if (featureFlagManager.getBooleanValue(FeatureFlags.KiotaNPMIntegration)) {
    const treeInfo = await listAPITreeInfo(specPath);

    const operations: ListAPIInfo[] = extractOperations(treeInfo);

    for (const operation of operations) {
      if (!operation.server) {
        operation.reason.push(ErrorType.NoServerInformation);
      } else {
        const serverValidateResult = Utils.checkServerUrl([{ url: operation.server }], true);
        operation.reason.push(...serverValidateResult.map((item) => item.type));
      }
      if (operation.reason.length > 0) {
        operation.isValid = false;
      }
    }
    return {
      specVersion: treeInfo.specVersion,
      allAPICount: operations.length,
      validAPICount: operations.filter((api) => api.isValid).length,
      APIs: operations,
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
    let apiInfo: ListAPIResult & { specVersion?: OpenApiSpecVersion };
    try {
      apiInfo = await listAPIInfo(specPath, platform);
    } catch (e) {
      return {
        status: ValidationStatus.Error,
        warnings: [],
        errors: [
          {
            type: ErrorType.SpecNotValid,
            content: getLocalizedString(
              "error.daSpecParser.InvalidSpecError",
              (e as Error).toString()
            ),
          },
        ],
        specHash: "",
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
        specHash: "",
      };
    }

    const result: ValidateResult = {
      status: ValidationStatus.Valid,
      warnings: [],
      errors: [],
      specHash: "",
    };

    if (apiInfo.specVersion === OpenApiSpecVersion.V2_0) {
      result.warnings.push({
        type: WarningType.ConvertSwaggerToOpenAPI,
        content: ConstantString.ConvertSwaggerToOpenAPI,
      });
    }

    const serverUrl = apiInfo.APIs.find((api) => api.isValid)?.server;
    if (serverUrl) {
      const serverString = JSON.stringify(serverUrl);
      result.specHash = createHash("sha256").update(serverString).digest("hex");
    }

    return result;
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

function extractOperations(treeInfo: KiotaTreeResult | undefined): ListAPIInfo[] {
  let operations: ListAPIInfo[] = [];
  if (treeInfo && treeInfo.rootNode) {
    operations = traverseTreeNodeForOperations(
      treeInfo.rootNode,
      treeInfo.servers ?? [],
      treeInfo.security ?? [],
      treeInfo.securitySchemes ?? {}
    );
  }
  return operations;
}

function traverseTreeNodeForOperations(
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
    const normalizedPath = node.path.replace(/\\/g, "/");
    const lastHashIndex = normalizedPath.lastIndexOf("#");
    const resourcePath = normalizedPath.substring(0, lastHashIndex);

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

    if (node.selected) {
      operations.push(apiInfo);
    }
  }

  if (node.children && node.children.length > 0) {
    for (const child of node.children) {
      const childOps: ListAPIInfo[] = traverseTreeNodeForOperations(
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

function isAuthTypeSupported(
  authScheme: AuthType,
  allowAPIKeyAuth: boolean,
  allowBearerTokenAuth: boolean,
  allowOauth2: boolean
): boolean {
  return (
    (allowAPIKeyAuth && Utils.isAPIKeyAuthButNotInCookie(authScheme)) ||
    (allowOauth2 && Utils.isOAuthWithAuthCodeFlow(authScheme)) ||
    (allowBearerTokenAuth && Utils.isBearerTokenAuth(authScheme))
  );
}
