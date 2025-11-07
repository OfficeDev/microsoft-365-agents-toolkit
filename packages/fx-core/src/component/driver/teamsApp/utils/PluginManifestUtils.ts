// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  Colors,
  DefaultApiSpecJsonFileName,
  DefaultApiSpecYamlFileName,
  FxError,
  IPlugin,
  ManifestUtil,
  Platform,
  PluginManifestSchema,
  FunctionObject,
  RuntimeObjectLocalplugin,
  Result,
  TeamsAppManifest,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { FileNotFoundError, JSONSyntaxError } from "../../../../error/common";
import stripBom from "strip-bom";
import path from "path";
import { manifestUtils } from "./ManifestUtils";
import { getResolvedManifest } from "./utils";
import { AppStudioResultFactory } from "../results";
import { AppStudioError } from "../errors";
import { getDefaultString, getLocalizedString } from "../../../../common/localizeUtils";
import { PluginManifestValidationResult } from "../interfaces/ValidationResult";
import { SummaryConstant } from "../../../configManager/constant";
import { EOL } from "os";
import { ManifestType } from "../../../utils/envFunctionUtils";
import { DriverContext } from "../../interface/commonArgs";
import { isJsonSpecFile } from "../../../../common/utils";
import { featureFlagManager, FeatureFlags } from "../../../../common/featureFlags";
import { ODRProvider, ODRServer, ODRTool } from "../../../utils/odrProvider";
import { LocalMcpPrefix } from "../../../constants";

export class PluginManifestUtils {
  public async readPluginManifestFile(
    path: string
  ): Promise<Result<PluginManifestSchema, FxError>> {
    if (!(await fs.pathExists(path))) {
      return err(new FileNotFoundError("PluginManifestUtils", path));
    }
    // Be compatible with UTF8-BOM encoding
    // Avoid Unexpected token error at JSON.parse()
    let content = await fs.readFile(path, { encoding: "utf-8" });
    content = stripBom(content);

    try {
      const manifest = JSON.parse(content) as PluginManifestSchema;
      return ok(manifest);
    } catch (e) {
      return err(new JSONSyntaxError(path, e, "PluginManifestUtils"));
    }
  }

  /**
   * Get plugin manifest with env value filled.
   * @param path path of declaraitve Copilot
   * @returns resolved manifest
   */
  public async getManifest(
    path: string,
    context: DriverContext
  ): Promise<Result<PluginManifestSchema, FxError>> {
    const manifestRes = await this.readPluginManifestFile(path);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }
    // Add environment variable keys to telemetry
    const resolvedManifestRes = await getResolvedManifest(
      JSON.stringify(manifestRes.value),
      path,
      ManifestType.PluginManifest,
      context
    );

    if (resolvedManifestRes.isErr()) {
      return err(resolvedManifestRes.error);
    }
    const resolvedManifestString = resolvedManifestRes.value;
    return ok(JSON.parse(resolvedManifestString));
  }

  public async validateAgainstSchema(
    plugin: IPlugin,
    path: string,
    context: DriverContext
  ): Promise<Result<PluginManifestValidationResult, FxError>> {
    const manifestRes = await this.getManifest(path, context);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }

    try {
      const schemaErrors = await ManifestUtil.validateManifest(manifestRes.value);
      const localPluginErrors = await this.validateLocalPluginRuntimes(manifestRes.value, context);
      const allErrors = [...schemaErrors, ...localPluginErrors];

      return ok({
        id: plugin.id,
        filePath: path,
        validationResult: allErrors,
      });
    } catch (e: any) {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.ValidationFailedError.name,
          AppStudioError.ValidationFailedError.message([
            getLocalizedString(
              "error.appstudio.validateFetchSchemaFailed",
              manifestRes.value.$schema,
              e.message
            ),
          ])
        )
      );
    }
  }

  public async getApiSpecFilePathFromTeamsManifest(
    manifest: TeamsAppManifest,
    manifestPath: string
  ): Promise<Result<string[], FxError>> {
    const pluginFilePathRes = await manifestUtils.getPluginFilePath(manifest, manifestPath);
    if (pluginFilePathRes.isErr()) {
      return err(pluginFilePathRes.error);
    }
    const pluginFilePath = pluginFilePathRes.value;
    const pluginContentRes = await this.readPluginManifestFile(pluginFilePath);
    if (pluginContentRes.isErr()) {
      return err(pluginContentRes.error);
    }
    const apiSpecFiles = await this.getApiSpecFilePathFromPlugin(
      pluginContentRes.value,
      pluginFilePath
    );
    return ok(apiSpecFiles);
  }

  public logValidationErrors(
    validationRes: PluginManifestValidationResult,
    platform: Platform
  ): string | Array<{ content: string; color: Colors }> {
    const validationErrors = validationRes.validationResult;
    const filePath = validationRes.filePath;
    if (validationErrors.length === 0) {
      return "";
    }

    if (platform !== Platform.CLI) {
      const errors = validationErrors
        .map((error: string) => {
          return `${SummaryConstant.Failed} ${error}`;
        })
        .join(EOL);
      return (
        getLocalizedString("driver.teamsApp.summary.validatePluginManifest.checkPath", filePath) +
        EOL +
        errors
      );
    } else {
      const outputMessage = [];
      outputMessage.push({
        content:
          getDefaultString("driver.teamsApp.summary.validatePluginManifest.checkPath", filePath) +
          "\n",
        color: Colors.BRIGHT_WHITE,
      });
      validationErrors.map((error: string) => {
        outputMessage.push({ content: `${SummaryConstant.Failed} `, color: Colors.BRIGHT_RED });
        outputMessage.push({
          content: `${error}\n`,
          color: Colors.BRIGHT_WHITE,
        });
      });

      return outputMessage;
    }
  }

  public async getDefaultNextAvailableApiSpecPath(
    apiSpecPath: string,
    apiSpecFolder: string,
    isKiotaIntegration = false
  ) {
    let isYaml = false;
    try {
      isYaml = !(await isJsonSpecFile(apiSpecPath));
    } catch (e) {}

    let openApiSpecFileName =
      isYaml || featureFlagManager.getBooleanValue(FeatureFlags.KiotaNPMIntegration)
        ? DefaultApiSpecYamlFileName
        : DefaultApiSpecJsonFileName;
    // Check if the default file name already exists
    if (!(await fs.pathExists(path.join(apiSpecFolder, openApiSpecFileName)))) {
      return path.join(apiSpecFolder, openApiSpecFileName);
    }

    const openApiSpecFileNamePrefix = openApiSpecFileName.split(".")[0];
    const openApiSpecFileType = openApiSpecFileName.split(".")[1];
    let apiSpecFileNameSuffix = 1;
    openApiSpecFileName = this.getApiSpecFileName(
      openApiSpecFileNamePrefix,
      openApiSpecFileType,
      apiSpecFileNameSuffix,
      isKiotaIntegration
    );

    while (await fs.pathExists(path.join(apiSpecFolder, openApiSpecFileName))) {
      apiSpecFileNameSuffix++;
      openApiSpecFileName = this.getApiSpecFileName(
        openApiSpecFileNamePrefix,
        openApiSpecFileType,
        apiSpecFileNameSuffix,
        isKiotaIntegration
      );
    }
    const openApiSpecFilePath = path.join(apiSpecFolder, openApiSpecFileName);

    return openApiSpecFilePath;
  }

  getApiSpecFileName(
    openApiSpecFileNamePrefix: string,
    openApiSpecFileType: string,
    apiSpecFileNameSuffix: number,
    isKiotaIntegration: boolean
  ): string {
    let openApiSpecFileName;
    if (isKiotaIntegration) {
      const apiSpecNameSplit = openApiSpecFileNamePrefix.split("-");
      openApiSpecFileName = `${apiSpecNameSplit[0]}_${apiSpecFileNameSuffix}-${apiSpecNameSplit[1]}.${openApiSpecFileType}`;
    } else {
      openApiSpecFileName = `${openApiSpecFileNamePrefix}_${apiSpecFileNameSuffix}.${openApiSpecFileType}`;
    }
    return openApiSpecFileName;
  }

  async getApiSpecFilePathFromPlugin(
    plugin: PluginManifestSchema,
    pluginPath: string
  ): Promise<string[]> {
    const runtimes = plugin.runtimes;
    const files: string[] = [];
    if (!runtimes) {
      return files;
    }
    for (const runtime of runtimes) {
      if (runtime.type === "OpenApi" && runtime.spec?.url) {
        const specFile = path.resolve(path.dirname(pluginPath), runtime.spec.url);
        if (await fs.pathExists(specFile)) {
          files.push(specFile);
        }
      }
    }

    return files;
  }

  /**
   * Validate LocalPlugin runtimes in the manifest.
   * Performs structural validation and optionally validates against MCP servers.
   * @param manifest The plugin manifest to validate
   * @param context Driver context for logging
   * @returns Array of validation error strings
   */
  public async validateLocalPluginRuntimes(
    manifest: PluginManifestSchema,
    context: DriverContext
  ): Promise<string[]> {
    const errors: string[] = [];

    if (!manifest.runtimes) {
      return errors;
    }

    const localPluginRuntimes = manifest.runtimes.filter(
      (rt): rt is RuntimeObjectLocalplugin => rt.type === "LocalPlugin"
    );

    if (localPluginRuntimes.length === 0) {
      return errors;
    }

    // local_endpoint format
    localPluginRuntimes.forEach((rt) => {
      const actualIdx = manifest.runtimes?.indexOf(rt) ?? -1;
      if (!rt.spec.local_endpoint.startsWith(LocalMcpPrefix)) {
        errors.push(
          `/runtimes/${actualIdx}/spec/local_endpoint must start with "${LocalMcpPrefix}" for LocalPlugin type. ` +
            `Found: "${rt.spec.local_endpoint}"`
        );
      }
    });

    // All functions must be referenced by some runtime
    const allReferencedFunctions = new Set<string>();
    manifest.runtimes.forEach((rt) => {
      rt.run_for_functions?.forEach((fn) => allReferencedFunctions.add(fn));
    });

    const unreferencedFunctions = (manifest.functions || [])
      .map((f) => f.name)
      .filter((name) => !allReferencedFunctions.has(name));

    if (unreferencedFunctions.length > 0) {
      errors.push(`Functions not referenced by any runtime: ${unreferencedFunctions.join(", ")}`);
    }

    // MCP tool validation
    if (localPluginRuntimes.length > 0) {
      if (process.platform !== "win32") {
        throw new Error("MCP validation failed: ODR is only available on Windows platform.");
      }

      try {
        const mcpServers = await ODRProvider.listServers();
        if (mcpServers.length > 0) {
          const mcpValidationErrors = this.validateAgainstMCPServers(
            manifest,
            localPluginRuntimes,
            mcpServers
          );
          errors.push(...mcpValidationErrors);
        } else {
          throw new Error("No MCP servers found. MCP tool validation failed.");
        }
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        throw new Error(`MCP validation failed: ${errorMessage}`);
      }
    }

    return errors;
  }

  /**
   * Validate LocalPlugin runtimes against actual MCP servers.
   * @param manifest The plugin manifest
   * @param localPluginRuntimes The LocalPlugin runtimes to validate
   * @param mcpServers The available MCP servers from ODR
   * @returns Array of validation error strings
   */
  private validateAgainstMCPServers(
    manifest: PluginManifestSchema,
    localPluginRuntimes: RuntimeObjectLocalplugin[],
    mcpServers: ODRServer[]
  ): string[] {
    const errors: string[] = [];

    localPluginRuntimes.forEach((runtime) => {
      const actualIdx = manifest.runtimes?.indexOf(runtime) ?? -1;

      const mcpIdentifier = runtime.spec.local_endpoint.replace(
        new RegExp(`^${LocalMcpPrefix}`),
        ""
      );

      // Find matching MCP server
      const mcpServer = mcpServers.find((s) => s.identifier === mcpIdentifier);

      if (!mcpServer) {
        const availableServers =
          mcpServers.length > 0
            ? mcpServers.map((s) => s.identifier).join(", ")
            : "(none available)";
        errors.push(
          `/runtimes/${actualIdx}/spec/local_endpoint: ` +
            `MCP server "${mcpIdentifier}" not found. Available servers: ${availableServers}`
        );
        return; // Skip further validation for this runtime
      }

      const mcpToolNames = mcpServer.tools.map((t) => t.name);

      // run_for_functions must exist in MCP server
      runtime.run_for_functions?.forEach((funcName, funcIdx) => {
        if (!mcpToolNames.includes(funcName)) {
          errors.push(
            `/runtimes/${actualIdx}/run_for_functions[${funcIdx}]: ` +
              `Tool "${funcName}" not found in MCP server "${mcpServer.display_name}". ` +
              `Available tools: ${mcpToolNames.join(", ")}`
          );
        }
      });

      // Function parameters must match MCP tools
      runtime.run_for_functions?.forEach((funcName) => {
        const manifestFunc = manifest.functions?.find((f) => f.name === funcName);
        const mcpTool = mcpServer.tools.find((t) => t.name === funcName);

        if (manifestFunc && mcpTool) {
          const paramErrors = this.compareParameters(manifestFunc, mcpTool, funcName);
          errors.push(...paramErrors);
        }
      });
    });

    return errors;
  }

  /**
   * Compare function parameters between manifest and MCP tool definition.
   * @param manifestFunc The function definition from the manifest
   * @param mcpTool The tool definition from MCP server
   * @param funcName The function name for error messages
   * @returns Array of validation error strings
   */
  private compareParameters(
    manifestFunc: FunctionObject,
    mcpTool: ODRTool,
    funcName: string
  ): string[] {
    const errors: string[] = [];
    const manifestParams = manifestFunc.parameters?.properties || {};
    const mcpParams = mcpTool.inputSchema?.properties || {};
    const manifestRequired = manifestFunc.parameters?.required || [];
    const mcpRequired = mcpTool.inputSchema?.required || [];

    // Check for extra parameters in manifest
    Object.keys(manifestParams).forEach((prop) => {
      if (!mcpParams[prop]) {
        errors.push(
          `/functions["${funcName}"]/parameters/properties/${prop}: ` +
            `Extra parameter not defined in MCP server. Remove this parameter or update MCP server.`
        );
      }
    });

    // Check for missing parameters from MCP
    Object.keys(mcpParams).forEach((prop) => {
      if (!manifestParams[prop]) {
        errors.push(
          `/functions["${funcName}"]/parameters/properties/${prop}: ` +
            `Missing parameter defined in MCP server. Add this parameter.`
        );
      }
    });

    // Check for required array differences
    const extraRequired = manifestRequired.filter((r) => !mcpRequired.includes(r));
    const missingRequired = mcpRequired.filter((r: string) => !manifestRequired.includes(r));

    if (extraRequired.length > 0) {
      errors.push(
        `/functions["${funcName}"]/parameters/required: ` +
          `Extra required parameters not in MCP server: ${extraRequired.join(", ")}`
      );
    }

    if (missingRequired.length > 0) {
      errors.push(
        `/functions["${funcName}"]/parameters/required: ` +
          `Missing required parameters from MCP server: ${(missingRequired as string[]).join(", ")}`
      );
    }

    // Deep type checking for matching parameters
    Object.keys(manifestParams).forEach((prop) => {
      if (mcpParams[prop]) {
        const manifestParam = manifestParams[prop];
        const mcpParam = mcpParams[prop];

        // Type check
        const manifestType = manifestParam.type;
        const mcpType = mcpParam.type;

        if (manifestType !== mcpType) {
          errors.push(
            `/functions["${funcName}"]/parameters/properties/${prop}/type: ` +
              `Type mismatch. Manifest has "${String(manifestType)}", MCP server has "${String(
                mcpType
              )}"`
          );
        }

        // Check enums
        const manifestEnum = manifestParam.enum;
        const mcpEnum = mcpParam.enum;
        if (manifestEnum || mcpEnum) {
          const manifestEnumSorted = manifestEnum ? [...manifestEnum].sort() : undefined;
          const mcpEnumSorted = mcpEnum ? [...mcpEnum].sort() : undefined;
          const enumMatch = JSON.stringify(manifestEnumSorted) === JSON.stringify(mcpEnumSorted);

          if (!enumMatch) {
            errors.push(
              `/functions["${funcName}"]/parameters/properties/${prop}/enum: ` +
                `Enum mismatch. Manifest: ${JSON.stringify(manifestEnum)}, ` +
                `MCP: ${JSON.stringify(mcpEnum)}`
            );
          }
        }
      }
    });

    return errors;
  }
}

export const pluginManifestUtils = new PluginManifestUtils();
