// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";
import { AuthType, ListAPIResult, ProjectType, SpecParser } from "@microsoft/m365-spec-parser";
import {
  DefaultApiSpecFolderName,
  FxError,
  Inputs,
  Platform,
  Result,
  Stage,
  SystemError,
  UserError,
  Warning,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import * as path from "path";
import { listAPIInfo } from "../common/daSpecParser";
import { ErrorContextMW, TOOLS, createContext } from "../common/globalVars";
import { getDefaultString, getLocalizedString } from "../common/localizeUtils";
import { TelemetryEvent } from "../common/telemetry";
import { MetadataV3, MetadataV4 } from "../common/versionMetadata";
import { ActionInjector } from "../component/configManager/actionInjector";
import { LocalMcpPrefix } from "../component/constants";
import { AppStudioError } from "../component/driver/teamsApp/errors";
import { AppStudioResultFactory } from "../component/driver/teamsApp/results";
import { copilotGptManifestUtils } from "../component/driver/teamsApp/utils/CopilotGptManifestUtils";
import { manifestUtils } from "../component/driver/teamsApp/utils/ManifestUtils";
import { pluginManifestUtils } from "../component/driver/teamsApp/utils/PluginManifestUtils";
import { normalizePath } from "../component/driver/teamsApp/utils/utils";
import { addExistingPlugin } from "../component/generator/declarativeAgent/helper";
import {
  generateFromApiSpec,
  generateScaffoldingSummary,
  getParserOptions,
} from "../component/generator/openApiSpec/helper";
import { QuestionMW } from "../component/middleware/questionMW";
import { outputScaffoldingWarningMessage } from "../component/utils/common";
import { resolveMCPOAuthMetadata } from "../component/utils/mcpToolFetcher";
import { pathUtils } from "../component/utils/pathUtils";
import { UserCancelError } from "../error/common";
import { ActionStartOptions, QuestionNames } from "../question/constants";
import { ConcurrentLockerMW } from "./middleware/concurrentLocker";
import { ErrorHandlerMW } from "./middleware/errorHandler";

// Non-translatable CLI command template used in warning messages
const mcpAddActionHint =
  "atk add action --api-plugin-type mcp --mcp-da-server-url <server-url> --mcp-tools-file-path <path-to-tools-json> --interactive false";

export class FxCoreDeclarativeAgentPart {
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: Stage.installApp }),
    ErrorHandlerMW,
    QuestionMW("updateActionWithMCP"),
    ConcurrentLockerMW,
  ])
  async updateActionWithMCP(inputs: Inputs): Promise<Result<any, FxError>> {
    const context = createContext();
    const projectPath = inputs.projectPath;
    if (!projectPath) {
      throw new Error("projectPath is undefined"); // should never happen
    }
    const aiPluginFilePath = inputs[QuestionNames.PluginManifestFilePath] as string;
    if (!(await fs.pathExists(aiPluginFilePath))) {
      const error = new SystemError(
        "MCPForDAPluginManifestNotFound",
        "PluginManifestNotFound",
        getDefaultString("core.MCPForDA.pluginManifestNotFound", aiPluginFilePath),
        getLocalizedString("core.MCPForDA.pluginManifestNotFound", aiPluginFilePath)
      );
      return err(error);
    }
    const aiPluginFilePathRelative = path.basename(aiPluginFilePath);

    const mcpServerUrl = inputs[QuestionNames.MCPForDAServerUrl];
    const serverName = inputs[QuestionNames.MCPForDAServerName] as string;
    const mcpAuth = inputs[QuestionNames.MCPForDAAuth];
    const authType = inputs[QuestionNames.MCPForDAAuthType];

    let oauthAuthorizationUrl: string | undefined = undefined;
    let oauthTokenUrl: string | undefined = undefined;
    let oauthRefreshUrl: string | undefined = undefined;
    let registrationId: string | undefined = undefined;

    if (mcpAuth === "OAuthPluginVault") {
      try {
        registrationId = `MCP_DA_AUTH_ID_${serverName.toUpperCase()}`;
        if (authType === "oauth") {
          const metadata = await resolveMCPOAuthMetadata(
            inputs[QuestionNames.MCPForDAAuthMetadataUrl],
            inputs[QuestionNames.MCPForDAAuthWellKnownUrl]
          );
          oauthAuthorizationUrl = metadata.authorizationUrl;
          oauthTokenUrl = metadata.tokenUrl;
          oauthRefreshUrl = metadata.refreshUrl;
        }
      } catch (error: any) {
        void context.userInteraction.showMessage(
          "error",
          getLocalizedString("core.MCPForDA.mcpAuthMetadataMissingError", error.message),
          false
        );
      }
    }

    // 2. Read ai-plugin.json
    const aiPluginContent = await fs.readJSON(aiPluginFilePath);

    // For dynamic fetch tools, keep the functions empty and add runtime info
    // TODO: support dynamic fetch tools in the future
    const mcpToolsDetail = inputs[QuestionNames.MCPForDAAvailableTools];
    const mcpToolsSelected = inputs[QuestionNames.MCPForDAPreFetchTools];
    if (!mcpToolsDetail || !mcpToolsSelected) {
      const error = new UserError(
        "MCPForDAPreFetchToolsNotFound",
        "PreFetchToolsNotFound",
        getDefaultString("core.MCPForDA.preFetchToolsNotFound"),
        getLocalizedString("core.MCPForDA.preFetchToolsNotFound")
      );
      return err(error);
    }

    const toolsSelectedPrevious: string[] = [];
    aiPluginContent.runtimes
      .filter(
        (runtime: any) =>
          (runtime.type === "RemoteMCPServer" && runtime.spec.url === mcpServerUrl) ||
          runtime.type === "LocalPlugin"
      )
      .forEach((runtime: any) => {
        toolsSelectedPrevious.push(...runtime.run_for_functions);
      });
    aiPluginContent.functions = aiPluginContent.functions.filter(
      (func: any) => !toolsSelectedPrevious.includes(func.name)
    );
    aiPluginContent.functions = [
      ...aiPluginContent.functions,
      ...mcpToolsDetail
        .filter((tool: any) => mcpToolsSelected.includes(tool.name))
        .map((tool: any) => {
          if (inputs[QuestionNames.MCPLocalServerIdentifier] != null) {
            return {
              name: tool.name,
              description: tool.description,
              parameters: {
                type: tool.inputSchema.type || "object",
                properties: tool.inputSchema.properties,
                required: tool.inputSchema.required || [],
              },
            };
          } else {
            return {
              name: tool.name,
              description: tool.description,
            };
          }
        }),
    ];

    const matchedRuntime = aiPluginContent.runtimes.find(
      (runtime: any) => runtime.type === "RemoteMCPServer" && runtime.spec.url === mcpServerUrl
    );

    aiPluginContent.runtimes = aiPluginContent.runtimes.filter(
      (runtime: any) =>
        (runtime.type !== "RemoteMCPServer" && runtime.type !== "LocalPlugin") ||
        runtime.spec.url !== mcpServerUrl
    );

    if (inputs[QuestionNames.MCPLocalServerIdentifier] != null) {
      (aiPluginContent.runtimes as any[]).push({
        type: "LocalPlugin",
        spec: {
          local_endpoint: `${LocalMcpPrefix}${
            inputs[QuestionNames.MCPLocalServerIdentifier] as string
          }`,
        },
        run_for_functions: mcpToolsSelected,
      });
    } else {
      let mcpFile = matchedRuntime?.spec.mcp_tool_description?.file;
      if (!mcpFile) {
        mcpFile = "mcp-tools.json";
        let suffix = 1;
        while (await fs.pathExists(path.join(path.dirname(aiPluginFilePath), mcpFile))) {
          mcpFile = `mcp-tools-${suffix}.json`;
          suffix += 1;
        }
      }
      await fs.writeJSON(
        path.join(path.dirname(aiPluginFilePath), mcpFile),
        {
          tools: [
            ...mcpToolsDetail
              .filter((tool: any) => mcpToolsSelected.includes(tool.name))
              .map((tool: any) => {
                return {
                  ...tool,
                  title: (tool.name as string)
                    .replace(/_/g, " ")
                    .replace(/^./, (str) => str.toUpperCase()),
                };
              }),
          ],
        },
        { spaces: 4 }
      );
      (aiPluginContent.runtimes as any[]).push({
        type: "RemoteMCPServer",
        spec: {
          url: mcpServerUrl,
          mcp_tool_description: {
            file: mcpFile,
          },
        },
        run_for_functions: mcpToolsSelected,
        auth:
          mcpAuth === "OAuthPluginVault" && !!registrationId
            ? {
                type: "OAuthPluginVault",
                reference_id: `$\{\{${registrationId}\}\}`,
              }
            : {
                type: "None",
              },
      });
    }

    if (mcpAuth === "OAuthPluginVault" && !!registrationId) {
      // insert oauth info in teamsapp.yaml
      await ActionInjector.injectCreateOAuthActionForMCP(
        pathUtils.getYmlFilePath(projectPath) as string,
        authType,
        serverName,
        registrationId,
        mcpServerUrl,
        oauthAuthorizationUrl,
        oauthTokenUrl,
        oauthRefreshUrl
      );
    }
    void context.userInteraction
      .showMessage(
        "info",
        getLocalizedString("core.MCPForDA.updatePluginManifest", aiPluginFilePathRelative),
        false,
        "Provision"
      )
      .then((result) => {
        if (result.isOk() && result.value === "Provision") {
          void this.provisionResources(inputs);
        }
      });
    await fs.writeJSON(aiPluginFilePath, aiPluginContent, { spaces: 4 });
    void context.userInteraction.openFile?.(aiPluginFilePath);
    return ok(undefined);
  }

  // This method will be implemented by FxCore
  provisionResources(_inputs: Inputs): Promise<Result<any, FxError>> {
    throw new Error("Method not implemented.");
  }

  // These methods will be implemented by FxCore
  protected parseAuthNameAndScheme(
    _listResult: ListAPIResult,
    _inputs: Inputs
  ): { authName: string; authScheme: AuthType }[] {
    throw new Error("Method not implemented.");
  }

  // eslint-disable-next-line @typescript-eslint/require-await
  protected async updateAuthActionInYaml(
    _authName: string | undefined,
    _authScheme: AuthType | undefined,
    _projectPath: string,
    _apSpecPath: string,
    _pluginManifestPath: string,
    _forceToAddNew?: boolean
  ): Promise<void> {
    throw new Error("Method not implemented.");
  }

  /**
   * Add plugin
   */
  @hooks([
    ErrorContextMW({ component: "FxCore", stage: Stage.addPlugin }),
    ErrorHandlerMW,
    QuestionMW("addPlugin"),
    ConcurrentLockerMW,
  ])
  async addPlugin(inputs: Inputs): Promise<Result<undefined | any, FxError>> {
    if (!inputs.projectPath) {
      throw new Error("projectPath is undefined"); // should never happen
    }

    const context = createContext();
    const teamsManifestPath = inputs[QuestionNames.ManifestPath];
    const appPackageFolder = path.dirname(teamsManifestPath);
    const isGenerateFromApiSpec =
      inputs[QuestionNames.ActionType] === ActionStartOptions.apiSpec().id;
    const isGenerateFromMCP = inputs[QuestionNames.ActionType] === ActionStartOptions.mcp().id;

    // validate the project is valid for adding plugin
    const manifestRes = await manifestUtils._readAppManifest(teamsManifestPath);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }

    const teamsManifest = manifestRes.value;
    const declarativeGpt = teamsManifest.copilotExtensions
      ? teamsManifest.copilotExtensions.declarativeCopilots?.[0]
      : teamsManifest.copilotAgents?.declarativeAgents?.[0];
    if (!declarativeGpt?.file) {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.TeamsAppRequiredPropertyMissingError.name,
          AppStudioError.TeamsAppRequiredPropertyMissingError.message(
            "declarativeAgents",
            teamsManifestPath
          )
        )
      );
    }
    const gptManifestFilePathRes = await copilotGptManifestUtils.getManifestPath(teamsManifestPath);
    if (gptManifestFilePathRes.isErr()) {
      return err(gptManifestFilePathRes.error);
    }

    const declarativeCopilotManifestPath = gptManifestFilePathRes.value;

    const declarativeCopilotManifesRes = await copilotGptManifestUtils.readCopilotGptManifestFile(
      declarativeCopilotManifestPath
    );
    if (declarativeCopilotManifesRes.isErr()) {
      return err(declarativeCopilotManifesRes.error);
    }

    const declarativeCopilotManifest = declarativeCopilotManifesRes.value;
    let confirmMessage = getLocalizedString(
      "core.addApi.confirm",
      path.relative(inputs.projectPath, appPackageFolder)
    );

    // Will be used if generating from API spec
    let specParser: SpecParser | undefined = undefined;
    let authNameAndSchemes: { authName: string; authScheme: AuthType }[] = [];

    if (isGenerateFromApiSpec) {
      specParser = new SpecParser(
        inputs[QuestionNames.ApiSpecLocation].trim(),
        getParserOptions(ProjectType.Copilot, true)
      );
      const listResult = await listAPIInfo(inputs[QuestionNames.ApiSpecLocation].trim());
      authNameAndSchemes = this.parseAuthNameAndScheme(listResult, inputs);

      if (authNameAndSchemes.length > 0) {
        const doesLocalYamlPathExists = await fs.pathExists(
          path.join(inputs.projectPath, MetadataV3.localConfigFile)
        );
        confirmMessage = doesLocalYamlPathExists
          ? getLocalizedString(
              "core.addApi.confirm.localTeamsYaml",
              path.relative(inputs.projectPath, appPackageFolder),
              MetadataV4.localConfigFile,
              MetadataV4.configFile
            )
          : getLocalizedString(
              "core.addApi.confirm.teamsYaml",
              path.relative(inputs.projectPath, appPackageFolder),
              MetadataV4.configFile
            );
      }
    }

    // confirm

    const confirmRes = await context.userInteraction.showMessage(
      "warn",
      confirmMessage,
      true,
      getLocalizedString("core.addApi.continue")
    );

    if (confirmRes.isErr()) {
      return err(confirmRes.error);
    } else if (confirmRes.value !== getLocalizedString("core.addApi.continue")) {
      return err(new UserCancelError());
    }

    // find the next available action id
    let actionId = "";
    let suffix = 1;
    actionId = `action_${suffix}`;
    const existingActionIds = declarativeCopilotManifest.actions?.map((action) => action.id);
    while (existingActionIds?.includes(actionId)) {
      suffix += 1;
      actionId = `action_${suffix}`;
    }

    let destinationPluginManifestPath: string;
    // generate files
    if (isGenerateFromApiSpec && specParser) {
      destinationPluginManifestPath =
        await copilotGptManifestUtils.getDefaultNextAvailablePluginManifestPath(
          appPackageFolder,
          undefined
        );
      const destinationApiSpecPath = await pluginManifestUtils.getDefaultNextAvailableApiSpecPath(
        inputs[QuestionNames.ApiSpecLocation].trim(),
        path.join(appPackageFolder, DefaultApiSpecFolderName)
      );

      const generateRes = await generateFromApiSpec(
        specParser,
        teamsManifestPath,
        inputs,
        context,
        Stage.addPlugin,
        ProjectType.Copilot,
        {
          destinationApiSpecFilePath: destinationApiSpecPath,
          pluginManifestFilePath: destinationPluginManifestPath,
        },
        inputs[QuestionNames.ApiSpecLocation].trim()
      );
      if (generateRes.isErr()) {
        return err(generateRes.error);
      }

      const warnings = generateRes.value.warnings;
      if (warnings && warnings.length > 0) {
        const warnSummary = await generateScaffoldingSummary(
          warnings,
          manifestRes.value,
          path.relative(inputs.projectPath, destinationApiSpecPath),
          path.relative(inputs.projectPath, destinationPluginManifestPath),
          inputs.projectPath
        );
        context.logProvider.info(warnSummary + "\n");
      }

      const addActionRes = await copilotGptManifestUtils.addAction(
        declarativeCopilotManifestPath,
        actionId,
        normalizePath(path.relative(appPackageFolder, destinationPluginManifestPath), true)
      );
      if (addActionRes.isErr()) {
        return err(addActionRes.error);
      }

      for (const authNameAndScheme of authNameAndSchemes) {
        await this.updateAuthActionInYaml(
          authNameAndScheme.authName,
          authNameAndScheme.authScheme,
          inputs.projectPath,
          destinationApiSpecPath,
          destinationPluginManifestPath
        );
      }
    } else if (isGenerateFromMCP) {
      // MCP action: create ai-plugin.json with MCP runtime config
      const mcpServerUrl = inputs[QuestionNames.MCPForDAServerUrl];
      const toolsFilePath = inputs[QuestionNames.MCPToolsFilePath];
      const mcpWarnings: Warning[] = [];

      // Validate: --mcp-da-server-url is required for MCP actions
      if (!mcpServerUrl) {
        return err(
          new UserError(
            Stage.addPlugin,
            "MissingMCPServerUrl",
            getLocalizedString("core.MCPForDA.missingServerUrl"),
            getLocalizedString("core.MCPForDA.missingServerUrl")
          )
        );
      }

      // Load tools from file if provided and not yet loaded
      const existingTools = inputs[QuestionNames.MCPForDAAvailableTools];
      if (toolsFilePath && (!existingTools || existingTools.length === 0)) {
        try {
          const { readMCPToolsFromFile } = await import("../component/utils/mcpToolFetcher");
          const fileTools = await readMCPToolsFromFile(toolsFilePath);
          inputs[QuestionNames.MCPForDAAvailableTools] = fileTools;
          if (!inputs[QuestionNames.MCPForDAPreFetchTools]) {
            inputs[QuestionNames.MCPForDAPreFetchTools] = fileTools.map((t: any) => t.name);
          }
        } catch {
          mcpWarnings.push({
            type: "mcpToolsFileReadError",
            content: getLocalizedString(
              "core.MCPForDA.toolsFileReadError",
              toolsFilePath,
              mcpAddActionHint
            ),
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
          const { probeMCPServerAuth } = await import("../component/utils/mcpToolFetcher");
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

      // Auto-fetch tools if not yet loaded (CLI flow for no-auth servers)
      const currentTools = inputs[QuestionNames.MCPForDAAvailableTools];
      if ((!currentTools || currentTools.length === 0) && mcpServerUrl) {
        try {
          const { fetchMCPTools } = await import("../component/utils/mcpToolFetcher");
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
            mcpWarnings.push({
              type: "mcpAuthRequired",
              content: getLocalizedString(
                "core.MCPForDA.authRequired",
                mcpServerUrl,
                mcpAddActionHint
              ),
            });
          } else {
            mcpWarnings.push({
              type: "mcpNoToolsFetched",
              content: getLocalizedString(
                "core.MCPForDA.noToolsFetched",
                mcpServerUrl,
                mcpAddActionHint
              ),
            });
          }
        } catch {
          mcpWarnings.push({
            type: "mcpFetchError",
            content: getLocalizedString("core.MCPForDA.fetchError", mcpServerUrl, mcpAddActionHint),
          });
        }
      }

      // If no tools available (auth required or fetch failed), skip creating plugin — just warn
      const mcpTools = inputs[QuestionNames.MCPForDAAvailableTools];
      const mcpToolsSelected = inputs[QuestionNames.MCPForDAPreFetchTools] || [];
      if (!mcpTools || mcpTools.length === 0 || mcpToolsSelected.length === 0) {
        if (mcpWarnings.length > 0) {
          for (const warning of mcpWarnings) {
            context.logProvider.warning(warning.content);
          }
        }
        return ok(undefined);
      }

      destinationPluginManifestPath =
        await copilotGptManifestUtils.getDefaultNextAvailablePluginManifestPath(
          appPackageFolder,
          undefined
        );

      const pluginManifest: any = {
        $schema: "https://developer.microsoft.com/json-schemas/copilot/plugin/v2.4/schema.json",
        schema_version: "v2.4",
        name_for_human: "MCP Action",
        description_for_human: "Action powered by MCP server",
        contact_email: "publisher-email@example.com",
        namespace: "mcpAction",
        functions: [],
        runtimes: [],
      };

      if (mcpTools && mcpTools.length > 0 && mcpToolsSelected.length > 0) {
        const selectedTools = mcpTools.filter((tool: any) => mcpToolsSelected.includes(tool.name));

        pluginManifest.functions = selectedTools.map((tool: any) => ({
          name: tool.name,
          description: tool.description || "",
        }));

        // Write mcp-tools file with full tool definitions (no field filtering)
        const mcpToolsFile = `mcp-tools-${suffix}.json`;
        const mcpToolsOutputPath = path.join(appPackageFolder, mcpToolsFile);
        await fs.writeJSON(mcpToolsOutputPath, { tools: selectedTools }, { spaces: 4 });

        pluginManifest.runtimes = [
          {
            type: "RemoteMCPServer",
            spec: {
              url: mcpServerUrl,
              mcp_tool_description: {
                file: mcpToolsFile,
              },
            },
            run_for_functions: pluginManifest.functions.map((f: any) => f.name),
          },
        ];

        // Auth handling
        const mcpAuth = inputs[QuestionNames.MCPForDAAuth];
        if (mcpAuth === "OAuthPluginVault") {
          const authType = inputs[QuestionNames.MCPForDAAuthType];
          if (!authType) {
            return err(
              new UserError(
                Stage.addPlugin,
                "MissingMCPAuthType",
                getLocalizedString("core.MCPForDA.missingAuthType"),
                getLocalizedString("core.MCPForDA.missingAuthType")
              )
            );
          }

          const registrationId = `MCP_DA_AUTH_ID_${actionId.toUpperCase()}`;
          pluginManifest.runtimes[0].auth = {
            type: "OAuthPluginVault",
            reference_id: `$\{\{${registrationId}\}\}`,
          };

          // Resolve OAuth metadata and inject oauth/register into m365agents.yml
          try {
            let authorizationUrl: string | undefined;
            let tokenUrl: string | undefined;
            let refreshUrl: string | undefined;

            if (authType === "oauth") {
              const { resolveMCPOAuthMetadata } = await import("../component/utils/mcpToolFetcher");
              const metadata = await resolveMCPOAuthMetadata(
                inputs[QuestionNames.MCPForDAAuthMetadataUrl]
              );
              authorizationUrl = metadata.authorizationUrl;
              tokenUrl = metadata.tokenUrl;
              refreshUrl = metadata.refreshUrl;
            }

            const { ActionInjector } = await import("../component/configManager/actionInjector");
            const ymlPath = pathUtils.getYmlFilePath(inputs.projectPath);
            if (ymlPath) {
              await ActionInjector.injectCreateOAuthActionForMCP(
                ymlPath,
                authType,
                actionId,
                registrationId,
                mcpServerUrl,
                authorizationUrl,
                tokenUrl,
                refreshUrl
              );
            }
          } catch (error: any) {
            mcpWarnings.push({
              type: "mcpAuthMetadataError",
              content: getLocalizedString(
                "core.MCPForDA.mcpAuthMetadataMissingError",
                error.message
              ),
            });
          }
        }
      }

      await fs.ensureFile(destinationPluginManifestPath);
      await fs.writeJSON(destinationPluginManifestPath, pluginManifest, { spaces: 4 });

      const addActionRes = await copilotGptManifestUtils.addAction(
        declarativeCopilotManifestPath,
        actionId,
        normalizePath(path.relative(appPackageFolder, destinationPluginManifestPath), true)
      );
      if (addActionRes.isErr()) {
        return err(addActionRes.error);
      }

      // Surface warnings for CLI
      if (mcpWarnings.length > 0) {
        for (const warning of mcpWarnings) {
          context.logProvider.warning(warning.content);
        }
      }
    } else {
      const addPluginRes = await addExistingPlugin(
        declarativeCopilotManifestPath,
        inputs[QuestionNames.PluginManifestFilePath].trim(),
        inputs[QuestionNames.PluginOpenApiSpecFilePath].trim(),
        actionId,
        context,
        Stage.addPlugin
      );

      if (addPluginRes.isErr()) {
        return err(addPluginRes.error);
      }
      destinationPluginManifestPath = addPluginRes.value.destinationPluginManifestPath;
      const warningMessage = outputScaffoldingWarningMessage(addPluginRes.value.warnings);
      context.logProvider.info(warningMessage);
    }

    if (inputs.platform === Platform.VSCode) {
      const successMessage = getLocalizedString("core.addPlugin.success.vsc", actionId);
      const viewPluginManifest = getLocalizedString("core.addPlugin.success.viewPluginManifest");
      void context.userInteraction
        .showMessage("info", successMessage, false, viewPluginManifest)
        .then((userRes) => {
          if (userRes.isOk() && userRes.value === viewPluginManifest) {
            context.telemetryReporter.sendTelemetryEvent(
              TelemetryEvent.ViewPluginManifestAfterAdded
            );
            void TOOLS?.ui?.openFile?.(destinationPluginManifestPath);
          }
        });
    } else {
      const successMessage = getLocalizedString(
        "core.addPlugin.success",
        actionId,
        destinationPluginManifestPath
      );
      void context.userInteraction.showMessage("info", successMessage, false);
    }

    return ok(undefined);
  }
}
