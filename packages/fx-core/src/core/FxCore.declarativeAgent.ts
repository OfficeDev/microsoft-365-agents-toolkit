// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks";
import {
  FxError,
  Inputs,
  Result,
  Stage,
  SystemError,
  UserError,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import axios from "axios";
import fs from "fs-extra";
import * as path from "path";
import { ErrorContextMW, createContext } from "../common/globalVars";
import { getDefaultString, getLocalizedString } from "../common/localizeUtils";
import { ActionInjector } from "../component/configManager/actionInjector";
import { LocalMcpPrefix } from "../component/constants";
import { QuestionMW } from "../component/middleware/questionMW";
import { pathUtils } from "../component/utils/pathUtils";
import { QuestionNames } from "../question/constants";
import { ConcurrentLockerMW } from "./middleware/concurrentLocker";
import { ErrorHandlerMW } from "./middleware/errorHandler";

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
          let wellKnownMetadataUrl = inputs[QuestionNames.MCPForDAAuthWellKnownUrl];
          if (!wellKnownMetadataUrl) {
            const mcpAuthMetadataUrl = inputs[QuestionNames.MCPForDAAuthMetadataUrl];
            if (!mcpAuthMetadataUrl) {
              throw new Error(getLocalizedString("core.MCPForDA.mcpAuthMetadataUrlNotFound"));
            }

            const response = await axios.get(mcpAuthMetadataUrl);
            if (
              response.status === 200 &&
              response.data &&
              response.data.authorization_servers &&
              response.data.authorization_servers.length > 0
            ) {
              const mcpServerMetadataUrl = response.data.authorization_servers?.[0];
              // Transform the URL to the proper OAuth authorization server metadata endpoint
              // According to RFC 8414, the well-known endpoint should be constructed as:
              // https://{domain}/.well-known/oauth-authorization-server{path}
              const serverUrl = new URL(mcpServerMetadataUrl);
              wellKnownMetadataUrl = `${serverUrl.protocol}//${serverUrl.host}/.well-known/oauth-authorization-server${serverUrl.pathname}`;
            } else {
              throw new Error(getLocalizedString("core.MCPForDA.mcpServerMetadataUrlNotFound"));
            }
          }
          const metadataResponse = await axios.get(wellKnownMetadataUrl);
          if (metadataResponse.status === 200) {
            oauthAuthorizationUrl = metadataResponse.data.authorization_endpoint;
            oauthTokenUrl = metadataResponse.data.token_endpoint;
            oauthRefreshUrl = metadataResponse.data.refresh_endpoint;
          }
          if (!oauthAuthorizationUrl || !oauthTokenUrl) {
            throw new Error(getLocalizedString("core.MCPForDA.authUrlNotFound"));
          }
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
      // let mcpFile = matchedRuntime?.spec.mcp_tool_description?.file;
      // if (!mcpFile) {
      //   mcpFile = "mcp-tools.json";
      //   let suffix = 1;
      //   while (await fs.pathExists(path.join(path.dirname(aiPluginFilePath), mcpFile))) {
      //     mcpFile = `mcp-tools-${suffix}.json`;
      //     suffix += 1;
      //   }
      // }
      // await fs.writeJSON(
      //   path.join(path.dirname(aiPluginFilePath), mcpFile),
      //   {
      //     tools: [
      //       ...mcpToolsDetail
      //         .filter((tool: any) => mcpToolsSelected.includes(tool.name))
      //         .map((tool: any) => {
      //           return {
      //             ...tool,
      //             title: (tool.name as string)
      //               .replace(/_/g, " ")
      //               .replace(/^./, (str) => str.toUpperCase()),
      //           };
      //         }),
      //     ],
      //   },
      //   { spaces: 4 }
      // );
      (aiPluginContent.runtimes as any[]).push({
        type: "RemoteMCPServer",
        spec: {
          url: mcpServerUrl,
          mcp_tool_description: {
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
}
