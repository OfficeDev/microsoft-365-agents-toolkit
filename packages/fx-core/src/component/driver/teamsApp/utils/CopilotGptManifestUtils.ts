// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  APIPluginManifestLatest,
  CapabilityName,
  Colors,
  DeclarativeAgentManifest,
  DeclarativeAgentManifestLatest,
  DeclarativeAgentManifestWrapper,
  DefaultPluginManifestFileName,
  err,
  FxError,
  IDeclarativeCopilot,
  ManifestUtil,
  ok,
  Platform,
  PluginManifestWrapper,
  Result,
} from "@microsoft/teamsfx-api";

// Type aliases derived from latest manifest types for capability operations
// We use DeclarativeAgentManifestLatest for types to get the most complete capability definition
type DeclarativeCapability = NonNullable<DeclarativeAgentManifestLatest["capabilities"]>[number];
type ConversationStarter = NonNullable<
  DeclarativeAgentManifestLatest["conversation_starters"]
>[number];
type PluginRuntime = NonNullable<APIPluginManifestLatest["runtimes"]>[number];
type PluginFunction = NonNullable<APIPluginManifestLatest["functions"]>[number];

// Capability-related type aliases
type OneDriveAndSharePointCapability = DeclarativeCapability & { name: "OneDriveAndSharePoint" };
type WebSearchCapability = DeclarativeCapability & { name: "WebSearch" };
type SharePointIDs = NonNullable<DeclarativeCapability["items_by_sharepoint_ids"]>[number];
type Site = NonNullable<DeclarativeCapability["items_by_url"]>[number];
import fs from "fs-extra";
import { EOL } from "os";
import path from "path";
import { Context } from "vm";
import { listAPIInfo } from "../../../../common/daSpecParser";
import { getDefaultString, getLocalizedString } from "../../../../common/localizeUtils";
import {
  FileNotFoundError,
  JSONSyntaxError,
  UserCancelError,
  WriteFileError,
} from "../../../../error/common";
import { SummaryConstant } from "../../../configManager/constant";
import { ManifestType } from "../../../utils/envFunctionUtils";
import { DriverContext } from "../../interface/commonArgs";
import { EmbeddedKnowledgeLocalDirectoryName } from "../constants";
import { AppStudioError } from "../errors";
import { DeclarativeCopilotManifestValidationResult } from "../interfaces/ValidationResult";
import { AppStudioResultFactory } from "../results";
import { manifestUtils } from "./ManifestUtils";
import { pluginManifestUtils } from "./PluginManifestUtils";
import { getResolvedManifest } from "./utils";

export class CopilotGptManifestUtils {
  public async readCopilotGptManifestFile(
    path: string
  ): Promise<Result<DeclarativeAgentManifest, FxError>> {
    if (!(await fs.pathExists(path))) {
      return err(new FileNotFoundError("CopilotGptManifestUtils", path));
    }

    try {
      const wrapper = await DeclarativeAgentManifestWrapper.read(path);
      return ok(wrapper.data);
    } catch (e) {
      return err(new JSONSyntaxError(path, e, "CopilotGptManifestUtils"));
    }
  }

  public async readDeclarativeAgentManifestFile(
    filePath: string
  ): Promise<Result<DeclarativeAgentManifest, FxError>> {
    if (!(await fs.pathExists(filePath))) {
      return err(new FileNotFoundError("CopilotGptManifestUtils", filePath));
    }

    try {
      const wrapper = await DeclarativeAgentManifestWrapper.read(filePath);
      return ok(wrapper.data);
    } catch (e) {
      return err(new JSONSyntaxError(filePath, e, "CopilotGptManifestUtils"));
    }
  }

  public readCopilotGptManifestFileSync(path: string): Result<DeclarativeAgentManifest, FxError> {
    if (!fs.pathExistsSync(path)) {
      return err(new FileNotFoundError("CopilotGptManifestUtils", path));
    }

    try {
      const wrapper = DeclarativeAgentManifestWrapper.readSync(path);
      return ok(wrapper.data);
    } catch (e) {
      return err(new FileNotFoundError("CopilotGptManifestUtils", path));
    }
  }

  public readDeclarativeAgentManifestFileSync(
    filePath: string
  ): Result<DeclarativeAgentManifest, FxError> {
    if (!fs.pathExistsSync(filePath)) {
      return err(new FileNotFoundError("CopilotGptManifestUtils", filePath));
    }

    try {
      const wrapper = DeclarativeAgentManifestWrapper.readSync(filePath);
      return ok(wrapper.data);
    } catch (e) {
      return err(new JSONSyntaxError(filePath, e, "CopilotGptManifestUtils"));
    }
  }

  /**
   * Get Declarative Copilot Manifest with env value filled.
   * @param path path of declaraitve Copilot
   * @returns resolved manifest
   */
  public async getManifest(
    path: string,
    context: DriverContext
  ): Promise<Result<DeclarativeAgentManifest, FxError>> {
    const manifestRes = await this.readCopilotGptManifestFile(path);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }
    // Add environment variable keys to telemetry
    const resolvedManifestRes = await getResolvedManifest(
      JSON.stringify(manifestRes.value),
      path,
      ManifestType.DeclarativeCopilotManifest,
      context
    );

    if (resolvedManifestRes.isErr()) {
      return err(resolvedManifestRes.error);
    }
    const resolvedManifestString = resolvedManifestRes.value;
    return ok(JSON.parse(resolvedManifestString));
  }

  public async writeCopilotGptManifestFile(
    manifest: DeclarativeAgentManifest,
    filePath: string
  ): Promise<Result<undefined, FxError>> {
    try {
      const wrapper = DeclarativeAgentManifestWrapper.fromJSON(JSON.stringify(manifest));
      await wrapper.save(filePath);
    } catch (e) {
      return err(new WriteFileError(e, "copilotGptManifestUtils"));
    }
    return ok(undefined);
  }

  public async writeDeclarativeAgentManifestFile(
    manifest: DeclarativeAgentManifest,
    filePath: string
  ): Promise<Result<undefined, FxError>> {
    try {
      const wrapper = DeclarativeAgentManifestWrapper.fromJSON(JSON.stringify(manifest));
      await wrapper.save(filePath);
    } catch (e) {
      return err(new WriteFileError(e, "copilotGptManifestUtils"));
    }
    return ok(undefined);
  }

  public async validateAgainstSchema(
    declaraitveCopilot: IDeclarativeCopilot,
    manifestPath: string,
    context: DriverContext
  ): Promise<Result<DeclarativeCopilotManifestValidationResult, FxError>> {
    const manifestRes = await this.getManifest(manifestPath, context);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }

    const manifest = manifestRes.value;
    try {
      const manifestValidationRes = await ManifestUtil.validateManifest(manifestRes.value);
      const res: DeclarativeCopilotManifestValidationResult = {
        id: declaraitveCopilot.id,
        filePath: manifestPath,
        validationResult: manifestValidationRes,
        actionValidationResult: [],
      };

      if (manifest.actions?.length) {
        // action
        for (const action of manifest.actions) {
          const actionPath = path.join(path.dirname(manifestPath), action.file);

          const actionValidationRes = await pluginManifestUtils.validateAgainstSchema(
            action,
            actionPath,
            context
          );
          if (actionValidationRes.isErr()) {
            return err(actionValidationRes.error);
          } else {
            res.actionValidationResult.push(actionValidationRes.value);
          }
        }
      }
      return ok(res);
    } catch (e: unknown) {
      const errorMessage = e instanceof Error ? e.message : String(e);
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.ValidationFailedError.name,
          AppStudioError.ValidationFailedError.message([
            getLocalizedString(
              "error.appstudio.validateFetchSchemaFailed",
              manifestRes.value.$schema,
              errorMessage
            ),
          ])
        )
      );
    }
  }

  public async getManifestPath(teamsManifestPath: string): Promise<Result<string, FxError>> {
    const teamsManifestRes = await manifestUtils._readAppManifest(teamsManifestPath);

    if (teamsManifestRes.isErr()) {
      return err(teamsManifestRes.error);
    }
    const filePath = teamsManifestRes.value.copilotExtensions
      ? teamsManifestRes.value.copilotExtensions.declarativeCopilots?.[0]?.file
      : teamsManifestRes.value.copilotAgents?.declarativeAgents?.[0]?.file;
    if (!filePath) {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.TeamsAppRequiredPropertyMissingError.name,
          AppStudioError.TeamsAppRequiredPropertyMissingError.message(
            "copilotExtensions.declarativeCopilots.file",
            teamsManifestPath
          )
        )
      );
    } else {
      return ok(path.resolve(path.dirname(teamsManifestPath), filePath));
    }
  }

  public async updateConversationStarters(
    actionPath: string,
    gptManifest: DeclarativeAgentManifest
  ): Promise<void> {
    const actionWrapper = await PluginManifestWrapper.read(actionPath);
    let conversationStarters = actionWrapper.data.capabilities?.conversation_starters;

    if (!conversationStarters || conversationStarters.length === 0) {
      const openApiRuntimes = actionWrapper.runtimes.filter(
        (runtime: PluginRuntime) => runtime.type === "OpenApi"
      );

      if (openApiRuntimes.length > 0) {
        for (const runtime of openApiRuntimes) {
          if (runtime.type === "OpenApi" && runtime.spec.url) {
            const specPathRelativePath = runtime.spec.url;
            const specPath = path.resolve(path.dirname(actionPath), specPathRelativePath);

            if (await fs.pathExists(specPath)) {
              const listResult = await listAPIInfo(specPath);
              const operationIds = actionWrapper.functions.map((func: PluginFunction) => func.name);
              const newStarters = listResult.APIs.filter(
                (item) =>
                  item.isValid &&
                  operationIds?.includes(item.operationId) &&
                  (item.description || item.summary)
              )
                .map((operation) => ({
                  text: operation.summary ?? operation.description ?? "",
                }))
                .filter((starter) => starter.text !== "");

              conversationStarters = (conversationStarters || []).concat(newStarters);
            }
          }
        }
      }
    }

    if (conversationStarters) {
      if (!gptManifest.conversation_starters) {
        gptManifest.conversation_starters = [];
      }

      for (const starter of conversationStarters) {
        if (gptManifest.conversation_starters.length >= 6) {
          break;
        }
        if (
          !gptManifest.conversation_starters.some(
            (existingStarter: ConversationStarter) => existingStarter.text === starter.text
          )
        ) {
          gptManifest.conversation_starters.push(starter);
        }
      }
    }
  }

  public async addAction(
    copilotGptPath: string,
    id: string,
    pluginFile: string
  ): Promise<Result<DeclarativeAgentManifest, FxError>> {
    const gptManifestRes = await copilotGptManifestUtils.readCopilotGptManifestFile(copilotGptPath);
    if (gptManifestRes.isErr()) {
      return err(gptManifestRes.error);
    } else {
      const gptManifest = gptManifestRes.value;
      if (!gptManifest.actions) {
        gptManifest.actions = [];
      }
      gptManifest.actions?.push({
        id,
        file: pluginFile,
      });

      const actionPath = path.join(path.dirname(copilotGptPath), pluginFile);
      await this.updateConversationStarters(actionPath, gptManifest);

      const updateGptManifestRes = await copilotGptManifestUtils.writeCopilotGptManifestFile(
        gptManifest,
        copilotGptPath
      );
      if (updateGptManifestRes.isErr()) {
        return err(updateGptManifestRes.error);
      } else {
        return ok(gptManifest);
      }
    }
  }

  public logValidationErrors(
    validationRes: DeclarativeCopilotManifestValidationResult,
    platform: Platform
  ): string | Array<{ content: string; color: Colors }> {
    const validationErrors = validationRes.validationResult;
    const filePath = validationRes.filePath;
    const hasDeclarativeCopilotError = validationErrors.length > 0;
    let hasActionError = false;

    for (const actionValidationRes of validationRes.actionValidationResult) {
      if (actionValidationRes.validationResult.length > 0) {
        hasActionError = true;
        break;
      }
    }
    if (!hasDeclarativeCopilotError && !hasActionError) {
      return "";
    }

    if (platform !== Platform.CLI) {
      let outputMessage = "";
      if (hasDeclarativeCopilotError) {
        const errors = validationErrors
          .map((error: string) => {
            return `${SummaryConstant.Failed} ${error}`;
          })
          .join(EOL);
        outputMessage +=
          getLocalizedString(
            "driver.teamsApp.summary.validateDeclarativeCopilotManifest.checkPath",
            filePath
          ) +
          EOL +
          errors;
      }

      for (const actionValidationRes of validationRes.actionValidationResult) {
        const actionValidationMessage = pluginManifestUtils.logValidationErrors(
          actionValidationRes,
          platform
        ) as string;
        if (actionValidationMessage) {
          outputMessage += (!outputMessage ? "" : EOL) + actionValidationMessage;
        }
      }

      return outputMessage;
    } else {
      const outputMessage = [];
      if (hasDeclarativeCopilotError) {
        outputMessage.push({
          content:
            getDefaultString(
              "driver.teamsApp.summary.validateDeclarativeCopilotManifest.checkPath",
              filePath
            ) + "\n",
          color: Colors.BRIGHT_WHITE,
        });
        validationErrors.map((error: string) => {
          outputMessage.push({ content: `${SummaryConstant.Failed} `, color: Colors.BRIGHT_RED });
          outputMessage.push({
            content: `${error}\n`,
            color: Colors.BRIGHT_WHITE,
          });
        });
      }

      for (const actionValidationRes of validationRes.actionValidationResult) {
        const actionValidationMessage = pluginManifestUtils.logValidationErrors(
          actionValidationRes,
          platform
        );
        if (actionValidationMessage) {
          outputMessage.push(
            ...(actionValidationMessage as Array<{ content: string; color: Colors }>)
          );
        }
      }

      return outputMessage;
    }
  }

  public async getDefaultNextAvailablePluginManifestPath(
    folder: string,
    pluginManifestFileName = DefaultPluginManifestFileName,
    isKiotaIntegration = false
  ): Promise<string> {
    if (!(await fs.pathExists(path.join(folder, pluginManifestFileName)))) {
      return path.join(folder, pluginManifestFileName);
    }
    const pluginManifestNamePrefix = pluginManifestFileName.split(".")[0];
    let pluginFileNameSuffix = 1;
    let pluginManifestName = this.getPluginManifestFileName(
      pluginManifestNamePrefix,
      pluginFileNameSuffix,
      isKiotaIntegration
    );
    while (await fs.pathExists(path.join(folder, pluginManifestName))) {
      pluginFileNameSuffix++;
      pluginManifestName = this.getPluginManifestFileName(
        pluginManifestNamePrefix,
        pluginFileNameSuffix,
        isKiotaIntegration
      );
    }
    return path.join(folder, pluginManifestName);
  }

  public async addEmbeddedKnowledgeFiles(
    manifestFilePath: string,
    filePathList: string[]
  ): Promise<Result<undefined, FxError>> {
    const declarativeAgentManifestPathRes = await copilotGptManifestUtils.getManifestPath(
      manifestFilePath
    );
    if (declarativeAgentManifestPathRes.isErr()) {
      return err(declarativeAgentManifestPathRes.error);
    }

    const declarativeAgentManifestPath = declarativeAgentManifestPathRes.value;
    const declarativeAgentManifesRes = await copilotGptManifestUtils.readCopilotGptManifestFile(
      declarativeAgentManifestPath
    );
    if (declarativeAgentManifesRes.isErr()) {
      return err(declarativeAgentManifesRes.error);
    }

    const declarativeAgentManifest = declarativeAgentManifesRes.value;
    if (!declarativeAgentManifest.capabilities) {
      declarativeAgentManifest.capabilities = [];
    }
    let embeddedKnowledgeCapability: DeclarativeCapability | undefined = (
      declarativeAgentManifest.capabilities as DeclarativeCapability[]
    ).find(
      (capability: DeclarativeCapability) => capability.name === CapabilityName.EmbeddedKnowledge
    );
    if (!embeddedKnowledgeCapability) {
      embeddedKnowledgeCapability = {
        name: CapabilityName.EmbeddedKnowledge,
        files: [],
      } as DeclarativeCapability;
      (declarativeAgentManifest.capabilities as DeclarativeCapability[]).push(
        embeddedKnowledgeCapability
      );
    }
    await fs.ensureDir(
      path.resolve(path.dirname(manifestFilePath), EmbeddedKnowledgeLocalDirectoryName)
    );
    // Ensure files array exists
    if (!embeddedKnowledgeCapability.files) {
      embeddedKnowledgeCapability.files = [];
    }
    for (const filePath of filePathList) {
      const savedAbsolutePath = path.resolve(
        path.dirname(manifestFilePath),
        EmbeddedKnowledgeLocalDirectoryName,
        path.basename(filePath)
      );
      await fs.copyFile(filePath, savedAbsolutePath);
      embeddedKnowledgeCapability.files.push({
        file: path.relative(path.dirname(manifestFilePath), savedAbsolutePath).replace(/\\/g, "/"),
      });
    }
    // save the updated declarativeCopilotManifestPath
    await copilotGptManifestUtils.writeCopilotGptManifestFile(
      declarativeAgentManifest,
      declarativeAgentManifestPath
    );
    return ok(undefined);
  }

  getPluginManifestFileName(
    pluginManifestNamePrefix: string,
    pluginFileNameSuffix: number,
    isKiotaIntegration: boolean
  ): string {
    let pluginManifestName;
    if (isKiotaIntegration) {
      const pluginManifestNameSplit = pluginManifestNamePrefix.split("-");
      pluginManifestName = `${pluginManifestNameSplit[0]}_${pluginFileNameSuffix}-${pluginManifestNameSplit[1]}.json`;
    } else {
      pluginManifestName = `${pluginManifestNamePrefix}_${pluginFileNameSuffix}.json`;
    }
    return pluginManifestName;
  }

  async updateDeclarativeAgentManifest(
    manifestPath: string,
    declarativeAgentManifestPath: string,
    declarativeCopilotActionId: string,
    pluginManifestPath: string
  ): Promise<Result<undefined, FxError>> {
    const gptManifestPath = path.join(path.dirname(manifestPath), declarativeAgentManifestPath);
    const addAcionResult = await copilotGptManifestUtils.addAction(
      gptManifestPath,
      declarativeCopilotActionId,
      path.basename(pluginManifestPath)
    );
    if (addAcionResult.isErr()) {
      return err(addAcionResult.error);
    }

    return ok(undefined);
  }

  public async addOneDriveSharePointCapability(
    agentManifestPath: string,
    items_by_sharepoint_ids: SharePointIDs | null,
    items_by_url: Site | null,
    manifestRes: Result<DeclarativeAgentManifest, FxError>
  ): Promise<Result<DeclarativeAgentManifest, FxError>> {
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }

    const agentManifest = manifestRes.value;
    if (!agentManifest.capabilities) {
      agentManifest.capabilities = [];
    }

    const newCapabilityData: Partial<OneDriveAndSharePointCapability> = {
      name: CapabilityName.OneDriveAndSharePoint,
    };

    const capability = (agentManifest.capabilities as DeclarativeCapability[]).find(
      (cap: DeclarativeCapability) => cap.name === CapabilityName.OneDriveAndSharePoint
    ) as OneDriveAndSharePointCapability | undefined;

    if (items_by_url) {
      newCapabilityData.items_by_url = capability ? capability.items_by_url || [] : [];
      newCapabilityData.items_by_url.push(items_by_url);
    }
    if (items_by_sharepoint_ids) {
      newCapabilityData.items_by_sharepoint_ids = capability
        ? capability.items_by_sharepoint_ids || []
        : [];
      newCapabilityData.items_by_sharepoint_ids.push(items_by_sharepoint_ids);
    }

    return this.addOrUpdateCapability(
      agentManifestPath,
      CapabilityName.OneDriveAndSharePoint,
      manifestRes,
      newCapabilityData
    );
  }

  public async addWebSearchCapability(
    context: Context,
    agentManifestPath: string,
    items_by_url: Site | null,
    manifestRes: Result<DeclarativeAgentManifest, FxError>
  ): Promise<Result<DeclarativeAgentManifest, FxError>> {
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }

    const agentManifest = manifestRes.value;
    if (!agentManifest.capabilities) {
      agentManifest.capabilities = [];
    }

    const newCapabilityData: Partial<WebSearchCapability> = {
      name: CapabilityName.WebSearch,
    };
    const capability = (agentManifest.capabilities as DeclarativeCapability[]).find(
      (cap: DeclarativeCapability) => cap.name === CapabilityName.WebSearch
    ) as WebSearchCapability | undefined;

    // del capability warning
    if (items_by_url === null && capability?.sites?.length) {
      const confirmRes = await context.userInteraction.showMessage(
        "warn",
        getLocalizedString("core.addKnowledge.doubleConfirm"),
        true,
        getLocalizedString("core.addKnowledge.continue")
      );

      if (confirmRes.isErr()) {
        return err(confirmRes.error);
      } else if (confirmRes.value !== getLocalizedString("core.addKnowledge.continue")) {
        return err(new UserCancelError());
      }
    }

    if (items_by_url) {
      newCapabilityData.sites = capability ? capability.sites || [] : [];
      if (items_by_url.url) {
        newCapabilityData.sites.push({ url: items_by_url.url });
      }
    }

    return this.addOrUpdateCapability(
      agentManifestPath,
      CapabilityName.WebSearch,
      manifestRes,
      newCapabilityData
    );
  }

  public async addGCCapability(
    agentManifestPath: string,
    inputConnectionIds: string[],
    manifestRes: Result<DeclarativeAgentManifest, FxError>
  ): Promise<Result<DeclarativeAgentManifest, FxError>> {
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }

    const agentManifest = manifestRes.value;
    if (!agentManifest.capabilities) {
      agentManifest.capabilities = [];
    }

    let capability: DeclarativeCapability | undefined = (
      agentManifest.capabilities as DeclarativeCapability[]
    ).find((cap: DeclarativeCapability) => cap.name === CapabilityName.GraphConnectors);
    if (!capability) {
      capability = {
        name: CapabilityName.GraphConnectors,
        connections: [],
      } as DeclarativeCapability;
      (agentManifest.capabilities as DeclarativeCapability[]).push(capability);
    }
    if (!capability.connections) {
      capability.connections = [];
    }
    const connections = capability.connections;
    inputConnectionIds.forEach((id) => {
      if (
        !connections.some(
          (connection: { connection_id: string }) => connection.connection_id === id
        )
      ) {
        connections.push({ connection_id: id });
      }
    });

    return this.addOrUpdateCapability(
      agentManifestPath,
      CapabilityName.GraphConnectors,
      manifestRes,
      {
        connections,
      }
    );
  }

  /**
   * Updates or adds a capability in the agent manifest.
   *
   * @param agentManifestPath - The path to the agent manifest file.
   * @param capabilityName - The name of the capability to update or add.
   * @param manifestRes - The result containing the agent manifest schema.
   * @param capabilityData - The data for the capability to update or add.
   */
  async addOrUpdateCapability(
    agentManifestPath: string,
    capabilityName: DeclarativeCapability["name"],
    manifestRes: Result<DeclarativeAgentManifest, FxError>,
    capabilityData: Partial<DeclarativeCapability>
  ): Promise<Result<DeclarativeAgentManifest, FxError>> {
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }

    const agentManifest = manifestRes.value;
    if (!agentManifest.capabilities) {
      agentManifest.capabilities = [];
    }

    const newCapability: DeclarativeCapability = {
      name: capabilityName,
      ...capabilityData,
    } as DeclarativeCapability;

    const capabilities = agentManifest.capabilities as DeclarativeCapability[];
    const capabilityIndex = capabilities.findIndex(
      (cap: DeclarativeCapability) => cap.name === capabilityName
    );
    if (capabilityIndex !== -1) {
      capabilities[capabilityIndex] = newCapability;
    } else {
      capabilities.push(newCapability);
    }

    const updateGptManifestRes = await this.writeCopilotGptManifestFile(
      agentManifest,
      agentManifestPath
    );
    if (updateGptManifestRes.isErr()) {
      return err(updateGptManifestRes.error);
    } else {
      return ok(agentManifest);
    }
  }
}

export const copilotGptManifestUtils = new CopilotGptManifestUtils();
