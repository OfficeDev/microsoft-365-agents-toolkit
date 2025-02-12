// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  FxError,
  Result,
  err,
  ok,
  DeclarativeCopilotManifestSchema,
  ManifestUtil,
  IDeclarativeCopilot,
  Platform,
  Colors,
  DefaultPluginManifestFileName,
} from "@microsoft/teamsfx-api";
import fs from "fs-extra";
import { FileNotFoundError, JSONSyntaxError, WriteFileError } from "../../../../error/common";
import stripBom from "strip-bom";
import { getResolvedManifest } from "./utils";
import { AppStudioResultFactory } from "../results";
import { AppStudioError } from "../errors";
import { getDefaultString, getLocalizedString } from "../../../../common/localizeUtils";
import { DeclarativeCopilotManifestValidationResult } from "../interfaces/ValidationResult";
import path from "path";
import { pluginManifestUtils } from "./PluginManifestUtils";
import { SummaryConstant } from "../../../configManager/constant";
import { EOL } from "os";
import { ManifestType } from "../../../utils/envFunctionUtils";
import { DriverContext } from "../../interface/commonArgs";
import { manifestUtils } from "./ManifestUtils";
import { ProjectType, SpecParser } from "@microsoft/m365-spec-parser";
import { getParserOptions } from "../../../generator/apiSpec/helper";

export class CopilotGptManifestUtils {
  public async readCopilotGptManifestFile(
    path: string
  ): Promise<Result<DeclarativeCopilotManifestSchema, FxError>> {
    if (!(await fs.pathExists(path))) {
      return err(new FileNotFoundError("CopilotGptManifestUtils", path));
    }
    // Be compatible with UTF8-BOM encoding
    // Avoid Unexpected token error at JSON.parse()
    let content = await fs.readFile(path, { encoding: "utf-8" });
    content = stripBom(content);

    try {
      const manifest = JSON.parse(content) as DeclarativeCopilotManifestSchema;
      return ok(manifest);
    } catch (e) {
      return err(new JSONSyntaxError(path, e, "CopilotGptManifestUtils"));
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
  ): Promise<Result<DeclarativeCopilotManifestSchema, FxError>> {
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
    manifest: DeclarativeCopilotManifestSchema,
    path: string
  ): Promise<Result<undefined, FxError>> {
    const content = JSON.stringify(manifest, undefined, 4);
    try {
      await fs.writeFile(path, content);
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

  public async getManifestPath(teamsManifestPath: string): Promise<Result<string, FxError>> {
    const teamsManifestRes = await manifestUtils._readAppManifest(teamsManifestPath);

    if (teamsManifestRes.isErr()) {
      return err(teamsManifestRes.error);
    }
    const filePath = teamsManifestRes.value.copilotExtensions
      ? teamsManifestRes.value.copilotExtensions.declarativeCopilots?.[0].file
      : teamsManifestRes.value.copilotAgents?.declarativeAgents?.[0].file;
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
    gptManifest: DeclarativeCopilotManifestSchema
  ): Promise<void> {
    const actionManifest = await fs.readJson(actionPath);
    let conversationStarters = actionManifest.capabilities?.conversation_starters;

    if (!conversationStarters || conversationStarters.length === 0) {
      const openApiRuntimes = actionManifest.runtimes?.filter(
        (runtime: any) => runtime.type === "OpenApi"
      );

      if (openApiRuntimes) {
        for (const runtime of openApiRuntimes) {
          const specPathRelativePath = runtime.spec.url;
          const specPath = path.resolve(path.dirname(actionPath), specPathRelativePath);

          if (await fs.pathExists(specPath)) {
            const specParser = new SpecParser(
              specPath,
              getParserOptions(ProjectType.Copilot, true)
            );
            const listResult = await specParser.list();
            const operationIds = actionManifest.functions?.map((func: any) => func.name);
            const newStarters = listResult.APIs.filter(
              (item) =>
                item.isValid &&
                operationIds?.includes(item.operationId) &&
                (item.description || item.summary)
            ).map((operation) => {
              return {
                text: operation.summary || operation.description,
              };
            });

            conversationStarters = (conversationStarters || []).concat(newStarters);
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
            (existingStarter) => existingStarter.text === starter.text
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
  ): Promise<Result<DeclarativeCopilotManifestSchema, FxError>> {
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
    platform: Platform,
    pluginPath: string
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
        if (!pluginPath || actionValidationRes.filePath !== pluginPath) {
          // do not output validation result of the Declarative Copilot if same file has been validated when validating plugin manifest.
          const actionValidationMessage = pluginManifestUtils.logValidationErrors(
            actionValidationRes,
            platform
          ) as string;
          if (actionValidationMessage) {
            outputMessage += (!outputMessage ? "" : EOL) + actionValidationMessage;
          }
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
        if (!pluginPath || actionValidationRes.filePath !== pluginPath) {
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
      }

      return outputMessage;
    }
  }

  public async getDefaultNextAvailablePluginManifestPath(
    folder: string,
    pluginManifestFileName = DefaultPluginManifestFileName,
    isKiotaIntegration = false
  ) {
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

  getPluginManifestFileName(
    pluginManifestNamePrefix: string,
    pluginFileNameSuffix: number,
    isKiotaIntegration: boolean
  ) {
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
  ): Promise<Result<any, FxError>> {
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
}

export const copilotGptManifestUtils = new CopilotGptManifestUtils();
