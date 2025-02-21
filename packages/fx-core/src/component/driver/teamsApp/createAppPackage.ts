// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  Colors,
  FxError,
  Result,
  err,
  ok,
  PluginManifestSchema,
  File,
} from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import fs from "fs-extra";
import * as path from "path";
import * as uuid from "uuid";
import { Service } from "typedi";
import { getLocalizedString } from "../../../common/localizeUtils";
import { ErrorContextMW } from "../../../common/globalVars";
import { FileNotFoundError, InvalidActionInputError, JSONSyntaxError } from "../../../error/common";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { WrapDriverContext } from "../util/wrapUtil";
import {
  Constants,
  EmbeddedKnowledgeLocalDirectoryName,
  EmbeddedKnowledgeCapabilityName,
} from "./constants";
import { CreateAppPackageArgs } from "./interfaces/CreateAppPackageArgs";
import { manifestUtils } from "./utils/ManifestUtils";
import { InvalidFileOutsideOfTheDirectotryError } from "../../../error/teamsApp";
import { getResolvedManifest, normalizePath } from "./utils/utils";
import { copilotGptManifestUtils } from "./utils/CopilotGptManifestUtils";
import { ManifestType } from "../../utils/envFunctionUtils";
import { getAbsolutePath } from "../../utils/common";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import { updateVersionForTeamsAppYamlFile } from "../util/utils";

export const actionName = "teamsApp/zipAppPackage";

@Service(actionName)
export class CreateAppPackageDriver implements StepDriver {
  description = getLocalizedString("driver.teamsApp.description.createAppPackageDriver");
  readonly progressTitle = getLocalizedString(
    "plugins.appstudio.createPackage.progressBar.message"
  );

  public async execute(
    args: CreateAppPackageArgs,
    context: DriverContext
  ): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(context, actionName, actionName);
    const res = await this.build(args, wrapContext);
    return {
      result: res,
      summaries: wrapContext.summaries,
    };
  }

  @hooks([
    ErrorContextMW({ source: "Teams", component: "CreateAppPackageDriver" }),
    addStartAndEndTelemetry(actionName, actionName),
  ])
  public async build(
    args: CreateAppPackageArgs,
    context: WrapDriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    const result = this.validateArgs(args);
    if (result.isErr()) {
      return err(result.error);
    }

    let manifestPath = args.manifestPath;
    if (!path.isAbsolute(manifestPath)) {
      manifestPath = path.join(context.projectPath, manifestPath);
    }

    const manifestRes = await manifestUtils.getManifestV3(manifestPath, context);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }
    const manifest = manifestRes.value;
    // Deal with relative path
    // Environment variables should have been replaced by value
    // ./build/appPackage/appPackage.dev.zip instead of ./build/appPackage/appPackage.${{TEAMSFX_ENV}}.zip
    const zipFileName = getAbsolutePath(args.outputZipPath, context.projectPath);
    const zipFileDir = path.dirname(zipFileName);
    await fs.mkdir(zipFileDir, { recursive: true });

    let jsonFileDir;
    let teamsManifestJsonFileName;
    const shouldwriteAllManifest = !!args.outputFolder;
    if (args.outputJsonPath) {
      teamsManifestJsonFileName = getAbsolutePath(args.outputJsonPath, context.projectPath);
      jsonFileDir = path.dirname(teamsManifestJsonFileName);
    } else {
      jsonFileDir = getAbsolutePath(args.outputFolder!, context.projectPath);
      teamsManifestJsonFileName = path.join(
        jsonFileDir,
        `manifest.${process.env.TEAMSFX_ENV!}.json`
      );
    }
    await fs.mkdir(jsonFileDir, { recursive: true });

    const appDirectory = path.dirname(manifestPath);

    const colorFile = path.resolve(appDirectory, manifest.icons.color);
    if (!(await fs.pathExists(colorFile))) {
      const error = new FileNotFoundError(
        actionName,
        colorFile,
        "https://aka.ms/teamsfx-actions/teamsapp-zipAppPackage"
      );
      return err(error);
    }
    const colorFileRelativePath = path.relative(appDirectory, colorFile);
    if (colorFileRelativePath.startsWith("..")) {
      return err(new InvalidFileOutsideOfTheDirectotryError(colorFile));
    }

    const outlineFile = path.resolve(appDirectory, manifest.icons.outline);
    if (!(await fs.pathExists(outlineFile))) {
      const error = new FileNotFoundError(
        actionName,
        outlineFile,
        "https://aka.ms/teamsfx-actions/teamsapp-zipAppPackage"
      );
      return err(error);
    }
    const outlineFileRelativePath = path.relative(appDirectory, outlineFile);
    if (outlineFileRelativePath.startsWith("..")) {
      return err(new InvalidFileOutsideOfTheDirectotryError(outlineFile));
    }

    // pre-check existence
    if (
      manifest.localizationInfo &&
      manifest.localizationInfo.additionalLanguages &&
      manifest.localizationInfo.additionalLanguages.length > 0
    ) {
      for (const language of manifest.localizationInfo.additionalLanguages) {
        const file = language.file;
        const fileName = `${appDirectory}/${file}`;
        if (!(await fs.pathExists(fileName))) {
          return err(
            new FileNotFoundError(
              actionName,
              fileName,
              "https://aka.ms/teamsfx-actions/teamsapp-zipAppPackage"
            )
          );
        }
      }
    }
    if (manifest.localizationInfo && manifest.localizationInfo.defaultLanguageFile) {
      const file = manifest.localizationInfo.defaultLanguageFile;
      const fileName = `${appDirectory}/${file}`;
      if (!(await fs.pathExists(fileName))) {
        return err(
          new FileNotFoundError(
            actionName,
            fileName,
            "https://aka.ms/teamsfx-actions/teamsapp-zipAppPackage"
          )
        );
      }
    }

    const zip = new AdmZip();
    zip.addFile(Constants.MANIFEST_FILE, Buffer.from(JSON.stringify(manifest, null, 4)));

    // outline.png & color.png, relative path
    let dir = path.dirname(manifest.icons.color);
    zip.addLocalFile(colorFile, dir === "." ? "" : dir);
    dir = path.dirname(manifest.icons.outline);
    zip.addLocalFile(outlineFile, dir === "." ? "" : dir);

    // localization file
    if (
      manifest.localizationInfo &&
      manifest.localizationInfo.additionalLanguages &&
      manifest.localizationInfo.additionalLanguages.length > 0
    ) {
      for (const language of manifest.localizationInfo.additionalLanguages) {
        const file = language.file;
        const fileName = path.resolve(appDirectory, file);
        const relativePath = path.relative(appDirectory, fileName);
        if (relativePath.startsWith("..")) {
          return err(new InvalidFileOutsideOfTheDirectotryError(fileName));
        }
        const resolvedLocFileRes = await manifestUtils.resolveLocFile(fileName);
        if (resolvedLocFileRes.isErr()) {
          return err(resolvedLocFileRes.error);
        }
        if (resolvedLocFileRes.value) {
          zip.addFile(relativePath, Buffer.from(resolvedLocFileRes.value));
        }
      }
    }
    if (manifest.localizationInfo && manifest.localizationInfo.defaultLanguageFile) {
      const file = manifest.localizationInfo.defaultLanguageFile;
      const fileName = path.resolve(appDirectory, file);
      const relativePath = path.relative(appDirectory, fileName);
      if (relativePath.startsWith("..")) {
        return err(new InvalidFileOutsideOfTheDirectotryError(fileName));
      }

      const resolvedLocFileRes = await manifestUtils.resolveLocFile(fileName);
      if (resolvedLocFileRes.isErr()) {
        return err(resolvedLocFileRes.error);
      }
      if (resolvedLocFileRes.value) {
        zip.addFile(relativePath, Buffer.from(resolvedLocFileRes.value));
      }
    }

    // API ME, API specification and Adaptive card templates
    if (
      manifest.composeExtensions &&
      manifest.composeExtensions.length > 0 &&
      manifest.composeExtensions[0].composeExtensionType == "apiBased" &&
      manifest.composeExtensions[0].apiSpecificationFile
    ) {
      const apiSpecificationFile = path.resolve(
        appDirectory,
        manifest.composeExtensions[0].apiSpecificationFile
      );
      const checkExistenceRes = await this.validateReferencedFile(
        apiSpecificationFile,
        appDirectory
      );
      if (checkExistenceRes.isErr()) {
        return err(checkExistenceRes.error);
      }

      const addFileWithVariableRes = await this.addFileWithVariable(
        zip,
        manifest.composeExtensions[0].apiSpecificationFile,
        apiSpecificationFile,
        ManifestType.ApiSpec,
        context
      );
      if (addFileWithVariableRes.isErr()) {
        return err(addFileWithVariableRes.error);
      }

      if (manifest.composeExtensions[0].commands.length > 0) {
        for (const command of manifest.composeExtensions[0].commands) {
          if (command.apiResponseRenderingTemplateFile) {
            const adaptiveCardFile = path.resolve(
              appDirectory,
              command.apiResponseRenderingTemplateFile
            );
            const checkExistenceRes = await this.validateReferencedFile(
              adaptiveCardFile,
              appDirectory
            );
            if (checkExistenceRes.isErr()) {
              return err(checkExistenceRes.error);
            }
            const dir = path.dirname(command.apiResponseRenderingTemplateFile);
            this.addFileInZip(zip, dir, adaptiveCardFile);
          }
        }
      }
    }

    const plugins = manifest.copilotExtensions
      ? manifest.copilotExtensions.plugins
      : manifest.copilotAgents?.plugins;
    if (plugins?.length && plugins[0].file) {
      // API plugin
      const addFilesRes = await this.addPlugin(
        zip,
        plugins[0].file,
        appDirectory,
        context,
        !shouldwriteAllManifest ? undefined : jsonFileDir
      );
      if (addFilesRes.isErr()) {
        return err(addFilesRes.error);
      }
    }

    const declarativeCopilots = manifest.copilotExtensions
      ? manifest.copilotExtensions.declarativeCopilots
      : manifest.copilotAgents?.declarativeAgents;
    // Copilot GPT
    if (declarativeCopilots?.length && declarativeCopilots[0].file) {
      const declarativeAgentManifestFile = path.resolve(appDirectory, declarativeCopilots[0].file);
      const checkExistenceRes = await this.validateReferencedFile(
        declarativeAgentManifestFile,
        appDirectory
      );
      if (checkExistenceRes.isErr()) {
        return err(checkExistenceRes.error);
      }

      const addFileWithVariableRes = await this.addFileWithVariable(
        zip,
        declarativeCopilots[0].file,
        declarativeAgentManifestFile,
        ManifestType.DeclarativeCopilotManifest,
        context,
        shouldwriteAllManifest
          ? path.join(jsonFileDir, path.relative(appDirectory, declarativeAgentManifestFile))
          : undefined
      );
      if (addFileWithVariableRes.isErr()) {
        return err(addFileWithVariableRes.error);
      }
      const getCopilotGptRes = await copilotGptManifestUtils.getManifest(
        declarativeAgentManifestFile,
        context
      );
      if (getCopilotGptRes.isOk()) {
        // Add action files
        if (getCopilotGptRes.value.actions) {
          const pluginFiles = getCopilotGptRes.value.actions.map((action) => action.file);

          for (const pluginFile of pluginFiles) {
            const pluginFileAbsolutePath = path.resolve(
              path.dirname(declarativeAgentManifestFile),
              pluginFile
            );

            const pluginFileRelativePath = path.relative(appDirectory, pluginFileAbsolutePath);
            const useForwardSlash = declarativeCopilots[0].file.concat(pluginFile).includes("/");

            const addPluginRes = await this.addPlugin(
              zip,
              normalizePath(pluginFileRelativePath, useForwardSlash),
              appDirectory,
              context,
              !shouldwriteAllManifest ? undefined : jsonFileDir
            );

            if (addPluginRes.isErr()) {
              return err(addPluginRes.error);
            }
          }
        }
        // Add embedded knowledge files
        if (featureFlagManager.getBooleanValue(FeatureFlags.BuilderAPIEnabled)) {
          if (getCopilotGptRes.value.capabilities) {
            const embeddedKnowledgeCapabilities = getCopilotGptRes.value.capabilities.filter(
              (capability) => capability.name === EmbeddedKnowledgeCapabilityName
            );
            if (embeddedKnowledgeCapabilities.length > 0) {
              const fileSet = new Set<string>();
              for (const capability of embeddedKnowledgeCapabilities) {
                for (const file of capability.files as File[]) {
                  if (file.file) {
                    fileSet.add(file.file);
                  }
                }
              }
              const fileArr = Array.from(fileSet);
              if (fileArr.length > 0) {
                for (const file of fileArr) {
                  const knowledgeFileAbsolutePath = path.resolve(appDirectory, file);
                  // check existence
                  const checkExistenceRes = await this.validateReferencedFile(
                    knowledgeFileAbsolutePath,
                    appDirectory
                  );
                  if (checkExistenceRes.isErr()) {
                    return err(checkExistenceRes.error);
                  }
                  const addKnowledgeRes = await this.addFileWithVariable(
                    zip,
                    file,
                    knowledgeFileAbsolutePath,
                    ManifestType.EmbeddedKnowledgeFile,
                    context
                  );
                  if (addKnowledgeRes.isErr()) {
                    return err(addKnowledgeRes.error);
                  }
                }
              }
            }
          }
        }
      } else {
        return err(getCopilotGptRes.error);
      }
    }

    zip.writeZip(zipFileName);

    await this.writeJsonFile(teamsManifestJsonFileName, JSON.stringify(manifest, null, 4));

    const builtSuccess = [
      { content: "(√)Done: ", color: Colors.BRIGHT_GREEN },
      { content: "Teams Package ", color: Colors.BRIGHT_WHITE },
      { content: zipFileName, color: Colors.BRIGHT_MAGENTA },
      { content: " built successfully!", color: Colors.BRIGHT_WHITE },
    ];
    context.logProvider.info(builtSuccess);
    return ok(new Map());
  }

  private static async expandEnvVars(
    filePath: string,
    ctx: WrapDriverContext,
    manifestType: ManifestType
  ): Promise<Result<string, FxError>> {
    const content = await fs.readFile(filePath, "utf8");
    return getResolvedManifest(content, filePath, manifestType, ctx);
  }

  private validateArgs(args: CreateAppPackageArgs): Result<any, FxError> {
    const invalidParams: string[] = [];
    if (!args || !args.manifestPath) {
      invalidParams.push("manifestPath");
    }
    if (!args || (!args.outputJsonPath && !args.outputFolder)) {
      invalidParams.push("outputJsonPath or outputFolder");
    }
    if (!args || !args.outputZipPath) {
      invalidParams.push("outputZipPath");
    }
    if (invalidParams.length > 0) {
      return err(
        new InvalidActionInputError(
          actionName,
          invalidParams,
          "https://aka.ms/teamsfx-actions/teamsapp-zipAppPackage"
        )
      );
    } else {
      return ok(undefined);
    }
  }

  private async validateReferencedFile(
    file: string,
    directory: string
  ): Promise<Result<undefined, FxError>> {
    if (!(await fs.pathExists(file))) {
      return err(
        new FileNotFoundError(
          actionName,
          file,
          "https://aka.ms/teamsfx-actions/teamsapp-zipAppPackage"
        )
      );
    }

    const relativePath = path.relative(directory, file);
    if (relativePath.startsWith("..")) {
      return err(new InvalidFileOutsideOfTheDirectotryError(file));
    }

    return ok(undefined);
  }

  /**
   * Add plugin file and plugin related files to zip.
   * @param zip zip
   * @param pluginRelativePath plugin file path relative to app package folder
   * @param appDirectory app package path containing manifest template.
   * @param context context
   * @param outputDirectory optional. Folder where we should put the resolved manifest in.
   * @returns result of adding plugin file and plugin related files
   */
  private async addPlugin(
    zip: AdmZip,
    pluginRelativePath: string,
    appDirectory: string,
    context: WrapDriverContext,
    outputDirectory?: string
  ): Promise<Result<undefined, FxError>> {
    const pluginFile = path.resolve(appDirectory, pluginRelativePath);
    const checkExistenceRes = await this.validateReferencedFile(pluginFile, appDirectory);
    if (checkExistenceRes.isErr()) {
      return err(checkExistenceRes.error);
    }

    let pluginFileContent;
    try {
      pluginFileContent = (await fs.readJSON(pluginFile)) as PluginManifestSchema;
    } catch (e) {
      return err(new JSONSyntaxError(pluginFile, e, actionName));
    }

    let containExternalAdaptiveCard = false;
    if (pluginFileContent.functions) {
      for (const func of pluginFileContent.functions) {
        if (func.capabilities?.response_semantics?.static_template?.file) {
          const staticTemplateFile = path.resolve(
            path.dirname(pluginFile),
            func.capabilities.response_semantics.static_template.file as string
          );
          const checkExistenceRes = await this.validateReferencedFile(
            staticTemplateFile,
            appDirectory
          );

          if (checkExistenceRes.isErr()) {
            delete func.capabilities.response_semantics.static_template.file;
            context.logProvider.warning(
              getLocalizedString(
                "plugins.appstudio.createPackage.aiPlugin.invalidFilePropertyWarning",
                pluginFile,
                func.name
              )
            );
            continue;
          }

          if (Object.keys(func.capabilities.response_semantics.static_template).length > 1) {
            context.logProvider.warning(
              getLocalizedString(
                "plugins.appstudio.createPackage.aiPlugin.overrideWarning",
                pluginFile,
                func.name
              )
            );
          }

          const staticTemplateFileContent = await fs.readJSON(staticTemplateFile);
          func.capabilities.response_semantics.static_template = staticTemplateFileContent;

          containExternalAdaptiveCard = true;
        }
      }
    }

    let tmpPluginFile = pluginFile;
    let tempFolder: string | undefined;

    if (containExternalAdaptiveCard) {
      await updateVersionForTeamsAppYamlFile(context.projectPath);
      tempFolder = path.join(appDirectory, ".tmp");
      await fs.ensureDir(tempFolder);
      tmpPluginFile = path.join(tempFolder, `tmp-ai-plugin-${uuid.v4().slice(0, 6)}.json`);
      await fs.writeJSON(tmpPluginFile, pluginFileContent, { spaces: 4 });
    }

    const addFileWithVariableRes = await this.addFileWithVariable(
      zip,
      pluginRelativePath,
      tmpPluginFile,
      ManifestType.PluginManifest,
      context,
      !outputDirectory
        ? undefined
        : path.join(outputDirectory, path.relative(appDirectory, pluginFile))
    );

    if (containExternalAdaptiveCard && tmpPluginFile !== pluginFile && tempFolder) {
      await fs.remove(tempFolder);
    }

    if (addFileWithVariableRes.isErr()) {
      return err(addFileWithVariableRes.error);
    }

    const addFilesRes = await this.addPluginRelatedFiles(
      zip,
      pluginRelativePath,
      appDirectory,
      context
    );
    if (addFilesRes.isErr()) {
      return err(addFilesRes.error);
    } else {
      return ok(undefined);
    }
  }

  /**
   * Add plugin related files (OpenAPI spec) to zip.
   * @param zip zip.
   * @param pluginFile plugin file path relative to app package folder.
   * @param appDirectory app package folder.
   * @param context context.
   * @returns results whether add files related to plugin is successful.
   */
  private async addPluginRelatedFiles(
    zip: AdmZip,
    pluginFile: string,
    appDirectory: string,
    context: WrapDriverContext
  ): Promise<Result<undefined, FxError>> {
    const pluginFilePath = path.join(appDirectory, pluginFile);
    const pluginContent = (await fs.readJSON(pluginFilePath)) as PluginManifestSchema;
    const runtimes = pluginContent.runtimes;
    if (runtimes && runtimes.length > 0) {
      for (const runtime of runtimes) {
        if (runtime.type === "OpenApi" && runtime.spec?.url) {
          const specFile = path.resolve(path.dirname(pluginFilePath), runtime.spec.url);
          // add openapi spec
          const checkExistenceRes = await this.validateReferencedFile(specFile, appDirectory);
          if (checkExistenceRes.isErr()) {
            return err(checkExistenceRes.error);
          }

          const entryName = path.relative(appDirectory, specFile);
          const useForwardSlash = pluginFile.concat(runtime.spec.url).includes("/");

          const addFileWithVariableRes = await this.addFileWithVariable(
            zip,
            normalizePath(entryName, useForwardSlash),
            specFile,
            ManifestType.ApiSpec,
            context
          );
          if (addFileWithVariableRes.isErr()) {
            return err(addFileWithVariableRes.error);
          }
        }
      }
    }

    return ok(undefined);
  }

  private async addFileWithVariable(
    zip: AdmZip,
    entryName: string,
    filePath: string,
    manifestType: ManifestType,
    context: WrapDriverContext,
    outputPath?: string // If outputPath exists, we will write down the file after replacing placeholders.
  ): Promise<Result<undefined, FxError>> {
    const expandedEnvVarResult = await CreateAppPackageDriver.expandEnvVars(
      filePath,
      context,
      manifestType
    );
    if (expandedEnvVarResult.isErr()) {
      return err(expandedEnvVarResult.error);
    }
    const content = expandedEnvVarResult.value;

    const attr = await fs.stat(filePath);
    zip.addFile(entryName, Buffer.from(content), "", attr.mode);

    if (outputPath && path.extname(outputPath).toLowerCase() === ".json") {
      await this.writeJsonFile(
        `${outputPath.substring(0, outputPath.length - 5)}.${process.env.TEAMSFX_ENV!}.json`,
        content
      );
    }

    return ok(undefined);
  }

  private addFileInZip(zip: AdmZip, zipPath: string, filePath: string) {
    zip.addLocalFile(filePath, zipPath === "." ? "" : zipPath);
  }

  private async writeJsonFile(jsonFileName: string, content: string) {
    if (await fs.pathExists(jsonFileName)) {
      await fs.chmod(jsonFileName, 0o777);
    }
    await fs.writeFile(jsonFileName, content);
    await fs.chmod(jsonFileName, 0o444);
  }
}
