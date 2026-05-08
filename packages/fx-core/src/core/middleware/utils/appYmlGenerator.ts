// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AppPackageFolderName } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import * as handlebars from "handlebars";
import * as path from "path";
import { MetadataV3 } from "../../../common/versionMetadata";
import {
  ComponentNames,
  convertProjectSettingsV2ToV3,
  getComponent,
} from "../../../component/migrate";
import { getTemplatesFolder } from "../../../folder";
import { FileType, namingConverterV3 } from "./MigrationUtils";
import { DebugPlaceholderMapping } from "./debug/debugV3MigrationUtils";
export abstract class BaseAppYmlGenerator {
  protected abstract handlebarsContext: any;
  constructor(protected oldProjectSettings: any) {}

  protected async buildHandlebarsTemplate(templateName: string): Promise<string> {
    const templatePath = path.join(getTemplatesFolder(), "core/v3Migration", templateName);
    const templateString = await fs.readFile(templatePath, "utf8");
    const template = handlebars.compile(templateString);
    return template(this.handlebarsContext);
  }
}
export function hasFunctionBot(projectSettings: any): boolean {
  const botComponent = getComponent(projectSettings, ComponentNames.TeamsBot);
  if (!botComponent) return false;
  return botComponent.hosting === ComponentNames.Function;
}
export class AppYmlGenerator extends BaseAppYmlGenerator {
  protected handlebarsContext: {
    activePlugins: Record<string, boolean>;
    placeholderMappings: Record<string, string>;
    aadAppName: string | undefined;
    teamsAppName: string | undefined;
    appName: string | undefined;
    isFunctionBot: boolean;
    botResourceId: string | undefined;
    isTypescript: boolean;
    defaultFunctionName: string | undefined;
    environmentFolder: string | undefined;
    projectId: string | undefined;
    dotnetPath: string | undefined;
    isM365: boolean | undefined;
  };
  constructor(
    oldProjectSettings: any,
    private bicepContent: string,
    private projectPath: string
  ) {
    super(oldProjectSettings);
    this.handlebarsContext = {
      activePlugins: {},
      placeholderMappings: {},
      aadAppName: undefined,
      teamsAppName: undefined,
      appName: undefined,
      isFunctionBot: false,
      botResourceId: undefined,
      isTypescript: false,
      defaultFunctionName: undefined,
      environmentFolder: undefined,
      projectId: undefined,
      dotnetPath: "DOTNET_PATH",
      isM365: undefined,
    };
  }

  public async generateAppYml(): Promise<string> {
    await this.generateCommonHandlerbarsContext();

    const solutionSettings = this.oldProjectSettings.solutionSettings;
    if (solutionSettings.hostType.toLowerCase() === "azure") {
      await this.generateAzureHandlebarsContext();
      switch (this.oldProjectSettings.programmingLanguage?.toLowerCase()) {
        case "javascript":
        case "typescript":
          return await this.buildHandlebarsTemplate("js.ts.app.yml");
        case "csharp":
          return await this.buildHandlebarsTemplate("csharp.app.yml");
      }
    } else if (solutionSettings.hostType.toLowerCase() === "spfx") {
      return await this.buildHandlebarsTemplate("spfx.app.yml");
    }
    throw new Error(
      "The current tooling cannot upgrade your project temporary. Please raise an issue in GitHub for your project."
    );
  }

  public async generateAppLocalYml(placeholderMappings: DebugPlaceholderMapping): Promise<string> {
    this.handlebarsContext.placeholderMappings = placeholderMappings as any;
    await this.generateAzureHandlebarsContext();

    const solutionSettings = this.oldProjectSettings.solutionSettings;
    if (solutionSettings.hostType === "Azure") {
      switch (this.oldProjectSettings.programmingLanguage?.toLowerCase()) {
        case "csharp":
          return await this.buildHandlebarsTemplate("csharp.app.local.yml");
      }
    }
    throw new Error(
      "The current tooling cannot upgrade your project temporary. Please raise an issue in GitHub for your project."
    );
  }

  private async generateCommonHandlerbarsContext(): Promise<void> {
    // project setting information
    this.handlebarsContext.appName = this.oldProjectSettings.appName;

    const azureSolutionSettings = this.oldProjectSettings.solutionSettings;
    for (const activePlugin of azureSolutionSettings.activeResourcePlugins) {
      this.handlebarsContext.activePlugins[activePlugin] = true; // convert array items to object properties to simplify handlebars template
    }

    // app names
    const aadManifestPath = path.join(this.projectPath, MetadataV3.aadManifestFileName);
    if (await fs.pathExists(aadManifestPath)) {
      const aadManifest = await fs.readJson(
        path.join(this.projectPath, MetadataV3.aadManifestFileName)
      );
      this.handlebarsContext.aadAppName = aadManifest.name;
    }

    const teamsAppManifestPath = path.join(
      this.projectPath,
      AppPackageFolderName,
      MetadataV3.teamsManifestFileName
    );
    if (await fs.pathExists(teamsAppManifestPath)) {
      const teamsAppManifest = await fs.readJson(
        path.join(this.projectPath, AppPackageFolderName, MetadataV3.teamsManifestFileName)
      );
      this.handlebarsContext.teamsAppName = teamsAppManifest.name.short;
    }

    // programming language
    this.handlebarsContext.isTypescript =
      this.oldProjectSettings.programmingLanguage?.toLowerCase() === "typescript";

    // default function name
    this.handlebarsContext.defaultFunctionName = this.oldProjectSettings.defaultFunctionName;

    // projectId
    this.handlebarsContext.projectId = this.oldProjectSettings.projectId;

    // env folder
    this.handlebarsContext.environmentFolder = MetadataV3.defaultEnvironmentFolder;

    this.handlebarsContext.isM365 = this.oldProjectSettings.isM365;
  }

  private generateAzureHandlebarsContext(): Promise<void> {
    // isFunctionBot
    const projectSettings: any = convertProjectSettingsV2ToV3(
      this.oldProjectSettings,
      this.projectPath
    );
    // placeholders
    this.setPlaceholderMapping("state.fx-resource-frontend-hosting.storageResourceId");
    this.setPlaceholderMapping("state.fx-resource-frontend-hosting.endpoint");
    this.setPlaceholderMapping("state.fx-resource-frontend-hosting.resourceId");
    this.setPlaceholderMapping("state.fx-resource-frontend-hosting.indexPath");
    this.setPlaceholderMapping("state.fx-resource-bot.resourceId");
    this.setPlaceholderMapping("state.fx-resource-bot.functionAppResourceId");
    this.setPlaceholderMapping("state.fx-resource-bot.webAppResourceId");
    this.setPlaceholderMapping("state.fx-resource-bot.botWebAppResourceId");
    this.setPlaceholderMapping("state.fx-resource-function.functionAppResourceId");
    this.setPlaceholderMapping("state.fx-resource-function.functionEndpoint");

    this.handlebarsContext.isFunctionBot = hasFunctionBot(projectSettings); // if not function bot but a resource bot, then a webApp bot

    // Match teams-bot or fx-resource-bot output obj
    const pluginRegex = new RegExp(
      "output +(\\S+) +object += +{" + // Mataches start of output declaration and capture output name. Example: output functionOutput object = {
        "[^{]*" + // Matches everything between '{' and plugin id declaration. For example: comments, extra properties. Will match multilines.
        "teamsFxPluginId: +'(teams-bot|fx-resource-bot)'" + // Matches given plugin id == teams-bot or fx-resource-bot
        "[^}]*" + // Mathches anything except '}'
        "(botWebAppResourceId|webAppResourceId|functionAppResourceId|resourceId) *:" + // Matches resource id and tries not to mismatch key and value
        "[^}]*}", // Matches until end of obj as '}'
      "g"
    );
    const outputContents = pluginRegex.exec(this.bicepContent);
    if (outputContents) {
      const prefix = "state.fx-resource-bot.";
      this.handlebarsContext.botResourceId =
        this.handlebarsContext.placeholderMappings[`${prefix}${outputContents[3]}`];
    }
    return Promise.resolve();
  }

  private setPlaceholderMapping(placeholder: string): void {
    const result = namingConverterV3(placeholder, FileType.STATE, this.bicepContent);
    if (result.isOk()) {
      this.handlebarsContext.placeholderMappings[placeholder] = result.value;
    }
    // ignore non-exist placeholder
  }
}
