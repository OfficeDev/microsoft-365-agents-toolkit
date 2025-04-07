// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Service } from "typedi";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { getLocalizedString } from "../../../common/localizeUtils";
import { DriverContext } from "../interface/commonArgs";
import { TypeSpecCompileArgs } from "./interface/typeSpecCompileArgs";
import { hooks } from "@feathersjs/hooks";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import {
  DeclarativeCopilotManifestSchema,
  err,
  ok,
  Platform,
  SystemError,
  TeamsAppManifest,
  UserError,
} from "@microsoft/teamsfx-api";
import path from "path";
import fs from "fs-extra";
import {
  assembleError,
  InputValidationError,
  InvalidActionInputError,
} from "../../../error/common";
import { defaultDAManifestFileName, defaultOpenApiOutputDir, helpLink } from "./constants";
import { NoSpecError } from "./error/noSpecError";
import { MultipleActionError } from "./error/multipleActionError";
import {
  injectAuthAction,
  parseAndUpdatePluginManifestForKiota,
} from "../../generator/openApiSpec/helper";
import { MetadataV3 } from "../../../common/versionMetadata";
import { ReProvisionError } from "./error/reProvisionError";

const actionName = "typeSpec/compile"; // DO NOT MODIFY the name

@Service(actionName) // DO NOT MODIFY the service name
export class TypeSpecCompileDriver implements StepDriver {
  description = getLocalizedString("driver.typeSpec.compile.description");
  readonly progressTitle = getLocalizedString("driver.typeSpec.progressBar");

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async execute(
    args: TypeSpecCompileArgs,
    ctx: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const summaries: string[] = [];
    const outputs: Map<string, string> = new Map<string, string>();

    try {
      ctx.logProvider?.info(getLocalizedString("driver.typeSpec.compile.start", actionName));
      this.validateArgs(args);

      const projectPath = ctx.projectPath;
      if (!projectPath) {
        throw new InputValidationError("projectPath", "undefined");
      }

      const mainFilePath = args.path;
      const outputFolderPath = path.isAbsolute(args.outputDir)
        ? args.outputDir
        : path.join(ctx.projectPath, args.outputDir);
      const openApiSpecsFolderPath = path.join(outputFolderPath, defaultOpenApiOutputDir);
      const daManifestFilePath = path.join(outputFolderPath, defaultDAManifestFileName);

      if (ctx.ui?.runCommand) {
        // 0. Delete output folder if exists
        if (fs.existsSync(outputFolderPath)) {
          this.removeGeneratedFiles(outputFolderPath);
        }

        // 1. Compile tsp file to openapi spec and declarative agent manifest
        const tspRes = await ctx.ui.runCommand({
          cmd: `npx --package=@typespec/compiler tsp compile ${mainFilePath} --config ${args.typeSpecConfigPath}`,
          workingDirectory: projectPath,
        });

        if (tspRes.isErr()) {
          throw tspRes.error;
        }

        // 2. Call Kiota to generate plugin manifest
        const openapiSpecs = fs.readdirSync(openApiSpecsFolderPath);
        if (openapiSpecs.length === 0) {
          throw new NoSpecError(actionName);
        }

        const daManifest = (await fs.readJSON(
          daManifestFilePath
        )) as DeclarativeCopilotManifestSchema;
        const actions = daManifest.actions;
        if (actions && actions.length > 0) {
          if (openapiSpecs.length === 1) {
            // only one openapi spec, the spac name should = openapi.yaml
            const spec = openapiSpecs[0];
            if (actions.length > 1) {
              throw new MultipleActionError(actionName);
            }

            const pluginManifestName = actions[0].id;
            await this.generatePluginManifestWithKiota(
              ctx,
              outputFolderPath,
              openApiSpecsFolderPath,
              spec,
              pluginManifestName
            );
          } else {
            for (const spec of openapiSpecs) {
              const action = actions.find(
                // action.id is the namespace (service name)
                // spec name follows the pattern openapi.{service-name-if-multiple}.json
                // directly match the action.id in the spec name here
                (action: any) =>
                  spec.toLowerCase().includes(`.${(action.id as string).toLowerCase()}.`)
              );
              if (!action) {
                continue;
              }
              const pluginManifestName = action.id;
              await this.generatePluginManifestWithKiota(
                ctx,
                outputFolderPath,
                openApiSpecsFolderPath,
                spec,
                pluginManifestName
              );
            }
          }
        }

        // 3. Update manifest
        const manifestFilePath = path.join(ctx.projectPath, args.manifestPath);
        const generatedManifestFilePath = path.join(
          outputFolderPath,
          path.basename(args.manifestPath)
        );
        const manifest = (await fs.readJSON(manifestFilePath)) as TeamsAppManifest;
        manifest.copilotAgents = manifest.copilotAgents || {};
        manifest.copilotAgents.declarativeAgents = manifest.copilotAgents.declarativeAgents || [];
        manifest.copilotAgents.declarativeAgents = [
          {
            file: defaultDAManifestFileName,
            id: "declarativeAgent",
          },
        ];
        await fs.writeJSON(generatedManifestFilePath, manifest, { spaces: 2 });

        // 4. If env exists in plugin manifest, update yaml file
        const generatedFolder = fs.readdirSync(outputFolderPath);
        let showAlert = false;
        for (const file of generatedFolder) {
          if (file.match(/[^-]+\-apiplugin\.json/)) {
            const pluginManifestPath = path.join(outputFolderPath, file);
            const authData = await parseAndUpdatePluginManifestForKiota(pluginManifestPath, true);
            for (const authInfo of authData) {
              const addAuthRes = await injectAuthAction(
                ctx.projectPath,
                authInfo.authName,
                undefined,
                path.join(outputFolderPath, authInfo.specPath),
                false,
                authInfo.authType === "apiKey" ? "ApiKeyPluginVault" : "OAuthPluginVault",
                false,
                authInfo.registrationId
              );

              if (addAuthRes) {
                showAlert = true;
              }
            }
          }
        }
        if (showAlert) {
          void ctx.ui.showMessage(
            "warn",
            getLocalizedString("driver.typeSpec.compile.reprovision", MetadataV3.configFile),
            false
          );
          return {
            result: err(new ReProvisionError(actionName, MetadataV3.configFile)),
            summaries: summaries,
          };
        }
      }

      if (ctx.platform === Platform.VSCode) {
        (ctx.logProvider as any).outputChannel.show();
      }

      ctx.logProvider?.info(
        getLocalizedString("driver.aadApp.log.successExecuteDriver", actionName)
      );
      return {
        result: ok(outputs),
        summaries: summaries,
      };
    } catch (error) {
      if (error instanceof UserError || error instanceof SystemError) {
        ctx.logProvider?.error(
          getLocalizedString(
            "driver.aadApp.log.failExecuteDriver",
            actionName,
            error.displayMessage
          )
        );
        return {
          result: err(error),
          summaries: summaries,
        };
      }

      const message = JSON.stringify(error);
      ctx.logProvider?.error(
        getLocalizedString("driver.aadApp.log.failExecuteDriver", actionName, message)
      );
      return {
        result: err(assembleError(error as Error, actionName)),
        summaries: summaries,
      };
    }
  }

  private validateArgs(args: TypeSpecCompileArgs): void {
    const invalidParameters: string[] = [];
    if (typeof args.path !== "string" || !args.path) {
      invalidParameters.push("path");
    }

    if (typeof args.manifestPath !== "string" || !args.manifestPath) {
      invalidParameters.push("manifestPath");
    }

    if (typeof args.outputDir !== "string" || !args.outputDir) {
      invalidParameters.push("outputDir");
    }

    if (typeof args.typeSpecConfigPath !== "string" || !args.typeSpecConfigPath) {
      invalidParameters.push("typeSpecConfigPath");
    }

    if (invalidParameters.length > 0) {
      throw new InvalidActionInputError(actionName, invalidParameters, helpLink);
    }
  }

  private async generatePluginManifestWithKiota(
    ctx: DriverContext,
    outputFolderPath: string,
    openApiSpecsFolderPath: string,
    specName: string,
    pluginManifestName: string
  ): Promise<void> {
    const generateRes = await ctx.ui!.runCommand!({
      cmd: `kiota plugin add \
        -d ${openApiSpecsFolderPath}/${specName} \
        --plugin-name ${pluginManifestName} \
        --output ${outputFolderPath} \
        --type apiplugin`,
      workingDirectory: ctx.projectPath,
      env: {
        KIOTA_CONFIG_PREVIEW: "true",
      },
    });

    if (generateRes.isErr()) {
      throw generateRes.error;
    }

    // Remove all plugins from Kiota to avoid error when re-provision
    const removeRes = await ctx.ui!.runCommand!({
      cmd: `kiota plugin remove \
        --plugin-name ${pluginManifestName}`,
      workingDirectory: ctx.projectPath,
      env: {
        KIOTA_CONFIG_PREVIEW: "true",
      },
    });
    if (removeRes.isErr()) {
      throw removeRes.error;
    }
  }

  private removeGeneratedFiles(outputFolderPath: string): void {
    const files = fs.readdirSync(outputFolderPath);
    for (const file of files) {
      if (file === defaultOpenApiOutputDir) {
        const openApiSpecsFolderPath = path.join(outputFolderPath, defaultOpenApiOutputDir);
        fs.rmSync(openApiSpecsFolderPath, { recursive: true, force: true });
      }

      if (
        file === defaultDAManifestFileName ||
        file.match(/[^-]+\-apiplugin\.json/) ||
        file.match(/[^-]+\-openapi\.json/) ||
        file.match(/[^-]+\-openapi\.yaml/) ||
        file.match(/[^-]+\-openapi\.yml/)
      ) {
        const filePath = path.join(outputFolderPath, file);
        fs.rmSync(filePath);
      }
    }
  }
}
