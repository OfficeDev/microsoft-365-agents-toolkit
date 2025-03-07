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
import {
  defaultDAManifestFileName,
  defaultOpenApiOutputDir,
  defaultOutputDir,
  helpLink,
} from "./constants";
import { NoSpecError } from "./error/noSpecError";
import { NoActionError } from "./error/noActionError";
import { MultipleActionError } from "./error/multipleActionError";

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
      const appPackageFolderPath = path.join(projectPath, "appPackage");
      const outputFolderPath = path.join(appPackageFolderPath, defaultOutputDir);
      const openApiSpecsFolderPath = path.join(outputFolderPath, defaultOpenApiOutputDir);
      const daManifestFilePath = path.join(outputFolderPath, defaultDAManifestFileName);

      if (ctx.ui?.runCommand) {
        // 0. Delete output folder if exists
        if (fs.existsSync(outputFolderPath)) {
          fs.rmSync(outputFolderPath, { recursive: true, force: true });
        }

        // 1. Compile tsp file to openapi spec and declarative agent manifest
        const tspRes = await ctx.ui.runCommand({
          cmd: `npx --package=@typespec/compiler tsp compile ${mainFilePath} \
            --emit @typespec/openapi3 \
            --emit @microsoft/typespec-copilot-skills \
            --options @microsoft/typespec-copilot-skills.file-type=json \
            --options @microsoft/typespec-copilot-skills.output-file=${daManifestFilePath} \
            --options @microsoft/typespec-copilot-skills.emitter-output-dir=${outputFolderPath} \
            --options @typespec/openapi3.emitter-output-dir=${openApiSpecsFolderPath}`,
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
        if (!actions || actions.length === 0) {
          throw new NoActionError(actionName);
        }

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
            const action = actions.find((action: any) =>
              spec.toLowerCase().includes((action.id as string).toLowerCase())
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
        getLocalizedString("driver.aadApp.log.failExecuteDriver", actionName, error.displayMessage)
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
      cmd: `npx --package=@microsoft/kiota-bundle kiota plugin add \
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
      cmd: `npx --package=@microsoft/kiota-bundle kiota plugin remove \
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
}
