// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author yefuwang@microsoft.com
 */

import {
  Context,
  devPreview,
  err,
  FxError,
  GeneratorResult,
  Inputs,
  ManifestUtil,
  ok,
  Result,
} from "@microsoft/teamsfx-api";
import { convertProject } from "office-addin-project";
import { getLocalizedString } from "../../../common/localizeUtils";
import { getUuid } from "../../../common/stringUtils";
import { assembleError } from "../../../error";
import { ProgrammingLanguage, QuestionNames } from "../../../question/constants";
import { ActionContext } from "../../middleware/actionExecutionMW";
import { envUtil } from "../../utils/envUtil";
import { DefaultTemplateGenerator } from "../defaultGenerator";
import { TemplateInfo } from "../templates/templateInfo";
import { TemplateNames } from "../templates/templateNames";
import { HelperMethods } from "./helperMethods";

/**
 * case 1: project-type=office-xml-addin-type AND addin-host=outlook
 * case 2: project-type=office-addin-type (addin-host=undefined)
 * case 3: project-type=outlook-addin-type (addin-host=undefined)
 */
export class OfficeAddinGenerator {
  public static async doScaffolding(
    context: Context,
    inputs: Inputs,
    destinationPath: string
  ): Promise<Result<undefined, FxError>> {
    const addinRoot = destinationPath;
    const fromFolder = inputs[QuestionNames.OfficeAddinFolder];
    const workingDir = process.cwd();
    const importProgressStr = getLocalizedString(
      "core.generator.officeAddin.importOfficeProject.title"
    );
    const importProgress = context.userInteraction.createProgressBar(importProgressStr, 3);

    process.chdir(addinRoot);
    try {
      if (fromFolder) {
        await importProgress.start();
        // from existing project
        await importProgress.next(
          getLocalizedString("core.generator.officeAddin.importProject.copyFiles")
        );
        HelperMethods.copyAddinFiles(fromFolder, addinRoot);
        const sourceManifestFile: string = inputs[QuestionNames.OfficeAddinManifest];
        let manifestFile: string = sourceManifestFile.replace(fromFolder, addinRoot);
        await importProgress.next(
          getLocalizedString("core.generator.officeAddin.importProject.convertProject")
        );
        if (manifestFile.endsWith(".xml")) {
          // Need to convert to json project first
          await convertProject(manifestFile, "./backup.zip", "", true);
          manifestFile = manifestFile.replace(/\.xml$/, ".json");
        }
        inputs[QuestionNames.OfficeAddinHost] = await getHost(manifestFile);
        await importProgress.next(
          getLocalizedString("core.generator.officeAddin.importProject.updateManifest")
        );
        await HelperMethods.updateManifest(destinationPath, manifestFile);
      }
      process.chdir(workingDir);
      await importProgress.end(true, true);
      return ok(undefined);
    } catch (e) {
      process.chdir(workingDir);
      await importProgress.end(false, true);
      return err(assembleError(e as Error));
    }
  }
}

// TODO: update to handle different hosts when support for them is implemented
// TODO: handle multiple scopes
type OfficeHost = "Outlook" | "Word" | "Excel" | "PowerPoint"; // | "OneNote" | "Project"
export async function getHost(addinManifestPath: string): Promise<OfficeHost> {
  // Read add-in manifest file
  const addinManifest: devPreview.DevPreviewSchema = await ManifestUtil.loadFromPath(
    addinManifestPath
  );
  let host: OfficeHost = "Outlook";
  switch (addinManifest.extensions?.[0].requirements?.scopes?.[0]) {
    case "document":
      host = "Word";
      break;
    case "mail":
      host = "Outlook";
      break;
    // case "notebook":
    //   host = "OneNote";
    case "presentation":
      host = "PowerPoint";
      break;
    // case "project":
    //   host = "Project";
    case "workbook":
      host = "Excel";
      break;
  }
  return host;
}

export class OfficeAddinGeneratorNew extends DefaultTemplateGenerator {
  componentName = "office-addin-generator";

  // activation condition
  public override activate(context: Context, inputs: Inputs): boolean {
    const templateName = inputs[QuestionNames.TemplateName];
    return [
      TemplateNames.OutlookTaskpane,
      TemplateNames.WXPTaskpane,
      TemplateNames.OfficeAddinCommon,
    ].includes(templateName);
  }

  public override async getTemplateInfos(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<TemplateInfo[], FxError>> {
    const templateName = inputs[QuestionNames.TemplateName];
    const res = await OfficeAddinGenerator.doScaffolding(context, inputs, destinationPath);
    if (res.isErr()) return err(res.error);
    return Promise.resolve(
      ok([
        {
          templateName: templateName,
          language: ProgrammingLanguage.TS,
          replaceMap: { manifestId: getUuid() },
        },
      ])
    );
  }

  public override async post(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<GeneratorResult, FxError>> {
    const fromFolder = inputs[QuestionNames.OfficeAddinFolder];
    if (fromFolder) {
      // reset all env files
      const envRes = await envUtil.listEnv(destinationPath);
      if (envRes.isOk()) {
        const envs = envRes.value;
        for (const env of envs) {
          await envUtil.resetEnv(destinationPath, env, ["TEAMSFX_ENV", "APP_NAME_SUFFIX"]);
        }
      }
    }
    return ok({});
  }
}
