// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author yuqzho@microsoft.com, Ning Tang
 */

import { ProjectType, SpecParserError } from "@microsoft/m365-spec-parser";
import {
  AppPackageFolderName,
  Context,
  err,
  FxError,
  GeneratorResult,
  Inputs,
  ManifestTemplateFileName,
  Platform,
  Result,
} from "@microsoft/teamsfx-api";
import { assembleError } from "../../../error";
import { QuestionNames } from "../../../question";
import { ActionContext } from "../../middleware/actionExecutionMW";
import { DefaultTemplateGenerator } from "../defaultGenerator";
import { TemplateInfo } from "../templates/templateInfo";
import { TemplateNames } from "../templates/templateNames";
import { generateFilesFromApiSpec, getTemplateInfosFromApiSpec } from "./common";
import { convertSpecParserErrorToFxError } from "./helper";
import { featureFlagManager, FeatureFlags } from "../../../common/featureFlags";
import path from "path";
import { EmbeddedKnowledgeLocalDirectoryName } from "../../driver/teamsApp/constants";
import fs from "fs-extra";
import { setGeneralSensitivityLabel } from "../utils";
import { copilotGptManifestUtils } from "../../driver/teamsApp/utils/CopilotGptManifestUtils";

export class DeclarativeAgentWithExistingApiSpecGenerator extends DefaultTemplateGenerator {
  componentName = "da-with-existing-api-generator";

  public override activate(context: Context, inputs: Inputs): boolean {
    return (
      TemplateNames.DeclarativeAgentWithActionFromExistingApiSpec ==
      inputs[QuestionNames.TemplateName]
    );
  }

  public override async getTemplateInfos(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<TemplateInfo[], FxError>> {
    return getTemplateInfosFromApiSpec(context, inputs, ProjectType.Copilot, actionContext);
  }

  public override async post(
    context: Context,
    inputs: Inputs,
    destinationPath: string,
    actionContext?: ActionContext
  ): Promise<Result<GeneratorResult, FxError>> {
    try {
      if (
        featureFlagManager.getBooleanValue(FeatureFlags.EmbeddedKnowledgeEnabled) &&
        (inputs.platform === Platform.CLI || inputs.platform === Platform.VSCode)
      ) {
        // ensure EmbeddedKnwoledge folder exists
        const embeddedKnowledgeFolderPath = path.join(
          destinationPath,
          AppPackageFolderName,
          EmbeddedKnowledgeLocalDirectoryName
        );
        await fs.ensureDir(embeddedKnowledgeFolderPath);
      }

      if (featureFlagManager.getBooleanValue(FeatureFlags.SensitivityLabelEnabled)) {
        const teamsManifestPath = path.join(
          destinationPath,
          AppPackageFolderName,
          ManifestTemplateFileName
        );
        const declarativeCopilotManifestPathRes = await copilotGptManifestUtils.getManifestPath(
          teamsManifestPath
        );
        if (declarativeCopilotManifestPathRes.isErr()) {
          return err(declarativeCopilotManifestPathRes.error);
        }
        // best-effort
        await setGeneralSensitivityLabel(context, declarativeCopilotManifestPathRes.value);
      }

      return await generateFilesFromApiSpec(
        context,
        inputs,
        destinationPath,
        ProjectType.Copilot,
        this.componentName
      );
    } catch (e) {
      let error: FxError;
      if (e instanceof SpecParserError) {
        error = convertSpecParserErrorToFxError(e);
      } else {
        error = assembleError(e);
      }
      return err(error);
    }
  }
}
