// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  Colors,
  err,
  FxError,
  ManifestUtil,
  ok,
  Platform,
  Result,
  TeamsAppManifest,
  TeamsManifest,
  TeamsManifestV1D19,
  TeamsManifestV1D5,
} from "@microsoft/teamsfx-api";
import { merge } from "lodash";
import { EOL } from "os";
import path from "path";
import semver from "semver";
import { Service } from "typedi";
import { HelpLinks } from "../../../common/constants";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";
import { InvalidActionInputError } from "../../../error/common";
import { SummaryConstant } from "../../configManager/constant";
import { getAbsolutePath } from "../../utils/common";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { WrapDriverContext } from "../util/wrapUtil";
import { AppStudioError } from "./errors";
import { ValidateManifestArgs } from "./interfaces/ValidateManifestArgs";
import { AppStudioResultFactory } from "./results";
import { copilotGptManifestUtils } from "./utils/CopilotGptManifestUtils";
import { manifestUtils } from "./utils/ManifestUtils";
import { TelemetryPropertyKey } from "./utils/telemetry";

const actionName = "teamsApp/validateManifest";

@Service(actionName)
export class ValidateManifestDriver implements StepDriver {
  description = getLocalizedString("driver.teamsApp.description.validateDriver");
  readonly progressTitle = getLocalizedString(
    "plugins.appstudio.validateManifest.progressBar.message"
  );

  public async execute(
    args: ValidateManifestArgs,
    context: DriverContext
  ): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(context, actionName, actionName);
    const res = await this.validate(args, wrapContext);
    return {
      result: res,
      summaries: wrapContext.summaries,
    };
  }

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async validate(
    args: ValidateManifestArgs,
    context: WrapDriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    const result = this.validateArgs(args);
    if (result.isErr()) {
      return err(result.error);
    }
    const manifestPath = getAbsolutePath(args.manifestPath, context.projectPath);
    const manifestRes = await manifestUtils.getManifestV3(manifestPath, context);
    if (manifestRes.isErr()) {
      return err(manifestRes.error);
    }
    const manifest = manifestRes.value;

    let manifestValidationResult;
    const telemetryProperties: Record<string, string> = {};
    if (manifest.$schema) {
      try {
        manifestValidationResult = await ManifestUtil.validateManifest(manifest);
        telemetryProperties[TelemetryPropertyKey.validationErrors] = manifestValidationResult
          .map((r: string) => r.replace(/\//g, ""))
          .join(";");
      } catch (e: any) {
        return err(
          AppStudioResultFactory.UserError(
            AppStudioError.ValidationFailedError.name,
            AppStudioError.ValidationFailedError.message([
              getLocalizedString(
                "error.appstudio.validateFetchSchemaFailed",
                manifest.$schema,
                e.message
              ),
            ]),
            HelpLinks.WhyNeedProvision
          )
        );
      }
    } else {
      return err(
        AppStudioResultFactory.UserError(
          AppStudioError.ValidationFailedError.name,
          AppStudioError.ValidationFailedError.message([
            getLocalizedString("error.appstudio.validateSchemaNotDefined"),
          ]),
          HelpLinks.WhyNeedProvision
        )
      );
    }

    // validate localization files
    const localizationFilesValidationRes = await this.validateLocalizatoinFiles(
      args,
      context,
      manifest
    );
    if (localizationFilesValidationRes.isErr()) {
      return err(localizationFilesValidationRes.error);
    }
    telemetryProperties[TelemetryPropertyKey.localizationValidationErrors] =
      localizationFilesValidationRes.value.error.map((r: string) => r.replace(/\//g, "")).join(";");

    const manifestVersion = semver.coerce(manifest.manifestVersion);
    let declarativeAgents: TeamsManifestV1D19.DeclarativeAgentRef[] | undefined;
    if (manifestVersion && semver.gte(manifestVersion, "1.19.0")) {
      declarativeAgents = (manifest as TeamsManifestV1D19.TeamsManifestV1D19).copilotAgents
        ?.declarativeAgents;
    }
    // Declarative Copilot
    let declarativeAgentValidationResult;
    if (declarativeAgents?.length && declarativeAgents[0].file) {
      const declarativeCopilotPath = path.join(
        path.dirname(manifestPath),
        declarativeAgents[0].file
      );

      const declarativeCopilotValidationRes = await copilotGptManifestUtils.validateAgainstSchema(
        declarativeAgents[0],
        declarativeCopilotPath,
        context
      );
      if (declarativeCopilotValidationRes.isErr()) {
        return err(declarativeCopilotValidationRes.error);
      } else {
        declarativeAgentValidationResult = declarativeCopilotValidationRes.value;
        telemetryProperties[TelemetryPropertyKey.gptValidationErrors] =
          declarativeAgentValidationResult?.validationResult
            .map((r: string) => r.replace(/\//g, ""))
            .join(";");

        if (declarativeAgentValidationResult.actionValidationResult.length > 0) {
          let errors: string[] = [];
          for (
            let index = 0;
            index < declarativeAgentValidationResult.actionValidationResult.length;
            index++
          ) {
            errors = errors.concat(
              declarativeAgentValidationResult.actionValidationResult[index].validationResult.map(
                (r: string) => index.toString() + ":" + r.replace(/\//g, "")
              )
            );
          }

          telemetryProperties[`${TelemetryPropertyKey.gptActionValidationErrors}`] =
            errors.join(";");
        }
      }
    }

    const actionErrorCount =
      declarativeAgentValidationResult?.actionValidationResult
        .filter((o) => o.filePath !== "")
        .reduce((acc, { validationResult }) => acc + validationResult.length, 0) ?? 0;

    const allErrorCount =
      manifestValidationResult.length +
      localizationFilesValidationRes.value.error.length +
      (declarativeAgentValidationResult?.validationResult.length ?? 0) +
      actionErrorCount;

    if (allErrorCount > 0) {
      const summaryStr = getLocalizedString(
        "driver.teamsApp.summary.validate.failed",
        allErrorCount
      );

      if (context.platform === Platform.CLI) {
        const outputMessage: Array<{ content: string; color: Colors }> = [
          {
            content:
              "Microsoft 365 Agents Toolkit has checked manifest(s) with corresponding schema:\n\nSummary: \n",
            color: Colors.BRIGHT_WHITE,
          },
          {
            content: `${allErrorCount} failed.\n`,
            color: Colors.BRIGHT_RED,
          },
        ];

        if (manifestValidationResult.length > 0) {
          outputMessage.push({
            content:
              getDefaultString(
                "driver.teamsApp.summary.validateTeamsManifest.checkPath",
                args.manifestPath
              ) + "\n",
            color: Colors.BRIGHT_WHITE,
          });
          manifestValidationResult.map((error: string) => {
            outputMessage.push({ content: `${SummaryConstant.Failed} `, color: Colors.BRIGHT_RED });
            outputMessage.push({
              content: `${error}\n`,
              color: Colors.BRIGHT_WHITE,
            });
          });
        }
        if (localizationFilesValidationRes.value.error.length > 0) {
          outputMessage.push({
            content:
              getDefaultString(
                "driver.teamsApp.summary.validateTeamsManifest.checkPath",
                localizationFilesValidationRes.value.filePath
              ) + "\n",
            color: Colors.BRIGHT_WHITE,
          });
          localizationFilesValidationRes.value.error.map((error: string) => {
            outputMessage.push({ content: `${SummaryConstant.Failed} `, color: Colors.BRIGHT_RED });
            outputMessage.push({
              content: `${error}\n`,
              color: Colors.BRIGHT_WHITE,
            });
          });
        }
        if (declarativeAgentValidationResult) {
          const validationMessage = copilotGptManifestUtils.logValidationErrors(
            declarativeAgentValidationResult,
            context.platform
          );
          if (validationMessage) {
            outputMessage.push(...(validationMessage as Array<{ content: string; color: Colors }>));
          }
        }

        context.ui?.showMessage("info", outputMessage, false);
      } else {
        // logs in output window
        const teamsManifestErrors = manifestValidationResult
          .map((error: string) => {
            return `${SummaryConstant.Failed} ${error}`;
          })
          .join(EOL);
        let outputMessage =
          EOL + getLocalizedString("driver.teamsApp.summary.validateManifest", summaryStr);

        if (teamsManifestErrors.length > 0) {
          outputMessage +=
            EOL +
            getLocalizedString(
              "driver.teamsApp.summary.validateTeamsManifest.checkPath",
              args.manifestPath
            ) +
            EOL +
            teamsManifestErrors;
        }

        if (localizationFilesValidationRes.value.error.length > 0) {
          const localizationErrors = localizationFilesValidationRes.value.error
            .map((error: string) => {
              return `${SummaryConstant.Failed} ${error}`;
            })
            .join(EOL);
          outputMessage +=
            EOL +
            getLocalizedString(
              "driver.teamsApp.summary.validateTeamsManifest.checkPath",
              localizationFilesValidationRes.value.filePath
            ) +
            EOL +
            localizationErrors;
        }
        if (declarativeAgentValidationResult) {
          const validationMessage = copilotGptManifestUtils.logValidationErrors(
            declarativeAgentValidationResult,
            context.platform
          ) as string;
          if (validationMessage) {
            outputMessage += EOL + validationMessage;
          }
        }

        outputMessage += EOL;

        context.logProvider?.info(outputMessage);
      }

      merge(context.telemetryProperties, telemetryProperties);

      return err(
        AppStudioResultFactory.UserError(AppStudioError.ValidationFailedError.name, [
          getDefaultString("driver.teamsApp.validate.result", summaryStr),
          getLocalizedString("driver.teamsApp.validate.result.display", summaryStr),
        ])
      );
    } else {
      // logs in output window
      const summaryStr = getLocalizedString(
        "driver.teamsApp.summary.validate.succeed",
        getLocalizedString("driver.teamsApp.summary.validate.all")
      );
      const outputMessage =
        EOL + getLocalizedString("driver.teamsApp.summary.validateManifest", summaryStr, "", "");
      context.logProvider?.info(outputMessage);

      const validationSuccess = getLocalizedString(
        "driver.teamsApp.validate.result.display",
        summaryStr
      );
      if (context.platform === Platform.VS) {
        context.logProvider.info(validationSuccess);
      }
      if (args.showMessage) {
        if (context.platform === Platform.CLI) {
          const outputMessage: Array<{ content: string; color: Colors }> = [
            {
              content:
                "Microsoft 365 Agents Toolkit has completed checking your app package against validation rules. " +
                summaryStr +
                ".",
              color: Colors.BRIGHT_GREEN,
            },
          ];
          context.logProvider.info(outputMessage);
        } else {
          context.ui?.showMessage("info", validationSuccess, false);
        }
      }
      return ok(new Map());
    }
  }

  private validateArgs(args: ValidateManifestArgs): Result<any, FxError> {
    if (!args || !args.manifestPath) {
      return err(
        new InvalidActionInputError(
          actionName,
          ["manifestPath"],
          "https://aka.ms/teamsfx-actions/teamsapp-validate"
        )
      );
    }
    return ok(undefined);
  }

  public async validateLocalizatoinFiles(
    args: ValidateManifestArgs,
    context: WrapDriverContext,
    manifest: TeamsManifest
  ): Promise<Result<{ error: string[]; filePath?: string }, FxError>> {
    const manifestVersion = semver.coerce(manifest.manifestVersion);
    let additionalLanguages: TeamsManifestV1D5.AdditionalLanguage[] | undefined;
    if (manifestVersion && semver.gte(manifestVersion, "1.5.0")) {
      additionalLanguages = (manifest as TeamsManifestV1D5.TeamsManifestV1D5).localizationInfo
        ?.additionalLanguages;
    }
    let defaultLanguageFile: string | undefined;
    let defaultLanguageTag: string | undefined;
    if (manifestVersion && semver.gte(manifestVersion, "1.19.0")) {
      defaultLanguageTag = (manifest as TeamsManifestV1D19.TeamsManifestV1D19).localizationInfo
        ?.defaultLanguageTag;
      defaultLanguageFile = (manifest as TeamsManifestV1D19.TeamsManifestV1D19).localizationInfo
        ?.defaultLanguageFile;
    }
    if (additionalLanguages?.length == 0 && !defaultLanguageFile) {
      return ok({ error: [] });
    }
    const languageList = additionalLanguages || [];
    if (defaultLanguageFile && defaultLanguageTag) {
      languageList.push({
        languageTag: defaultLanguageTag,
        file: defaultLanguageFile,
      });
    }
    for (const language of languageList) {
      const filePath = language?.file;
      if (!filePath) {
        return err(
          AppStudioResultFactory.UserError(
            AppStudioError.ValidationFailedError.name,
            AppStudioError.ValidationFailedError.message([
              getLocalizedString("error.appstudio.localizationFile.pathNotDefined", filePath),
            ])
          )
        );
      }
      const localizationFileDir = path.dirname(
        getAbsolutePath(args.manifestPath, context.projectPath)
      );
      const localizationFilePath = getAbsolutePath(filePath, localizationFileDir);

      const resolvedLocFileRes = await manifestUtils.resolveLocFile(localizationFilePath);
      if (resolvedLocFileRes.isErr()) {
        return err(resolvedLocFileRes.error);
      }
      const localizationFile = JSON.parse(resolvedLocFileRes.value) as TeamsAppManifest;
      try {
        const schema = await ManifestUtil.fetchSchema(localizationFile);
        // the current localization schema has invalid regex sytax, we need to manually fix the properties temporarily
        const activityDespString =
          "^activities.activityTypes\\[\\b([0-9]|[1-8][0-9]|9[0-9]|1[01][0-9]|12[0-7])\\b]\\.description$";
        const fixedActivityDespString =
          "^activities.activityTypes\\[\\b([0-9]|[1-8][0-9]|9[0-9]|1[01][0-9]|12[0-7])\\b\\]\\.description$";
        if (schema.patternProperties?.[activityDespString]) {
          schema.patternProperties[fixedActivityDespString] =
            schema.patternProperties[activityDespString];
          delete schema.patternProperties[activityDespString];
        }
        const activityTemplateString =
          "^activities.activityTypes\\[\\b([0-9]|[1-8][0-9]|9[0-9]|1[01][0-9]|12[0-7])\\b]\\.templateText$";
        const fixedActivityTemplateString =
          "^activities.activityTypes\\[\\b([0-9]|[1-8][0-9]|9[0-9]|1[01][0-9]|12[0-7])\\b\\]\\.templateText$";
        if (schema.patternProperties?.[activityTemplateString]) {
          schema.patternProperties[fixedActivityTemplateString] =
            schema.patternProperties[activityTemplateString];
          delete schema.patternProperties[activityTemplateString];
        }

        const validationRes = await ManifestUtil.validateManifestAgainstSchema(
          localizationFile,
          schema
        );
        if (validationRes.length > 0) {
          return ok({ error: validationRes, filePath: localizationFilePath });
        }
      } catch (e: any) {
        return err(
          AppStudioResultFactory.UserError(
            AppStudioError.ValidationFailedError.name,
            AppStudioError.ValidationFailedError.message([
              getLocalizedString(
                "error.appstudio.localizationFile.validationException",
                filePath,
                e.message
              ),
            ])
          )
        );
      }
    }
    return ok({ error: [] });
  }
}
