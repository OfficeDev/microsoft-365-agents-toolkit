// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author Ning Liu <nliu@microsoft.com>
 */

import { hooks } from "@feathersjs/hooks/lib";
import {
  Colors,
  FxError,
  LogLevel,
  Platform,
  Result,
  TeamsManifestConverter,
  err,
  ok,
} from "@microsoft/teamsfx-api";
import AdmZip from "adm-zip";
import fs from "fs-extra";
import { merge } from "lodash";
import { EOL } from "os";
import * as path from "path";
import { Service } from "typedi";
import { teamsDevPortalClient } from "../../../client/teamsDevPortalClient";
import { AppStudioScopes } from "../../../common/constants";
import { getDefaultString, getLocalizedString } from "../../../common/localizeUtils";
import { FileNotFoundError, InvalidActionInputError } from "../../../error/common";
import { SummaryConstant } from "../../configManager/constant";
import { metadataUtil } from "../../utils/metadataUtil";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { WrapDriverContext } from "../util/wrapUtil";
import { Constants, GeneralValidationErrorId } from "./constants";
import { AppStudioError } from "./errors";
import { ValidateAppPackageArgs } from "./interfaces/ValidateAppPackageArgs";
import { AppStudioResultFactory } from "./results";
import { TelemetryPropertyKey } from "./utils/telemetry";
import { manifestUtils } from "./utils/ManifestUtils";
import { ODRProvider } from "../../utils/odrProvider";
import { exec } from "child_process";
import { promisify } from "util";
import { LocalMcpPrefix } from "../../constants";

const actionName = "teamsApp/validateAppPackage";

@Service(actionName)
export class ValidateAppPackageDriver implements StepDriver {
  description = getLocalizedString("driver.teamsApp.description.validateDriver");
  readonly progressTitle = getLocalizedString(
    "plugins.appstudio.validateAppPackage.progressBar.message"
  );

  public async execute(
    args: ValidateAppPackageArgs,
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
    args: ValidateAppPackageArgs,
    context: WrapDriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    const result = this.validateArgs(args);
    if (result.isErr()) {
      return err(result.error);
    }

    let appPackagePath = args.appPackagePath;
    if (!path.isAbsolute(appPackagePath)) {
      appPackagePath = path.join(context.projectPath, appPackagePath);
    }
    if (!(await fs.pathExists(appPackagePath))) {
      return err(new FileNotFoundError(actionName, appPackagePath));
    }
    const archivedFile = await fs.readFile(appPackagePath);

    const zipEntries = new AdmZip(archivedFile).getEntries();
    const manifestFile = zipEntries.find((x) => x.entryName === Constants.MANIFEST_FILE);

    const pluginFile = zipEntries.find((x) => x.entryName === "ai-plugin.json");
    if (pluginFile) {
      const isValid = await this.verifyLocalPluginCerts(pluginFile);

      if (!isValid) {
        const message =
          "MSIX package certificate verification failed. One or more local mcp servers do not have a cert or have a self-signed cert.";
        return err(
          AppStudioResultFactory.UserError(AppStudioError.ValidationFailedError.name, [
            message,
            message,
          ])
        );
      }
    }

    if (manifestFile) {
      const manifestContent = manifestFile.getData().toString();
      const manifest = TeamsManifestConverter.jsonToManifest(manifestContent);
      metadataUtil.parseManifest(manifest);

      // Add common properties like isCopilotPlugin: boolean
      const manifestTelemetries = manifestUtils.parseCommonTelemetryProperties(manifest as any);
      merge(context.telemetryProperties, manifestTelemetries);
    }

    const appStudioTokenRes = await context.m365TokenProvider.getAccessToken({
      scopes: AppStudioScopes,
    });
    if (appStudioTokenRes.isErr()) {
      return err(appStudioTokenRes.error);
    }
    const appStudioToken = appStudioTokenRes.value;

    try {
      const validationResult = await teamsDevPortalClient.partnerCenterAppPackageValidation(
        appStudioToken,
        archivedFile
      );

      if (context.platform === Platform.CLI) {
        const outputMessage: Array<{ content: string; color: Colors }> = [
          {
            content:
              "Microsoft 365 Agents Toolkit has checked against all validation rules:\n\nSummary: \n",
            color: Colors.BRIGHT_WHITE,
          },
        ];
        if (validationResult.errors.length > 0) {
          outputMessage.push({
            content: `${validationResult.errors.length} failed, `,
            color: Colors.BRIGHT_RED,
          });
          merge(context.telemetryProperties, {
            [TelemetryPropertyKey.validationErrors]: validationResult.errors
              .map((x) => x.title)
              .join(";"),
          });
        }
        if (validationResult.warnings.length > 0) {
          outputMessage.push({
            content:
              `${validationResult.warnings.length} warning` +
              (validationResult.warnings.length > 1 ? "s" : "") +
              ", ",
            color: Colors.BRIGHT_YELLOW,
          });
          merge(context.telemetryProperties, {
            [TelemetryPropertyKey.validationWarnings]: validationResult.warnings
              .map((x) => x.title)
              .join(";"),
          });
        }
        outputMessage.push({
          content: `${validationResult.notes.length} passed.\n`,
          color: Colors.BRIGHT_GREEN,
        });

        if (validationResult.errors.length > 0 || validationResult.warnings.length > 0) {
          outputMessage.push({
            content: getDefaultString(
              "driver.teamsApp.summary.validate.checkPath",
              args.appPackagePath
            ),
            color: Colors.BRIGHT_WHITE,
          });
        }

        validationResult.errors.map((error) => {
          outputMessage.push({ content: `${SummaryConstant.Failed} `, color: Colors.BRIGHT_RED });
          outputMessage.push({
            content:
              error.id === GeneralValidationErrorId && error.code
                ? `${this.processErrorCode(error.code)}`
                : `${error.content} \nFile path: ${error.filePath}, title: ${error.title}`,
            color: Colors.BRIGHT_WHITE,
          });
          if (error.helpUrl) {
            outputMessage.push({
              content: `\n${getLocalizedString("core.option.learnMore")}: `,
              color: Colors.BRIGHT_WHITE,
            });
            outputMessage.push({ content: `${error.helpUrl}\n`, color: Colors.BRIGHT_CYAN });
          }
        });
        validationResult.warnings.map((warning) => {
          outputMessage.push({
            content: `${SummaryConstant.NotExecuted} `,
            color: Colors.BRIGHT_YELLOW,
          });
          outputMessage.push({
            content: `${warning.content}  \nFile path: ${warning.filePath}, title: ${warning.title}`,
            color: Colors.BRIGHT_WHITE,
          });
          if (warning.helpUrl) {
            outputMessage.push({
              content: `\n${getLocalizedString("core.option.learnMore")}: `,
              color: Colors.BRIGHT_WHITE,
            });
            outputMessage.push({ content: `${warning.helpUrl}\n`, color: Colors.BRIGHT_CYAN });
          }
        });
        validationResult.notes.map((note) => {
          // It might be undefined in some cases
          if (note.content) {
            outputMessage.push({
              content: `${SummaryConstant.Succeeded} `,
              color: Colors.BRIGHT_GREEN,
            });
            outputMessage.push({
              content: `${note.content}\n`,
              color: Colors.BRIGHT_WHITE,
            });
          }
        });
        context.ui?.showMessage("info", outputMessage, false);
        if (validationResult.errors.length > 0) {
          const message = `Microsoft 365 Agents Toolkit has completed checking your app package against validation rules. ${validationResult.errors.length} failed.`;
          return err(
            AppStudioResultFactory.UserError(AppStudioError.ValidationFailedError.name, [
              message,
              message,
            ])
          );
        } else {
          return ok(new Map());
        }
      } else {
        // logs in output window
        const errors = validationResult.errors
          .map((error) => {
            const errorContent =
              error.id === GeneralValidationErrorId && error.code
                ? this.processErrorCode(error.code)
                : `${error.content} \n${getLocalizedString(
                    "error.teamsApp.validate.details",
                    error.filePath,
                    error.title
                  )}`;
            let message = `${SummaryConstant.Failed} ${errorContent}\n`;
            if (error.helpUrl) {
              message += getLocalizedString("core.option.learnMore", error.helpUrl);
            }
            return message;
          })
          .join(EOL);
        const warnings = validationResult.warnings
          .map((warning) => {
            let message = `${SummaryConstant.NotExecuted} ${warning.content} \n${getLocalizedString(
              "error.teamsApp.validate.details",
              warning.filePath,
              warning.title
            )} \n`;
            if (warning.helpUrl) {
              message += getLocalizedString("core.option.learnMore", warning.helpUrl);
            }
            return message;
          })
          .join(EOL);
        const notes = validationResult.notes
          .filter((note) => {
            return note.content !== undefined;
          })
          .map((note) => {
            return `${SummaryConstant.Succeeded} ${note.content}`;
          })
          .join(EOL);

        const passed = validationResult.notes.length;
        const failed = validationResult.errors.length;
        const warns = validationResult.warnings.length;
        const summaryStr = [];
        if (failed > 0) {
          summaryStr.push(getLocalizedString("driver.teamsApp.summary.validate.failed", failed));
          merge(context.telemetryProperties, {
            [TelemetryPropertyKey.validationErrors]: validationResult.errors
              .map((x) => x.title)
              .join(";"),
          });
        }
        if (warns > 0) {
          summaryStr.push(
            getLocalizedString("driver.teamsApp.summary.validate.warning", warns) +
              (warns > 1 ? "s" : "")
          );
          merge(context.telemetryProperties, {
            [TelemetryPropertyKey.validationWarnings]: validationResult.warnings
              .map((x) => x.title)
              .join(";"),
          });
        }
        if (passed > 0) {
          summaryStr.push(getLocalizedString("driver.teamsApp.summary.validate.succeed", passed));
        }

        const outputMessage =
          EOL +
          getLocalizedString(
            "driver.teamsApp.summary.validate",
            summaryStr.join(", "),
            errors.length > 0 || warnings.length > 0
              ? getLocalizedString(
                  "driver.teamsApp.summary.validate.checkPath",
                  args.appPackagePath
                )
              : "",
            errors,
            warnings,
            path.resolve(context.logProvider?.getLogFilePath())
          );
        context.logProvider?.info(outputMessage);
        // logs in log file
        await context.logProvider?.logInFile(
          LogLevel.Info,
          `${outputMessage}\n${errors}\n${warnings}\n${notes}`
        );

        const defaultMesage = getDefaultString(
          "driver.teamsApp.validate.result",
          summaryStr.join(", ")
        );
        const displayMessage = getLocalizedString(
          "driver.teamsApp.validate.result.display",
          summaryStr.join(", ")
        );
        if (args.showMessage) {
          // For non-lifecycle commands, just show the message
          if (validationResult.errors.length > 0) {
            return err(
              AppStudioResultFactory.UserError(AppStudioError.ValidationFailedError.name, [
                defaultMesage,
                displayMessage,
              ])
            );
          } else if (validationResult.warnings.length > 0) {
            context.ui?.showMessage("warn", displayMessage, false);
          } else {
            context.ui?.showMessage("info", displayMessage, false);
          }
        } else {
          // For lifecycle like provision, stop-on-error
          if (validationResult.errors.length > 0) {
            return err(
              AppStudioResultFactory.UserError(AppStudioError.ValidationFailedError.name, [
                defaultMesage,
                displayMessage,
              ])
            );
          }
        }
      }
    } catch (e: any) {
      context.logProvider?.warning(
        getLocalizedString("error.teamsApp.validate.apiFailed", e.message)
      );
      context.ui?.showMessage(
        "warn",
        getLocalizedString("error.teamsApp.validate.apiFailed.display"),
        false
      );
    }
    return ok(new Map());
  }

  private validateArgs(args: ValidateAppPackageArgs): Result<any, FxError> {
    if (!args || !args.appPackagePath) {
      return err(
        new InvalidActionInputError(
          actionName,
          ["appPackagePath"],
          "https://aka.ms/teamsfx-actions/teamsapp-validate"
        )
      );
    }
    return ok(undefined);
  }

  private processErrorCode(errorCode: string): string {
    if (errorCode.startsWith("Invalid TypeB ")) {
      // A temporary solution to update the error message.
      return errorCode.substring(0, 8) + "API " + errorCode.substring(14);
    } else {
      return errorCode;
    }
  }

  private async verifyLocalPluginCerts(pluginFile: AdmZip.IZipEntry): Promise<boolean> {
    const pluginContent = pluginFile.getData().toString();
    const pluginContentAsJson = JSON.parse(pluginContent);

    const servers = await ODRProvider.listServers();

    let allValidCerts = true;

    pluginContentAsJson.runtimes
      .filter((runtime: any) => runtime.type === "LocalPlugin")
      .forEach(async (runtime: any) => {
        const serverInfo = servers.find(
          (x) => x.identifier === runtime.spec.local_endpoint.substring(LocalMcpPrefix.length)
        );
        if (serverInfo) {
          const valid = await this.verifyPackageFamilyCertIsValid(serverInfo.packageFamily);

          if (!valid) {
            allValidCerts = false;
          }
        } else {
          allValidCerts = false;
        }
      });

    return allValidCerts;
  }

  private async verifyPackageFamilyCertIsValid(packageName: string): Promise<boolean> {
    const execAsync = promisify(exec);
    const command = `powershell.exe -Command "& Get-AppxPackage | where { $_.PackageFamilyName -eq '${packageName}' } | select { $_.SignatureKind }"`;

    try {
      const { stdout } = await execAsync(command);

      if (!stdout) {
        return false;
      }

      if (stdout.toLowerCase().includes("developer")) {
        return false;
      }
      return true;
    } catch (error) {
      console.error("Unable to get cert info for package name", error);
      return false;
    }
  }
}
