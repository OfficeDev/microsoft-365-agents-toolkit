// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/* eslint-disable @typescript-eslint/no-non-null-assertion */
import * as fs from "fs-extra";
import { Service } from "typedi";

import { hooks } from "@feathersjs/hooks/lib";
import { Colors, FxError, Result, SystemError, UserError } from "@microsoft/teamsfx-api";

import { getLocalizedString } from "../../../common/localizeUtils";
import { AppScope, PackageService } from "../../m365/packageService";
import { MosServiceEndpoint, MosServiceScope } from "../../m365/serviceConstant";
import { FileNotFoundError, InvalidActionInputError, assembleError } from "../../../error/common";
import { getAbsolutePath, wrapRun } from "../../utils/common";
import { logMessageKeys } from "../aad/utility/constants";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";

export interface ShareArgs {
  appPackagePath?: string; // The path of the app package
}

export const actionName = "teamsApp/shareToOthers";

const outputKeys = {
  titleId: "titleId",
  appId: "appId",
  shareLink: "shareLink",
};

@Service(actionName) // DO NOT MODIFY the service name
export class ShareToOthersDriver implements StepDriver {
  description = getLocalizedString("driver.shareToOthers.description");
  readonly progressTitle = getLocalizedString("driver.shareToOthers.progress.message");

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async run(
    args: ShareArgs,
    context: DriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    return wrapRun(async () => {
      const result = await this.handler(args, context);
      return result.output;
    }, actionName);
  }

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async execute(
    args: ShareArgs,
    ctx: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    let summaries: string[] = [];
    const outputResult = await wrapRun(async () => {
      const result = await this.handler(args, ctx, outputEnvVarNames);
      summaries = result.summaries;
      return result.output;
    }, actionName);
    return {
      result: outputResult,
      summaries,
    };
  }

  private async handler(
    args: ShareArgs,
    context: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<{
    output: Map<string, string>;
    summaries: string[];
  }> {
    try {
      this.validateArgs(args);
      this.validateOutputEnvVarNames(outputEnvVarNames);
      const appPackagePath = getAbsolutePath(args.appPackagePath!, context.projectPath);
      if (!(await fs.pathExists(appPackagePath))) {
        throw new FileNotFoundError(actionName, appPackagePath);
      }

      // get sideloading service settings
      const sideloadingServiceEndpoint =
        process.env.SIDELOADING_SERVICE_ENDPOINT ?? MosServiceEndpoint;
      const sideloadingServiceScope = process.env.SIDELOADING_SERVICE_SCOPE ?? MosServiceScope;

      const packageService = new PackageService(sideloadingServiceEndpoint, context.logProvider);
      const sideloadingTokenRes = await context.m365TokenProvider.getAccessToken({
        scopes: [sideloadingServiceScope],
      });
      if (sideloadingTokenRes.isErr()) {
        throw sideloadingTokenRes.error;
      }
      const sideloadingToken = sideloadingTokenRes.value;
      const sideloadingRes = await packageService.sideLoading(
        sideloadingToken,
        appPackagePath,
        AppScope.Shared
      );

      const shareSuccess = [
        { content: "(√)Done: ", color: Colors.BRIGHT_GREEN },
        { content: "Share Link: ", color: Colors.BRIGHT_WHITE },
        { content: sideloadingRes[2], color: Colors.BRIGHT_MAGENTA },
      ];
      context.logProvider?.info(shareSuccess);

      return {
        output: new Map([
          [outputEnvVarNames!.get(outputKeys.titleId)!, sideloadingRes[0]],
          [outputEnvVarNames!.get(outputKeys.appId)!, sideloadingRes[1]],
          [outputEnvVarNames!.get(outputKeys.shareLink)!, sideloadingRes[2]],
        ]),
        summaries: [getLocalizedString("driver.shareToOthers.summary", sideloadingRes[2])],
      };
    } catch (error) {
      if (error instanceof UserError || error instanceof SystemError) {
        context.logProvider?.error(
          getLocalizedString(logMessageKeys.failExecuteDriver, actionName, error.displayMessage)
        );
        throw error;
      }
      const message = JSON.stringify(error);
      context.logProvider?.error(
        getLocalizedString(logMessageKeys.failExecuteDriver, actionName, message)
      );
      throw assembleError(error as Error, actionName);
    }
  }

  private validateArgs(args: ShareArgs): void {
    const invalidParameters: string[] = [];

    if (!args.appPackagePath || typeof args.appPackagePath !== "string") {
      invalidParameters.push("appPackagePath");
    }

    if (invalidParameters.length > 0) {
      throw new InvalidActionInputError(actionName, invalidParameters);
    }
  }

  private validateOutputEnvVarNames(outputEnvVarNames?: Map<string, string>): void {
    if (
      !outputEnvVarNames?.get(outputKeys.titleId) ||
      !outputEnvVarNames.get(outputKeys.appId) ||
      !outputEnvVarNames.get(outputKeys.shareLink)
    ) {
      throw new InvalidActionInputError(actionName, ["writeToEnvironmentFile"]);
    }
  }
}
