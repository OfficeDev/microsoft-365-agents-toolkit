// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { err, FxError, ok, Result, UserError, Void } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../common/localizeUtils";
import { SolutionSource } from "./constants";
import { DriverContext } from "./driver/interface/commonArgs";
import {
  azureOpenAIKeyQuestion,
  azureOpenAIDeploymentNameQuestion,
  azureOpenAIEndpointQuestion,
  openAIKeyQuestion,
} from "../question";
import { OpenAIEnvironmentVariables } from "./constants";

class DeployUtils {
  async askForDeployConsentV3(ctx: DriverContext): Promise<Result<Void, FxError>> {
    const msg = getLocalizedString("core.deploy.confirmEnvNoticeV3", process.env.TEAMSFX_ENV);
    const deployOption = getLocalizedString("core.option.deploy");
    const result = await ctx.ui?.showMessage("warn", msg, true, deployOption);
    const choice = result?.isOk() ? result.value : undefined;
    if (choice === deployOption) {
      return ok(Void);
    }
    return err(new UserError(SolutionSource, "UserCancel", "UserCancel"));
  }

  /**
   * Pop up input text to input OpenAI environment variables, or return UserCancel error.
   * @param ctx
   * @param unresolvedPlaceholders
   * @param envOutput Used to store the resolved environment variables, which will be written to the environment file.
   * @returns
   */
  async askForOpenAIEnvironmentVariables(
    ctx: DriverContext,
    unresolvedPlaceholders: string[],
    envOutput: Map<string, string>
  ): Promise<Result<Void, FxError>> {
    if (unresolvedPlaceholders.includes(OpenAIEnvironmentVariables.AZURE_OPENAI_API_KEY)) {
      const result = await ctx.ui!.inputText({
        name: azureOpenAIKeyQuestion().name,
        title: azureOpenAIKeyQuestion().title as string,
        password: azureOpenAIKeyQuestion().password,
        validation: (input: string): string | undefined => {
          if (input.length < 1) {
            return getLocalizedString(
              "driver.file.createOrUpdateEnvironmentFile.OpenAIKey.validation"
            );
          }
        },
      });
      if (result.isErr()) {
        return result;
      } else {
        process.env[OpenAIEnvironmentVariables.AZURE_OPENAI_API_KEY] = result.value.result;
        envOutput.set(OpenAIEnvironmentVariables.AZURE_OPENAI_API_KEY, result.value.result!);
      }
    }

    if (unresolvedPlaceholders.includes(OpenAIEnvironmentVariables.AZURE_OPENAI_ENDPOINT)) {
      const result = await ctx.ui!.inputText({
        name: azureOpenAIEndpointQuestion().name,
        title: azureOpenAIEndpointQuestion().title as string,
        validation: (input: string): string | undefined => {
          if (!input.startsWith("https://") && !input.startsWith("http://")) {
            return getLocalizedString(
              "driver.file.createOrUpdateEnvironmentFile.OpenAIDeploymentEndpoint.validation"
            );
          }
        },
      });
      if (result.isErr()) {
        return result;
      } else {
        process.env[OpenAIEnvironmentVariables.AZURE_OPENAI_ENDPOINT] = result.value.result;
        envOutput.set(OpenAIEnvironmentVariables.AZURE_OPENAI_ENDPOINT, result.value.result!);
      }
    }

    if (unresolvedPlaceholders.includes(OpenAIEnvironmentVariables.AZURE_OPENAI_DEPLOYMENT_NAME)) {
      const result = await ctx.ui!.inputText({
        name: azureOpenAIDeploymentNameQuestion().name,
        title: azureOpenAIDeploymentNameQuestion().title as string,
        validation: (input: string): string | undefined => {
          if (input.length < 1) {
            return getLocalizedString(
              "driver.file.createOrUpdateEnvironmentFile.OpenAIDeploymentName.validation"
            );
          }
        },
      });
      if (result.isErr()) {
        return result;
      } else {
        process.env[OpenAIEnvironmentVariables.AZURE_OPENAI_DEPLOYMENT_NAME] = result.value.result;
        envOutput.set(
          OpenAIEnvironmentVariables.AZURE_OPENAI_DEPLOYMENT_NAME,
          result.value.result!
        );
      }
    }

    if (unresolvedPlaceholders.includes(OpenAIEnvironmentVariables.OPENAI_API_KEY)) {
      const result = await ctx.ui!.inputText({
        name: openAIKeyQuestion().name,
        title: openAIKeyQuestion().title as string,
        validation: (input: string): string | undefined => {
          if (input.length < 1) {
            return getLocalizedString(
              "driver.file.createOrUpdateEnvironmentFile.OpenAIKey.validation"
            );
          }
        },
      });
      if (result.isErr()) {
        return result;
      } else {
        process.env[OpenAIEnvironmentVariables.OPENAI_API_KEY] = result.value.result;
        envOutput.set(OpenAIEnvironmentVariables.OPENAI_API_KEY, result.value.result!);
      }
    }
    return ok(Void);
  }
}
export const deployUtils = new DeployUtils();
