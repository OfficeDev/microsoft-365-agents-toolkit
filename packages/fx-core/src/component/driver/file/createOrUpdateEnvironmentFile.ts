// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import {
  FxError,
  Result,
  SystemError,
  UserError,
  Void,
  ok,
  TextInputQuestion,
} from "@microsoft/teamsfx-api";
import * as dotenv from "dotenv";
import * as fs from "fs-extra";
import * as os from "os";
import * as path from "path";
import { Service } from "typedi";
import { getLocalizedString } from "../../../common/localizeUtils";
import { InvalidActionInputError, assembleError } from "../../../error/common";
import {
  azureOpenAIAssistantIdQuestion,
  azureOpenAIDeploymentNameQuestion,
  azureOpenAIEmbeddingDeploymentNameQuestion,
  azureOpenAIEndpointQuestion,
  azureOpenAIKeyQuestion,
  openAIAssistantIdQuestion,
  openAIKeyQuestion,
} from "../../../question";
import { OpenAIEnvironmentVariables } from "../../constants";
import { wrapRun } from "../../utils/common";
import { pathUtils } from "../../utils/pathUtils";
import { logMessageKeys } from "../aad/utility/constants";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { GenerateEnvArgs } from "./interface/generateEnvArgs";

const actionName = "file/createOrUpdateEnvironmentFile";
const helpLink = "https://aka.ms/teamsfx-actions/file-createOrUpdateEnvironmentFile";

@Service(actionName) // DO NOT MODIFY the service name
export class CreateOrUpdateEnvironmentFileDriver implements StepDriver {
  description = getLocalizedString("driver.file.createOrUpdateEnvironmentFile.description");
  readonly progressTitle = getLocalizedString("driver.file.progressBar.env");

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async run(
    args: GenerateEnvArgs,
    context: DriverContext
  ): Promise<Result<Map<string, string>, FxError>> {
    return wrapRun(async () => {
      const result = await this.handler(args, context);
      return result.output;
    }, actionName);
  }

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  public async execute(args: GenerateEnvArgs, ctx: DriverContext): Promise<ExecutionResult> {
    let summaries: string[] = [];
    const outputResult = await wrapRun(async () => {
      const result = await this.handler(args, ctx);
      summaries = result.summaries;
      return result.output;
    }, actionName);
    return {
      result: outputResult,
      summaries,
    };
  }

  private async handler(
    args: GenerateEnvArgs,
    context: DriverContext
  ): Promise<{
    output: Map<string, string>;
    summaries: string[];
  }> {
    try {
      this.validateArgs(args);
      const target = pathUtils.resolveFilePath(context.projectPath, args.target);
      await fs.ensureFile(target);
      const envs = dotenv.parse(await fs.readFile(target));
      context.logProvider?.debug(`Existing envs: ${JSON.stringify(envs)}`);
      const map = new Map<string, string>();
      const res = await this.askForOpenAIEnvironmentVariables(context, args, map);
      if (res.isErr()) {
        throw res.error;
      }
      const updatedEnvs = Object.entries({ ...envs, ...args.envs }).map(
        ([key, value]) => `${key}=${value}`
      );
      context.logProvider?.debug(`Updated envs: ${JSON.stringify(updatedEnvs)}`);
      await fs.writeFile(target, updatedEnvs.join(os.EOL));
      const envFilePathRes = await pathUtils.getEnvFilePath(
        context.projectPath,
        process.env.TEAMSFX_ENV || "dev"
      );
      if (envFilePathRes.isOk()) {
        if (path.resolve(target) === path.resolve(envFilePathRes.value!)) {
          for (const key of Object.keys(args.envs)) {
            map.set(key, args.envs[key]);
          }
        }
      }
      return {
        output: map,
        summaries: [
          getLocalizedString(
            "driver.file.createOrUpdateEnvironmentFile.summary",
            path.normalize(target)
          ),
        ],
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

  /**
   * Pop up input text to input OpenAI environment variables, or return UserCancel error.
   * @param ctx
   * @param args The arguments passed to the driver.
   * @param envOutput Used to store the resolved environment variables, which will be written to the environment file.
   * @returns
   */
  async askForOpenAIEnvironmentVariables(
    ctx: DriverContext,
    args: GenerateEnvArgs,
    envOutput: Map<string, string>
  ): Promise<Result<Void, FxError>> {
    for (const [envKey, envValue] of Object.entries(args.envs)) {
      const result = await this.processOpenAIEnvironmentVariable(
        ctx,
        args,
        envOutput,
        envKey,
        envValue
      );
      if (result.isErr()) {
        return result;
      }
    }
    return ok(Void);
  }

  private async processOpenAIEnvironmentVariable(
    ctx: DriverContext,
    args: GenerateEnvArgs,
    envOutput: Map<string, string>,
    envKey: string,
    envValue: string
  ): Promise<Result<Void, FxError>> {
    if (envValue) {
      const placeHolderReg = /\${{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*}}/;
      const matches = placeHolderReg.exec(envValue);
      if (matches != null && matches.length > 1) {
        const envVarName = matches[1];
        const config =
          envVarConfigs.find((c) => c.envKey === envKey) || getGenericEnvVarConfig(envKey);

        const result = await ctx.ui?.inputText({
          name: config.question.name,
          title: config.question.title as string,
          password: config.isPassword || config.question.password,
          validation: config.validation,
        });

        if (result?.isErr()) {
          return result;
        } else if (result?.isOk() && result.value.result) {
          envOutput.set(envVarName, result.value.result);
          args.envs[envKey] = result.value.result;
        }
      }
    }
    return ok(Void);
  }

  private validateArgs(args: GenerateEnvArgs): void {
    const invalidParameters: string[] = [];
    if (!args.target || typeof args.target !== "string" || args.target.length === 0) {
      invalidParameters.push("target");
    }

    if (!args.envs || typeof args.envs !== "object") {
      invalidParameters.push("envs");
    } else {
      for (const value of Object.values(args.envs)) {
        if (value === undefined || value === null || typeof value === "object") {
          invalidParameters.push("envs");
        }
      }
    }

    if (invalidParameters.length > 0) {
      throw new InvalidActionInputError(actionName, invalidParameters, helpLink);
    }
  }
}

interface OpenAIEnvVarConfig {
  envKey: string;
  question: TextInputQuestion;
  validation: (input: string) => string | undefined;
  isPassword?: boolean;
}

const getNonEmptyStringValidation = (key: string): ((input: string) => string | undefined) => {
  return (input: string) => (input.length < 1 ? getLocalizedString(key) : undefined);
};

const getEndpointValidation = (key: string): ((input: string) => string | undefined) => {
  return (input: string) =>
    !input.startsWith("https://") && !input.startsWith("http://")
      ? getLocalizedString(key)
      : undefined;
};

const getGenericEnvVarConfig = (envKey: string): OpenAIEnvVarConfig => ({
  envKey,
  question: {
    type: "text",
    name: envKey,
    title: envKey,
  },
  validation: getNonEmptyStringValidation(
    "driver.file.createOrUpdateEnvironmentFile.genericEnvVar.validation"
  ),
  isPassword: envKey.toLowerCase().includes("key") || envKey.toLowerCase().includes("secret"),
});

const envVarConfigs: OpenAIEnvVarConfig[] = [
  {
    envKey: OpenAIEnvironmentVariables.AZURE_OPENAI_API_KEY,
    question: azureOpenAIKeyQuestion(),
    validation: getNonEmptyStringValidation(
      "driver.file.createOrUpdateEnvironmentFile.OpenAIKey.validation"
    ),
    isPassword: true,
  },
  {
    envKey: OpenAIEnvironmentVariables.AZURE_OPENAI_ENDPOINT,
    question: azureOpenAIEndpointQuestion(),
    validation: getEndpointValidation(
      "driver.file.createOrUpdateEnvironmentFile.OpenAIDeploymentEndpoint.validation"
    ),
  },
  {
    envKey: OpenAIEnvironmentVariables.AZURE_OPENAI_DEPLOYMENT_NAME,
    question: azureOpenAIDeploymentNameQuestion(),
    validation: getNonEmptyStringValidation(
      "driver.file.createOrUpdateEnvironmentFile.OpenAIDeploymentName.validation"
    ),
  },
  {
    envKey: OpenAIEnvironmentVariables.AZURE_OPENAI_MODEL_DEPLOYMENT_NAME,
    question: azureOpenAIDeploymentNameQuestion(),
    validation: getNonEmptyStringValidation(
      "driver.file.createOrUpdateEnvironmentFile.OpenAIDeploymentName.validation"
    ),
  },
  {
    envKey: OpenAIEnvironmentVariables.OPENAI_API_KEY,
    question: openAIKeyQuestion(),
    validation: getNonEmptyStringValidation(
      "driver.file.createOrUpdateEnvironmentFile.OpenAIKey.validation"
    ),
    isPassword: true,
  },
  {
    envKey: OpenAIEnvironmentVariables.OPENAI_ASSISTANT_ID,
    question: openAIAssistantIdQuestion(),
    validation: getNonEmptyStringValidation(
      "driver.file.createOrUpdateEnvironmentFile.OpenAIAssistantID.validation"
    ),
  },
  {
    envKey: OpenAIEnvironmentVariables.AZURE_OPENAI_ASSISTANT_ID,
    question: azureOpenAIAssistantIdQuestion(),
    validation: getNonEmptyStringValidation(
      "driver.file.createOrUpdateEnvironmentFile.OpenAIAssistantID.validation"
    ),
  },
  {
    envKey: OpenAIEnvironmentVariables.AZURE_OPENAI_EMBEDDING_DEPLOYMENT,
    question: azureOpenAIEmbeddingDeploymentNameQuestion(),
    validation: getNonEmptyStringValidation(
      "driver.file.createOrUpdateEnvironmentFile.OpenAIEmbeddingDeploymentName.validation"
    ),
  },
  {
    envKey: OpenAIEnvironmentVariables.AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME,
    question: azureOpenAIEmbeddingDeploymentNameQuestion(),
    validation: getNonEmptyStringValidation(
      "driver.file.createOrUpdateEnvironmentFile.OpenAIEmbeddingDeploymentName.validation"
    ),
  },
];
