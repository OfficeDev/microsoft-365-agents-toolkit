// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * @author xzf0587 <zhaofengxu@microsoft.com>
 */
import { Deployment, DeploymentMode, ResourceManagementClient } from "@azure/arm-resources";
import { Context, FxError, Result, SystemError, UserError, err, ok } from "@microsoft/teamsfx-api";
import * as fs from "fs-extra";
import { ConstantString } from "../../../common/constants";
import { getLocalizedString } from "../../../common/localizeUtils";
import { CompileBicepError, DeployArmError } from "../../../error/arm";
import { InvalidAzureCredentialError } from "../../../error/azure";
import { InvalidActionInputError, MissingEnvironmentVariablesError } from "../../../error/common";
import {
  expandEnvironmentVariable,
  getAbsolutePath,
  getEnvironmentVariables,
} from "../../utils/common";
import { cpUtils } from "../../utils/depsChecker/cpUtils";
import { WrapDriverContext } from "../util/wrapUtil";
import { Constants, TelemetryProperties, TemplateType } from "./constant";
import { deployArgs, deploymentOutput, templateArgs } from "./interface";
import { ensureBicepForDriver } from "./util/bicepChecker";
import { ArmErrorHandle, DeployContext } from "./util/handleError";
import { convertOutputs, getFileExtension, hasBicepTemplate } from "./util/util";
import { validateArgs } from "./validator";
import { ErrorContextMW } from "../../../common/globalVars";
import { hooks } from "@feathersjs/hooks";
import { azureClientHelper } from "../../utils/azureClient";

const helpLink = "https://aka.ms/teamsfx-actions/arm-deploy";

export class ArmDeployImpl {
  args: deployArgs;
  context: WrapDriverContext;
  bicepCommand?: string;
  client?: ResourceManagementClient;

  constructor(args: deployArgs, context: WrapDriverContext) {
    this.args = args;
    this.context = context;
  }

  public async run(): Promise<Map<string, string>> {
    await this.validateArgs();
    this.client = await azureClientHelper.createRmClient(
      this.context.azureAccountProvider,
      this.args.subscriptionId
    );
    const needBicepCli = hasBicepTemplate(this.args.templates);

    if (needBicepCli && this.args.bicepCliVersion) {
      this.context.logProvider.debug(
        `Ensure bicep cli version ${this.args.bicepCliVersion} for ${Constants.actionName}`
      );
      this.bicepCommand = await this.ensureBicepCli();
    } else {
      this.bicepCommand = "bicep";
    }
    const deployRes = await this.deployTemplates();
    if (deployRes.isOk()) {
      const outputs = deployRes.value;
      return convertOutputs(outputs);
    } else {
      throw deployRes.error;
    }
  }

  private async validateArgs(): Promise<void> {
    const invalidParameters = await validateArgs(this.args);

    if (invalidParameters.length > 0) {
      throw new InvalidActionInputError(Constants.actionName, invalidParameters, helpLink);
    }
  }

  public async ensureBicepCli(): Promise<string> {
    return await ensureBicepForDriver(this.context, this.args.bicepCliVersion!);
  }

  async deployTemplates(): Promise<Result<deploymentOutput[], FxError>> {
    const outputs: deploymentOutput[] = [];
    this.setTelemetries();
    await Promise.all(
      this.args.templates.map(async (template) => {
        this.context.logProvider.debug(
          `Deploy template ${template.deploymentName} from ${template.path} to resource group ${this.args.resourceGroupName}`
        );
        const res = await this.deployTemplate(template);
        if (res.isOk() && res.value) {
          this.context.addSummary(
            getLocalizedString(
              "core.deployArmTemplates.ActionSuccess",
              this.args.resourceGroupName,
              template.deploymentName
            )
          );
          outputs.push(res.value);
        } else if (res.isErr()) {
          throw res.error;
        }
      })
    );
    return ok(outputs);
  }

  async deployTemplate(
    templateArg: templateArgs
  ): Promise<Result<deploymentOutput | undefined, FxError>> {
    const deployCtx: DeployContext = {
      ctx: this.context as any as Context,
      finished: false,
      deploymentStartTime: Date.now(),
      client: this.client!,
      resourceGroupName: this.args.resourceGroupName,
      deploymentName: templateArg.deploymentName,
    };
    try {
      const parameters = await this.getDeployParameters(templateArg.parameters);
      const template = await this.getDeployTemplate(templateArg.path);
      const deploymentParameters: Deployment = {
        properties: {
          parameters: parameters ? parameters.parameters : null,
          template: template as any,
          mode: "Incremental" as DeploymentMode,
        },
      };
      const res = await this.executeDeployment(templateArg, deploymentParameters, deployCtx);
      return res;
    } catch (error: any) {
      if (error instanceof UserError || error instanceof SystemError) return err(error);
      return err(new DeployArmError(deployCtx.deploymentName, deployCtx.resourceGroupName, error));
    }
  }

  async executeDeployment(
    templateArg: templateArgs,
    deploymentParameters: Deployment,
    deployCtx: DeployContext
  ): Promise<Result<deploymentOutput | undefined, FxError>> {
    try {
      return await this.innerExecuteDeployment(templateArg, deploymentParameters);
    } catch (error) {
      const errRes = ArmErrorHandle.handleArmDeploymentError(error, deployCtx);
      return errRes;
    }
  }
  @hooks([ErrorContextMW({ source: "Azure", component: "ArmDeployImpl" })])
  async innerExecuteDeployment(
    templateArg: templateArgs,
    deploymentParameters: Deployment
  ): Promise<Result<deploymentOutput | undefined, FxError>> {
    const result = await this.client?.deployments.beginCreateOrUpdateAndWait(
      this.args.resourceGroupName,
      templateArg.deploymentName,
      deploymentParameters
    );
    return ok(result?.properties?.outputs);
  }

  async getDeployParameters(parameters?: string): Promise<any> {
    if (!parameters) {
      return null;
    }
    const filePath = getAbsolutePath(parameters, this.context.projectPath);
    const template = await fs.readFile(filePath, ConstantString.UTF8Encoding);
    const parameterJsonString = expandEnvironmentVariable(template);
    this.checkPlaceholderInTemplate(parameterJsonString, filePath);
    return JSON.parse(parameterJsonString);
  }

  checkPlaceholderInTemplate(parameterJsonString: string, filePath: string): void {
    const tokens = getEnvironmentVariables(parameterJsonString);
    if (tokens.length > 0) {
      throw new MissingEnvironmentVariablesError("arm", tokens.join(","), filePath);
    }
  }

  async getDeployTemplate(templatePath: string): Promise<string> {
    const templateType = getFileExtension(templatePath);
    const filePath = getAbsolutePath(templatePath, this.context.projectPath);
    let templateJsonString;
    if (templateType === TemplateType.Bicep) {
      templateJsonString = await this.compileBicepToJson(filePath);
    } else {
      const template = await fs.readFile(filePath, ConstantString.UTF8Encoding);
      templateJsonString = JSON.parse(template);
    }
    return templateJsonString;
  }

  async compileBicepToJson(filePath: string): Promise<JSON> {
    try {
      this.context.logProvider.debug(`Compile bicep template ${filePath} to json`);
      const result = await cpUtils.executeCommand(
        undefined,
        this.context.logProvider,
        { shell: false },
        this.bicepCommand!,
        ...["build", filePath, "--stdout"]
      );
      return JSON.parse(result);
    } catch (err) {
      throw new CompileBicepError(filePath, err as Error);
    }
  }

  private setTelemetries(): void {
    let bicepCount = 0;
    let jsonCount = 0;
    for (const template of this.args.templates) {
      const templateType = getFileExtension(template.path);
      if (templateType === TemplateType.Bicep) {
        bicepCount++;
      } else {
        jsonCount++;
      }
    }
    this.context.addTelemetryProperties({
      [TelemetryProperties.bicepTemplateCount]: bicepCount.toString(),
      [TelemetryProperties.jsonTemplateCount]: jsonCount.toString(),
    });
  }
}
