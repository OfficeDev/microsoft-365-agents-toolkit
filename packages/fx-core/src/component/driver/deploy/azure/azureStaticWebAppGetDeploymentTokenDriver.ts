// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Service } from "typedi";
import { ExecutionResult, StepDriver } from "../../interface/stepDriver";
import { DriverContext } from "../../interface/commonArgs";
import { hooks } from "@feathersjs/hooks";
import { addStartAndEndTelemetry } from "../../middleware/addStartAndEndTelemetry";
import { TelemetryConstant } from "../../../constant/commonConstant";
import { asFactory, asString, errorHandle } from "../../../utils/common";
import { AzureStaticWebAppConfigArgs } from "../../interface/provisionArgs";
import {
  getAzureAccountCredential,
  parseAzureResourceId,
} from "../../../utils/azureResourceOperation";
import { WebSiteManagementClient } from "@azure/arm-appservice";
import { FxError, ok, Result } from "@microsoft/teamsfx-api";
import { getLocalizedString } from "../../../../common/localizeUtils";

const ACTION_NAME = "azureStaticWebApps/getDeploymentToken";

@Service(ACTION_NAME)
export class AzureStaticWebAppGetDeploymentTokenDriver implements StepDriver {
  readonly description: string = getLocalizedString("driver.deploy.getSWADeploymentToken");

  protected static readonly STORAGE_CONFIG_ARGS = asFactory<AzureStaticWebAppConfigArgs>({
    resourceId: asString,
  });
  protected static readonly HELP_LINK = "https://aka.ms/teamsfx-actions/swa-get-deployment-key";
  protected static readonly RESOURCE_PATTERN =
    /\/subscriptions\/([^\/]*)\/resourceGroups\/([^\/]*)\/providers\/Microsoft.Web\/staticSites\/([^\/]*)/i;

  @hooks([addStartAndEndTelemetry(ACTION_NAME, TelemetryConstant.DEPLOY_COMPONENT_NAME)])
  async execute(
    args: unknown,
    ctx: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    try {
      return await AzureStaticWebAppGetDeploymentTokenDriver.run(args, ctx, outputEnvVarNames);
    } catch (e) {
      return { result: await errorHandle(e, ACTION_NAME, ctx.logProvider), summaries: [] };
    }
  }

  static async run(
    args: unknown,
    ctx: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const input = AzureStaticWebAppGetDeploymentTokenDriver.STORAGE_CONFIG_ARGS(
      args,
      AzureStaticWebAppGetDeploymentTokenDriver.HELP_LINK
    );
    const outputKey = !outputEnvVarNames?.get("deploymentToken")
      ? "SECRET_TAB_SWA_DEPLOYMENT_TOKEN"
      : outputEnvVarNames.get("deploymentToken")!;
    const resourceInfo = parseAzureResourceId(
      input.resourceId,
      AzureStaticWebAppGetDeploymentTokenDriver.RESOURCE_PATTERN
    );
    const azureTokenCredential = await getAzureAccountCredential(ctx.azureAccountProvider);
    const client = new WebSiteManagementClient(azureTokenCredential, resourceInfo.subscriptionId);
    const secrets = await client.staticSites.listStaticSiteSecrets(
      resourceInfo.resourceGroupName,
      resourceInfo.instanceId,
      {
        requestOptions: {
          customHeaders: {
            "User-Agent": "TeamsToolkit",
          },
        },
      }
    );
    const deploymentKey = secrets?.properties?.apiKey ?? "";
    const result: Result<Map<string, string>, FxError> = ok(new Map([[outputKey, deploymentKey]]));
    // always set the deployment token to the output environment variable
    process.env[outputKey] = deploymentKey;
    return {
      result: result,
      summaries: [getLocalizedString("driver.deploy.getSWADeploymentTokenSummary")],
    };
  }
}
