// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
/**
 * @author Siglud <siglud@gmail.com>
 */

import { Service } from "typedi";
import { ExecutionResult, StepDriver } from "../../interface/stepDriver";
import { DriverContext } from "../../interface/commonArgs";
import { hooks } from "@feathersjs/hooks";
import { addStartAndEndTelemetry } from "../../middleware/addStartAndEndTelemetry";
import { TelemetryConstant } from "../../../constant/commonConstant";
import { getLocalizedString } from "../../../../common/localizeUtils";
import { AzureZipDeployImpl } from "./impl/AzureZipDeployImpl";

const ACTION_NAME = "azureFunctions/zipDeploy";

@Service(ACTION_NAME)
export class AzureFunctionDeployDriver implements StepDriver {
  readonly description: string = getLocalizedString(
    "driver.deploy.deployToAzureFunctionsDescription"
  );
  private static readonly SERVICE_NAME = "Azure Function App";
   
  private static readonly SUMMARY = ["driver.deploy.azureFunctionsDeployDetailSummary"];
  private static readonly SUMMARY_PREPARE = ["driver.deploy.notice.deployDryRunComplete"];

  @hooks([addStartAndEndTelemetry(ACTION_NAME, TelemetryConstant.DEPLOY_COMPONENT_NAME)])
  async execute(args: unknown, ctx: DriverContext): Promise<ExecutionResult> {
    const impl = new AzureZipDeployImpl(
      args,
      ctx,
      AzureFunctionDeployDriver.SERVICE_NAME,
      "https://aka.ms/teamsfx-actions/azure-functions-deploy",
      AzureFunctionDeployDriver.SUMMARY,
      AzureFunctionDeployDriver.SUMMARY_PREPARE
    );
    return await impl.run();
  }
}
