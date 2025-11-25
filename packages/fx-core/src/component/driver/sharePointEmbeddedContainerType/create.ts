// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import { FxError, SystemError, UserError, err, ok, Result } from "@microsoft/teamsfx-api";
import axios from "axios";
import { Service } from "typedi";
import { GraphScopes } from "../../../common/constants";
import { getLocalizedString } from "../../../common/localizeUtils";
import { CreateSPEContainerTypeArgs } from "./interface/createSPEContainerTypeArgs";
import { CreateSPEContainerTypeOutput, OutputKeys } from "./interface/createSPEContainerTypeOutput";
import { SPContainerTypeBillingClassification } from "./interface/sharePointEmbeddedContainerType";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { loadStateFromEnv, mapStateToEnv } from "../util/utils";
import { logMessageKeys, descriptionMessageKeys, progressBarKeys } from "./utility/constants";
import { WrapDriverContext } from "../util/wrapUtil";
import { SPEContainerTypeAppClient } from "./utility/speContainerTypeAppClient";
import { OutputEnvironmentVariableUndefinedError } from "../error/outputEnvironmentVariableUndefinedError";
import { HttpServerError, InvalidActionInputError, assembleError } from "../../../error/common";

const actionName = "sharePointEmbeddedContainerType/create"; // DO NOT MODIFY THE NAME

@Service(actionName) // DO NOT MODIFY THE SERVICE NAME
export class CreateSharePointEmbeddedContainerTypeDriver implements StepDriver {
  // TODO localize these strings
  description = getLocalizedString(descriptionMessageKeys.createContainerType);
  readonly progressTitle = getLocalizedString(progressBarKeys.progressBarTitle);

  public async execute(
    args: CreateSPEContainerTypeArgs,
    context: DriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    const wrapDriverContext = new WrapDriverContext(context, actionName, actionName);
    return await this.executeInternal(args, wrapDriverContext, outputEnvVarNames);
  }

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  private async executeInternal(
    args: CreateSPEContainerTypeArgs,
    context: WrapDriverContext,
    outputEnvVarNames?: Map<string, string>
  ): Promise<ExecutionResult> {
    let outputs: Map<string, string> = new Map<string, string>();
    const summaries: string[] = [];
    if (!outputEnvVarNames) {
      const error = new OutputEnvironmentVariableUndefinedError(actionName);
      context.logProvider?.error(
        getLocalizedString(logMessageKeys.failExecuteDriver, actionName, error.displayMessage)
      );
      return {
        result: err(error),
        summaries: summaries,
      };
    }

    const speContainerTypeState: CreateSPEContainerTypeOutput = loadStateFromEnv(outputEnvVarNames);

    try {
      context.logProvider?.info(getLocalizedString(logMessageKeys.startExecuteDriver, actionName));
      this.validateArgs(args);

      if (!speContainerTypeState.containerTypeId) {
        context.logProvider?.info(
          getLocalizedString(
            logMessageKeys.startCreateContainerType,
            outputEnvVarNames.get(OutputKeys.containerTypeId)
          )
        );

        const speContainerTypeClient = new SPEContainerTypeAppClient(
          context.m365TokenProvider,
          context.logProvider
        );

        const speContainerType = await speContainerTypeClient.createSPEContainerType(
          args.owningApplicationId,
          args.billingClassification,
          args.name,
          args.discoverable
        );

        speContainerTypeState.containerTypeId = speContainerType.id;
        outputs = mapStateToEnv(speContainerTypeState, outputEnvVarNames);

        const summary = getLocalizedString(
          logMessageKeys.successCreateContainerType,
          speContainerType.id
        );

        context.logProvider?.info(summary);
        summaries.push(summary);
      } else {
        context.logProvider?.info(
          getLocalizedString(
            logMessageKeys.skipCreateContainerType,
            outputEnvVarNames.get(OutputKeys.containerTypeId)
          )
        );
      }
      return {
        result: ok(outputs),
        summaries: summaries,
      };
    } catch (error) {
      if (error instanceof UserError || error instanceof SystemError) {
        context.logProvider?.error(
          getLocalizedString(logMessageKeys.failExecuteDriver, actionName, error.displayMessage)
        );
        return {
          result: err(error),
          summaries: summaries,
        };
      }

      if (axios.isAxiosError(error)) {
        const message = JSON.stringify(error.response!.data);
        context.logProvider?.error(
          getLocalizedString(logMessageKeys.failExecuteDriver, actionName, message)
        );
        return {
          result: err(new HttpServerError(error, actionName, message)),
          summaries: summaries,
        };
      }
      const message = JSON.stringify(error);
      context.logProvider?.error(
        getLocalizedString(logMessageKeys.failExecuteDriver, actionName, message)
      );

      return {
        result: err(assembleError(error as Error, actionName)),
        summaries: summaries,
      };
    }
  }

  private validateArgs(args: CreateSPEContainerTypeArgs): void {
    const invalidParameters: string[] = [];
    if (typeof args.name !== "string" || !args.name) {
      invalidParameters.push("name");
    }

    if (typeof args.owningApplicationId !== "string" || !args.owningApplicationId) {
      invalidParameters.push("owningApplicationId");
    }

    if (
      (args.billingClassification && typeof args.billingClassification !== "string") ||
      !Object.values(SPContainerTypeBillingClassification).includes(args.billingClassification)
    ) {
      invalidParameters.push("billingClassification");
    }

    if (typeof args.discoverable !== "boolean") {
      invalidParameters.push("discoverable");
    }

    if (invalidParameters.length > 0) {
      throw new InvalidActionInputError(actionName, invalidParameters);
    }
  }
}
