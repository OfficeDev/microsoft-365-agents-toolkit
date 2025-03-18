// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { hooks } from "@feathersjs/hooks/lib";
import { FxError, Result, err, ok } from "@microsoft/teamsfx-api";
import { Service } from "typedi";
import axios from "axios";
import { DriverContext } from "../interface/commonArgs";
import { ExecutionResult, StepDriver } from "../interface/stepDriver";
import { addStartAndEndTelemetry } from "../middleware/addStartAndEndTelemetry";
import { WrapDriverContext } from "../util/wrapUtil";
import { CreateDevChannelArgs } from "./interfaces/CreateDevChannelArgs";
import { getLocalizedString } from "../../../common/localizeUtils";
import { GraphClient } from "../../../client/graphClient";
import { HttpClientError } from "../../../error/common";

const actionName = "devChannel/create";

@Service(actionName)
export class CreateDevChannelDriver implements StepDriver {
  description = getLocalizedString("driver.devChannel.description");
  readonly progressTitle = getLocalizedString("driver.devChannel.progress.message");

  public async execute(
    args: CreateDevChannelArgs,
    context: DriverContext,
    outputEnvVarNames: Map<string, string>
  ): Promise<ExecutionResult> {
    const wrapContext = new WrapDriverContext(context, actionName, actionName);
    const res = await this.create(args, wrapContext, outputEnvVarNames);
    return {
      result: res,
      summaries: wrapContext.summaries,
    };
  }

  @hooks([addStartAndEndTelemetry(actionName, actionName)])
  async create(
    args: CreateDevChannelArgs,
    context: WrapDriverContext,
    outputEnvVarNames: Map<string, string>
  ): Promise<Result<Map<string, string>, FxError>> {
    try {
      const graphClient = new GraphClient(context.m365TokenProvider);
      const res = await graphClient.CreateTeamAndChannelAsync(
        args.teamName,
        args.teamDescription,
        args.channelName
      );
      const channelId = res.channelId;
      const teamId = res.teamId;
      context.logProvider.info(
        getLocalizedString(
          "driver.devChannel.success",
          args.teamName,
          args.channelName,
          channelId,
          teamId
        )
      );
      context.addSummary(
        getLocalizedString("driver.devChannel.summary", args.teamName, args.channelName)
      );
      return ok(
        new Map([
          [outputEnvVarNames.get("channelId") as string, channelId],
          [outputEnvVarNames.get("teamId") as string, teamId],
        ])
      );
    } catch (error: any) {
      if (axios.isAxiosError(error)) {
        const message = JSON.stringify(error.response!.data);
        context.logProvider.error(message);
        return err(new HttpClientError(error, actionName, message));
      } else {
        return err(error);
      }
    }
  }
}
