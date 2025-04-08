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
import { loadStateFromEnv, mapStateToEnv } from "../util/utils";
import { TelemetryProperty } from "../../../common/telemetry";
import { CreateDevChannelOutput } from "./interfaces/CreateDevChannelOutput";

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
    // Skip creation if the team and channel already exist
    const state: CreateDevChannelOutput = loadStateFromEnv(outputEnvVarNames);
    const graphClient = new GraphClient(context.m365TokenProvider);

    try {
      if (state.teamId && state.channelId) {
        const message = getLocalizedString(
          "driver.devChannel.summary.exists",
          outputEnvVarNames.get("teamId"),
          outputEnvVarNames.get("channelId")
        );
        context.logProvider.info(message);
        context.addSummary(message);
        context.addTelemetryProperties({ [TelemetryProperty.SkipCreation]: "true" });

        // Update the channel web URL, incase user manually modify teamId or channelId
        state.channelWebUrl = await this.getChannelWebUrl(
          graphClient,
          state.teamId,
          state.channelId
        );
        const outputs = mapStateToEnv(state, outputEnvVarNames);
        return ok(outputs);
      }

      const res = await graphClient.CreateTeamAndChannelAsync(
        args.teamName,
        args.teamDescription,
        args.channelName
      );
      state.channelId = res.channelId;
      state.teamId = res.teamId;
      context.logProvider.info(
        getLocalizedString(
          "driver.devChannel.success",
          args.teamName,
          args.channelName,
          state.channelId,
          state.teamId
        )
      );
      context.addSummary(
        getLocalizedString("driver.devChannel.summary", args.teamName, args.channelName)
      );
      state.channelWebUrl = await this.getChannelWebUrl(graphClient, state.teamId, state.channelId);
      const outputs = mapStateToEnv(state, outputEnvVarNames);
      return ok(outputs);
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

  private async getChannelWebUrl(
    graphClient: GraphClient,
    channelId: string,
    teamId: string
  ): Promise<string> {
    const res = await graphClient.GetChannelDeeplinkAsync(teamId, channelId);
    return res;
  }
}
