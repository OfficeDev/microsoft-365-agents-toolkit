// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  CancellationToken,
  ChatRequest,
  ChatResponseStream,
  LanguageModelChatMessage,
} from "vscode";
import { OfficeChatCommand } from "../consts";
import { ISkill } from "./skills/iSkill";
import { SkillsManager } from "./skills/skillsManager";
import { Spec } from "./skills/spec";
import { ICopilotChatOfficeResult } from "../types";
import { TelemetryEvent } from "../../telemetry/extTelemetryEvents";
import { ExtTelemetry } from "../../telemetry/extTelemetry";
import { ExecutionResultEnum } from "./skills/executionResultEnum";
import {
  MeasurementCodeGenExecutionTimeInTotalSec,
  MeasurementCodeGenGetSampleTimeInTotalSec,
  MeasurementCodeGenPreScanTimeInTotalSec,
  MeasurementCodeGenTaskBreakdownTimeInTotalSec,
  MeasurementCommandExcutionTimeSec,
  MeasurementErrorsAfterCorrection,
  MeasurementSelfReflectionExecutionTimeInTotalSec,
  PropertySystemFailureFromSkill,
  PropertySystemRequesRejected,
  PropertySystemRequestCancelled,
  PropertySystemRequestFailed,
  PropertySystemRequestFailedAndGoNext,
  PropertySystemRequestSucceeded,
} from "./telemetryConsts";
import { purifyUserMessage } from "../utils";
import { localize } from "../../utils/localizeUtils";
import { OfficeChatTelemetryBlockReasonEnum, OfficeChatTelemetryData } from "../telemetry";

export class Planner {
  private static instance: Planner;

  private constructor() {
    // Private constructor to prevent direct instantiation
  }

  public static getInstance(): Planner {
    if (!Planner.instance) {
      Planner.instance = new Planner();
    }
    return Planner.instance;
  }

  public async processRequest(
    languageModel: LanguageModelChatMessage,
    request: ChatRequest,
    response: ChatResponseStream,
    token: CancellationToken,
    command: OfficeChatCommand,
    telemetryData: OfficeChatTelemetryData
  ): Promise<ICopilotChatOfficeResult> {
    const candidates: ISkill[] = SkillsManager.getInstance().getCapableSkills(command);
    const t0 = performance.now();
    token.onCancellationRequested(() => {
      const t1 = performance.now();
      const duration = (t1 - t0) / 1000;
      telemetryData.extendBy(
        { [PropertySystemRequestCancelled]: "true" },
        { [MeasurementCommandExcutionTimeSec]: duration }
      );
      telemetryData.markComplete();
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CopilotChat, telemetryData.properties);
    });
    const chatResult: ICopilotChatOfficeResult = {
      metadata: {
        command: command,
        requestId: telemetryData.requestId,
      },
    };

    if (!candidates || candidates.length === 0) {
      chatResult.errorDetails = { message: "No skill is available to process the request." };
      return chatResult;
    }

    // dispatcher
    const purified = await purifyUserMessage(request.prompt, token, telemetryData);
    telemetryData.setTimeToFirstToken();
    response.markdown(`
${localize("teamstoolkit.chatParticipants.officeAddIn.printer.outputTemplate.intro")}\n
${purified}
`);
    const spec = new Spec(purified, telemetryData.requestId);
    try {
      for (let index = 0; index < candidates.length; index++) {
        const candidate = candidates[index];
        if (!candidate.canInvoke(spec)) {
          throw new Error("Internal error: the prior skill failed to produce necessary data.");
        }
        const { result: invokeResult, spec: newSpec }: { result: ExecutionResultEnum; spec: Spec } =
          await candidate.invoke(languageModel, response, token, spec);
        spec.clone(newSpec);
        if (invokeResult == ExecutionResultEnum.Failure) {
          spec.appendix.telemetryData.properties[PropertySystemRequestFailed] = "true";
          spec.appendix.telemetryData.properties[PropertySystemFailureFromSkill] =
            candidate.name || "unknown";
          if (spec.appendix.telemetryData.isHarmful === true) {
            telemetryData.setBlockReason(OfficeChatTelemetryBlockReasonEnum.RAI);
          } else {
            telemetryData.setBlockReason(OfficeChatTelemetryBlockReasonEnum.PlannerFailure);
          }
          throw new Error("Failed to process the request.");
        }

        if (invokeResult == ExecutionResultEnum.Rejected) {
          // hard stop if one of the skill reject to process the request
          // for example, the user ask is not what we target to address
          spec.appendix.telemetryData.properties[PropertySystemRequesRejected] = "true";
          spec.appendix.telemetryData.properties[PropertySystemFailureFromSkill] =
            candidate.name || "unknown";
          telemetryData.setBlockReason(OfficeChatTelemetryBlockReasonEnum.OffTopic);
          throw new Error(
            `The skill "${candidate.name || "Unknown"}" is rejected to process the request.`
          );
        }

        if (invokeResult == ExecutionResultEnum.FailedAndGoNext) {
          spec.appendix.telemetryData.properties[PropertySystemRequestFailedAndGoNext] = "true";
          spec.appendix.telemetryData.properties[PropertySystemFailureFromSkill] =
            candidate.name || "unknown";
        } else {
          spec.appendix.telemetryData.properties[PropertySystemRequestSucceeded] = "true";
        }

        console.log(`Skill ${candidate.name || "unknown"} is executed.`);
      }
    } catch (error) {
      // console.log("Purified user message: ", purified);
      // console.error(error);
      const errorDetails = localize(
        "teamstoolkit.chatParticipants.officeAddIn.default.canNotAssist"
      );
      response.markdown(errorDetails);
      throw new Error("Failed or rejected to process the request.");
    }
    const t1 = performance.now();
    const duration = (t1 - t0) / 1000;
    spec.appendix.telemetryData.measurements[MeasurementCommandExcutionTimeSec] = duration;
    telemetryData.extendBy(
      spec.appendix.telemetryData.properties,
      spec.appendix.telemetryData.measurements
    );
    telemetryData.setHostType(spec.appendix.host.toLowerCase());
    telemetryData.setRelatedSampleName(spec.appendix.telemetryData.relatedSampleName.toString());
    for (const chatMessage of spec.appendix.telemetryData.chatMessages) {
      telemetryData.chatMessages.push(chatMessage);
    }
    for (const responseChatMessage of spec.appendix.telemetryData.responseChatMessages) {
      telemetryData.responseChatMessages.push(responseChatMessage);
    }
    const debugInfo = `
      ## Time cost:\n
      In total ${Math.ceil(duration)} seconds.\n
      - Task pre scan: ${Math.ceil(
        spec.appendix.telemetryData.measurements[MeasurementCodeGenPreScanTimeInTotalSec]
      )} seconds.
      - Task breakdown: ${Math.ceil(
        spec.appendix.telemetryData.measurements[MeasurementCodeGenTaskBreakdownTimeInTotalSec]
      )} seconds.
      - Download sample: ${Math.ceil(
        spec.appendix.telemetryData.measurements[MeasurementCodeGenGetSampleTimeInTotalSec]
      )} seconds.
      - Code gen: ${Math.ceil(
        spec.appendix.telemetryData.measurements[MeasurementCodeGenExecutionTimeInTotalSec]
      )} seconds.
      - Self reflection: ${Math.ceil(
        spec.appendix.telemetryData.measurements[MeasurementSelfReflectionExecutionTimeInTotalSec]
      )} seconds.\n\n
      ## Compile error remains:\n
      ${Math.ceil(spec.appendix.telemetryData.measurements[MeasurementErrorsAfterCorrection])}
      `;
    console.debug(debugInfo);
    // response.markdown(debugInfo);

    return chatResult;
  }
}
