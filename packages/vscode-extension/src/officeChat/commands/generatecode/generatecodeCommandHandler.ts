// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import {
  CancellationToken,
  ChatContext,
  ChatFollowup,
  ChatRequest,
  ChatResponseStream,
  LanguageModelChatMessage,
  LanguageModelChatMessageRole,
} from "vscode";
import { ExtTelemetry } from "../../../telemetry/extTelemetry";
import { TelemetryEvent } from "../../../telemetry/extTelemetryEvents";
import { localize } from "../../../utils/localizeUtils";
import { OfficeChatCommand, officeChatParticipantId } from "../../consts";
import { Planner } from "../../common/planner";
import { isInputHarmful } from "../../utils";
import { ICopilotChatOfficeResult } from "../../types";
import { OfficeChatTelemetryBlockReasonEnum, OfficeChatTelemetryData } from "../../telemetry";
import followupProvider from "../../../chat/followupProvider";

export default async function generatecodeCommandHandler(
  request: ChatRequest,
  context: ChatContext,
  response: ChatResponseStream,
  token: CancellationToken
): Promise<ICopilotChatOfficeResult> {
  const officeChatTelemetryData = OfficeChatTelemetryData.createByParticipant(
    officeChatParticipantId,
    OfficeChatCommand.GenerateCode
  );
  ExtTelemetry.sendTelemetryEvent(
    TelemetryEvent.CopilotChatStart,
    officeChatTelemetryData.properties
  );

  if (request.prompt.trim() === "") {
    officeChatTelemetryData.setTimeToFirstToken();
    response.markdown(
      localize("teamstoolkit.chatParticipants.officeAddIn.generateCode.noPromptAnswer")
    );
    const followUps: ChatFollowup[] = [
      {
        label: "@office /generatecode create a chart based on the selected range in Excel",
        command: "generatecode",
        prompt: "create a chart based on the selected range in Excel",
      },
      {
        label: "@office /generatecode insert a content control in a Word document",
        command: "generatecode",
        prompt: "insert a content control in a Word document",
      },
    ];
    followupProvider.addFollowups(followUps);
    officeChatTelemetryData.setBlockReason(OfficeChatTelemetryBlockReasonEnum.UnsupportedInput);
    officeChatTelemetryData.markComplete("fail");
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChat,
      officeChatTelemetryData.properties,
      officeChatTelemetryData.measurements
    );
    return {
      metadata: {
        command: OfficeChatCommand.GenerateCode,
        requestId: officeChatTelemetryData.requestId,
      },
    };
  }

  const isHarmful = await isInputHarmful(request, token, officeChatTelemetryData);
  if (!isHarmful) {
    let chatResult: ICopilotChatOfficeResult = {};
    try {
      chatResult = await Planner.getInstance().processRequest(
        new LanguageModelChatMessage(LanguageModelChatMessageRole.User, request.prompt),
        request,
        response,
        token,
        OfficeChatCommand.GenerateCode,
        officeChatTelemetryData
      );
      officeChatTelemetryData.markComplete();
    } catch (error) {
      officeChatTelemetryData.markComplete("fail");
    }
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChat,
      officeChatTelemetryData.properties,
      officeChatTelemetryData.measurements
    );
    return chatResult;
  } else {
    officeChatTelemetryData.setTimeToFirstToken();
    response.markdown(localize("teamstoolkit.chatParticipants.officeAddIn.harmfulInputResponse"));
    officeChatTelemetryData.setBlockReason(OfficeChatTelemetryBlockReasonEnum.RAI);
    officeChatTelemetryData.markComplete("fail");
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChat,
      officeChatTelemetryData.properties,
      officeChatTelemetryData.measurements
    );
    return {
      metadata: {
        command: OfficeChatCommand.GenerateCode,
        requestId: officeChatTelemetryData.requestId,
      },
    };
  }
}
