// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  CancellationToken,
  ChatContext,
  ChatRequest,
  ChatResponseStream,
  LanguageModelChatMessage,
  LanguageModelChatMessageRole,
} from "vscode";

import * as util from "util";
import { CommandKey } from "../../../constants";
import { ExtTelemetry } from "../../../telemetry/extTelemetry";
import { TelemetryEvent, TelemetryTriggerFrom } from "../../../telemetry/extTelemetryEvents";
import { localize } from "../../../utils/localizeUtils";
import { CHAT_EXECUTE_COMMAND_ID, TeamsChatCommand, chatParticipantId } from "../../consts";
import { brieflyDescribeProjectSystemPrompt, describeProjectSystemPrompt } from "../../prompts";
import { ChatTelemetryData } from "../../telemetry";
import { ICopilotChatResult } from "../../types";
import { verbatimCopilotInteraction } from "../../utils";
import * as helper from "./helper";

export default async function createCommandHandler(
  request: ChatRequest,
  context: ChatContext,
  response: ChatResponseStream,
  token: CancellationToken
): Promise<ICopilotChatResult> {
  const chatTelemetryData = ChatTelemetryData.createByParticipant(
    chatParticipantId,
    TeamsChatCommand.Create
  );
  ExtTelemetry.sendTelemetryEvent(TelemetryEvent.CopilotChatStart, chatTelemetryData.properties);

  if (request.prompt.trim() === "") {
    response.markdown(localize("teamstoolkit.chatParticipants.create.noPromptAnswer"));

    chatTelemetryData.markComplete();
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChat,
      chatTelemetryData.properties,
      chatTelemetryData.measurements
    );
    return {
      metadata: {
        command: TeamsChatCommand.Create,
        requestId: chatTelemetryData.requestId,
      },
    };
  }

  const matchedResult = await helper.matchProject(request, token, chatTelemetryData);

  if (matchedResult.length === 0) {
    response.markdown(localize("teamstoolkit.chatParticipants.create.noMatched"));
    chatTelemetryData.markComplete();
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChat,
      chatTelemetryData.properties,
      chatTelemetryData.measurements
    );
    return {
      metadata: {
        command: TeamsChatCommand.Create,
        requestId: chatTelemetryData.requestId,
      },
    };
  }
  if (matchedResult.length === 1) {
    response.markdown(localize("teamstoolkit.chatParticipants.create.oneMatched"));
    const firstMatch = matchedResult[0];
    const describeProjectChatMessages = [
      describeProjectSystemPrompt(),
      new LanguageModelChatMessage(
        LanguageModelChatMessageRole.User,
        `The project you are looking for is '${JSON.stringify({
          name: firstMatch.name,
          description: firstMatch.description,
        })}'.`
      ),
    ];
    chatTelemetryData.chatMessages.push(...describeProjectChatMessages);

    await verbatimCopilotInteraction(
      "copilot-gpt-3.5-turbo",
      describeProjectChatMessages,
      response,
      token
    );
    if (firstMatch.type === "sample") {
      const folder = await helper.showFileTree(firstMatch, response);
      const sampleTitle = localize("teamstoolkit.chatParticipants.create.sample");
      response.button({
        command: CommandKey.DownloadSample,
        arguments: [TelemetryTriggerFrom.CopilotChat, firstMatch.id],
        title: sampleTitle,
      });
    } else if (firstMatch.type === "template") {
      const templateTitle = localize("teamstoolkit.chatParticipants.create.template");
      response.button({
        command: CHAT_EXECUTE_COMMAND_ID,
        arguments: [CommandKey.Create, chatTelemetryData.requestId, firstMatch.data],
        title: templateTitle,
      });
    }

    chatTelemetryData.markComplete();
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChat,
      chatTelemetryData.properties,
      chatTelemetryData.measurements
    );
    return {
      metadata: {
        command: TeamsChatCommand.Create,
        requestId: chatTelemetryData.requestId,
      },
    };
  } else if (matchedResult.length <= 5) {
    response.markdown(
      util.format(
        localize("teamstoolkit.chatParticipants.create.multipleMatched"),
        matchedResult.slice(0, 3).length
      )
    );
    for (const project of matchedResult.slice(0, 3)) {
      response.markdown(`- ${project.name}: `);

      const brieflyDescribeProjectChatMessages = [
        brieflyDescribeProjectSystemPrompt(),
        new LanguageModelChatMessage(
          LanguageModelChatMessageRole.User,
          `The project you are looking for is '${JSON.stringify(project)}'.`
        ),
      ];
      chatTelemetryData.chatMessages.push(...brieflyDescribeProjectChatMessages);

      await verbatimCopilotInteraction(
        "copilot-gpt-3.5-turbo",
        brieflyDescribeProjectChatMessages,
        response,
        token
      );
      if (project.type === "sample") {
        const sampleTitle = localize("teamstoolkit.chatParticipants.create.sample");
        response.button({
          command: CommandKey.DownloadSample,
          arguments: [TelemetryTriggerFrom.CopilotChat, project.id],
          title: sampleTitle,
        });
      } else if (project.type === "template") {
        const templateTitle = localize("teamstoolkit.chatParticipants.create.template");
        response.button({
          command: CHAT_EXECUTE_COMMAND_ID,
          arguments: [CommandKey.Create, chatTelemetryData.requestId, project.data],
          title: templateTitle,
        });
      }
    }

    chatTelemetryData.markComplete();
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChat,
      chatTelemetryData.properties,
      chatTelemetryData.measurements
    );
    return {
      metadata: {
        command: TeamsChatCommand.Create,
        requestId: chatTelemetryData.requestId,
      },
    };
  } else {
    response.markdown(localize("teamstoolkit.chatParticipants.create.tooGeneric"));

    chatTelemetryData.markComplete();
    ExtTelemetry.sendTelemetryEvent(
      TelemetryEvent.CopilotChat,
      chatTelemetryData.properties,
      chatTelemetryData.measurements
    );
    return {
      metadata: {
        command: TeamsChatCommand.Create,
        requestId: chatTelemetryData.requestId,
      },
    };
  }
}
