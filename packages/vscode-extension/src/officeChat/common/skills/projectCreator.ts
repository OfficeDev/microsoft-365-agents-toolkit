// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ChatResponseStream, LanguageModelChatMessage, CancellationToken } from "vscode";
import { ISkill } from "./iSkill";
import { Spec } from "./spec";
import { ExecutionResultEnum } from "./executionResultEnum";
import { CHAT_CREATE_OFFICE_PROJECT_COMMAND_ID } from "../../consts";
import { localize } from "../../../utils/localizeUtils";
import { showOfficeTemplateFileTree } from "../../commands/create/helper";

export class projectCreator implements ISkill {
  name: string | undefined;
  capability: string | undefined;

  constructor() {
    this.name = "Project Creator";
    this.capability = "Create a new project template";
  }

  public canInvoke(spec: Spec): boolean {
    return (
      !!spec.userInput &&
      !!spec.appendix.codeSnippet &&
      !!spec.appendix.codeTaskBreakdown &&
      spec.appendix.codeTaskBreakdown.length > 0
    );
  }

  // eslint-disable-next-line @typescript-eslint/require-await
  public async invoke(
    languageModel: LanguageModelChatMessage,
    response: ChatResponseStream,
    token: CancellationToken,
    spec: Spec
  ): Promise<{ result: ExecutionResultEnum; spec: Spec }> {
    const host = spec.appendix.host.toLowerCase();
    const capabilities = spec.appendix.isCustomFunction
      ? "excel-custom-functions-shared"
      : `${host}-taskpane`;
    const createInputs = {
      capabilities: capabilities,
      "project-type": "office-xml-addin-type",
      "addin-host": host,
      "programming-language": "typescript",
      agent: "office",
    };
    const rootFolder = await showOfficeTemplateFileTree(
      createInputs,
      response,
      spec.appendix.codeSnippet
    );
    const sampleTitle = localize("teamstoolkit.chatParticipants.officeAddIn.create.project");
    response.button({
      command: CHAT_CREATE_OFFICE_PROJECT_COMMAND_ID,
      arguments: [
        rootFolder,
        spec.appendix.telemetryData.requestId,
        "No Match Result Type",
        capabilities,
      ],
      title: sampleTitle,
    });
    return { result: ExecutionResultEnum.Success, spec: spec };
  }
}
