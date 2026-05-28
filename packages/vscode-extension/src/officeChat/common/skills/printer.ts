// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CancellationToken, ChatResponseStream, LanguageModelChatMessage } from "vscode";
import { ISkill } from "./iSkill";
import { Spec } from "./spec";
import { ExecutionResultEnum } from "./executionResultEnum";
import { isOutputHarmful } from "../../utils";
import { localize } from "../../../utils/localizeUtils";
import { TelemetryProperty } from "../../../telemetry/extTelemetryEvents";

export class Printer implements ISkill {
  name: string | undefined;
  capability: string | undefined;

  constructor() {
    this.name = "printer";
    this.capability = "Print the output in a readable format to user";
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
    const template = `
${localize("teamstoolkit.chatParticipants.officeAddIn.printer.outputTemplate.codeIntro")}\n
\`\`\`typescript
${spec.appendix.codeSnippet}
\`\`\`

${localize("teamstoolkit.chatParticipants.officeAddIn.printer.outputTemplate.ending")}\n
`;
    const isHarmful = await isOutputHarmful(template, token, spec);
    if (isHarmful) {
      response.markdown(localize("teamstoolkit.chatParticipants.officeAddIn.printer.raiBlock"));
      spec.appendix.telemetryData.isHarmful = true;
      return { result: ExecutionResultEnum.Failure, spec: spec };
    } else {
      response.markdown(template);
      spec.appendix.telemetryData.properties[TelemetryProperty.CopilotChatHasCodeBlock] = "true";
      return { result: ExecutionResultEnum.Success, spec: spec };
    }
  }
}
