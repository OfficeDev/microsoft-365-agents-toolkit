// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { CancellationToken, ChatResponseStream, LanguageModelChatMessage } from "vscode";
import { ISkill } from "./iSkill";
import { Spec } from "./spec";
import { ExecutionResultEnum } from "./executionResultEnum";

export class SkillSet implements ISkill {
  name: string | undefined;
  capability: string | undefined;
  skills: ISkill[] | undefined;
  retriableTimes: number;

  constructor(skills: ISkill[], retriableTimes?: number) {
    this.name = "skillSet";
    this.capability = "A container for muultiple skills";
    this.skills = skills;
    this.retriableTimes = retriableTimes ?? 1;
  }

  public canInvoke(spec: Spec): boolean {
    if (!this.skills) {
      return false;
    }
    return true;
  }

  // eslint-disable-next-line @typescript-eslint/require-await
  public async invoke(
    languageModel: LanguageModelChatMessage,
    response: ChatResponseStream,
    token: CancellationToken,
    spec: Spec
  ): Promise<{ result: ExecutionResultEnum; spec: Spec }> {
    if (!this.skills) {
      return { result: ExecutionResultEnum.Success, spec: spec };
    }
    let specCopy = new Spec("");
    specCopy.clone(spec);
    let retried = 0;
    let isSuccessed = true;
    let isFailedAndGoNext = false;
    while (retried < this.retriableTimes) {
      retried++;

      for (const skill of this.skills) {
        if (!skill.canInvoke(specCopy)) {
          isSuccessed = false;
          continue;
        }
        const { result: result, spec: newSpec }: { result: ExecutionResultEnum; spec: Spec } =
          await skill.invoke(languageModel, response, token, specCopy);
        if (result === ExecutionResultEnum.Rejected) {
          // We want to keep the telemetry data anyway
          return { result: result, spec: newSpec };
        }
        if (result === ExecutionResultEnum.Failure) {
          isSuccessed = false;
        }
        if (result === ExecutionResultEnum.FailedAndGoNext) {
          isSuccessed = false;
          isFailedAndGoNext = true;
        }
        if (result === ExecutionResultEnum.Success) {
          isSuccessed = true;
        }
        specCopy = newSpec;
      }

      if (isSuccessed) {
        return { result: ExecutionResultEnum.Success, spec: specCopy };
      }
    }
    if (isFailedAndGoNext) {
      return { result: ExecutionResultEnum.FailedAndGoNext, spec: specCopy };
    } else {
      return { result: ExecutionResultEnum.Failure, spec: specCopy };
    }
  }
}
