// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as vscode from "vscode";
import { globalStateGet, globalStateUpdate } from "@microsoft/teamsfx-core";
import { isValidProject } from "@microsoft/teamsfx-core";
import { ExtTelemetry } from "../telemetry/extTelemetry";
import { TelemetryEvent } from "../telemetry/extTelemetryEvents";
import * as globalVariables from "../globalVariables";
import { getDefaultString, localize } from "./localizeUtils";
import * as extensionPackage from "../../package.json";

const SURVEY_URL_V3 = "https://aka.ms/ttk-feedback";

enum ExtensionSurveyStateKeys {
  DoNotShowAgain = "survey/doNotShowAgain",
  RemindMeLater = "survey/remindMeLater",
  DisableSurveyForTime = "survey/disableSurveyForTime",
}

const TIME_TO_DISABLE_SURVEY = 1000 * 60 * 60 * 24 * 7 * 12; // 12 weeks
const TIME_TO_REMIND_ME_LATER = 1000 * 60 * 60 * 24 * 7 * 2; // 2 weeks
const TIME_TO_SHOW_SURVEY = 1000 * 60 * 15; // 15 minutes
const SAMPLE_PERCENTAGE = 25; // 25 percent for public preview
const V3PREVIEW_SAMPLE_PERCENTAGE = 100; // always pop up survey
const V3PREVIEW_TIME_TO_DISABLE_SURVEY = -1; // negtive value stands for never show again
const V3PREVIEW_TIME_TO_REMIND_ME_LATER = 1000 * 60 * 60 * 1; // 1 hour

export class ExtensionSurvey {
  private timeToShowSurvey: number;
  private timeToDisableSurvey: number;
  private timeToRemindMeLater: number;
  private checkSurveyInterval?: NodeJS.Timeout;
  private showSurveyTimeout?: NodeJS.Timeout;
  private needToShow = false;
  private static instance: ExtensionSurvey;

  public static getInstance(): ExtensionSurvey {
    if (!ExtensionSurvey.instance) {
      ExtensionSurvey.instance = new ExtensionSurvey(
        TIME_TO_SHOW_SURVEY,
        V3PREVIEW_SAMPLE_PERCENTAGE,
        V3PREVIEW_TIME_TO_DISABLE_SURVEY,
        V3PREVIEW_TIME_TO_REMIND_ME_LATER
      );
    }

    return ExtensionSurvey.instance;
  }

  private constructor(
    timeToShowSurvey?: number,
    samplePercentage?: number,
    timeToDisableSurvey?: number,
    timeToRemindMeLater?: number
  ) {
    this.timeToShowSurvey = timeToShowSurvey ? timeToShowSurvey : TIME_TO_SHOW_SURVEY;

    const randomSample: number = Math.floor(Math.random() * 100) + 1;
    if (randomSample <= (samplePercentage ? samplePercentage : SAMPLE_PERCENTAGE)) {
      this.needToShow = true;
    }
    this.timeToDisableSurvey = timeToDisableSurvey ? timeToDisableSurvey : TIME_TO_DISABLE_SURVEY;
    this.timeToRemindMeLater = timeToRemindMeLater ? timeToRemindMeLater : TIME_TO_REMIND_ME_LATER;
  }

  public activate(): void {
    if (this.needToShow && !this.checkSurveyInterval) {
      this.checkSurveyInterval = setInterval(async () => {
        if (!(await this.shouldShowBanner())) {
          return;
        }

        if (!this.showSurveyTimeout && isValidProject(globalVariables.workspaceUri?.fsPath)) {
          this.showSurveyTimeout = setTimeout(() => this.showSurvey(), this.timeToShowSurvey);
        }
      }, 2000);
    }
  }

  private async shouldShowBanner(): Promise<boolean> {
    const doNotShowAgain = await globalStateGet(ExtensionSurveyStateKeys.DoNotShowAgain, false);
    if (doNotShowAgain) {
      return false;
    }

    const currentTime = Date.now();
    const remindMeLaterTime = await globalStateGet(ExtensionSurveyStateKeys.RemindMeLater, 0);
    if (remindMeLaterTime > currentTime) {
      return false;
    }

    const disableSurveyForTime = await globalStateGet(
      ExtensionSurveyStateKeys.DisableSurveyForTime,
      0
    );
    if (disableSurveyForTime < 0 || disableSurveyForTime > currentTime) {
      return false;
    }

    return true;
  }

  public async openSurveyLink() {
    const link = SURVEY_URL_V3;
    void vscode.commands.executeCommand(
      "vscode.open",
      vscode.Uri.parse(
        `${link}?o=${encodeURIComponent(process.platform)}&v=${encodeURIComponent(
          extensionPackage.version
        )}`
      )
    );

    const disableSurveyForTime =
      this.timeToDisableSurvey < 0
        ? this.timeToDisableSurvey
        : Date.now() + this.timeToDisableSurvey;
    await globalStateUpdate(ExtensionSurveyStateKeys.DisableSurveyForTime, disableSurveyForTime);
  }

  private async showSurvey(): Promise<void> {
    const take = {
      title: localize("teamstoolkit.survey.takeSurvey.title"),
      run: async (): Promise<void> => {
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Survey, {
          message: getDefaultString("teamstoolkit.survey.takeSurvey.message"),
        });

        await this.openSurveyLink();
      },
    };

    const remind = {
      title: localize("teamstoolkit.survey.remindMeLater.title"),
      run: async (): Promise<void> => {
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Survey, {
          message: getDefaultString("teamstoolkit.survey.remindMeLater.message"),
        });
        const disableSurveyForTime = Date.now() + this.timeToRemindMeLater;
        await globalStateUpdate(ExtensionSurveyStateKeys.RemindMeLater, disableSurveyForTime);
      },
    };

    const never = {
      title: localize("teamstoolkit.survey.dontShowAgain.title"),
      run: async (): Promise<void> => {
        ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Survey, {
          message: getDefaultString("teamstoolkit.survey.dontShowAgain.message"),
        });
        await globalStateUpdate(ExtensionSurveyStateKeys.DoNotShowAgain, true);
      },
    };

    const selection = await vscode.window.showInformationMessage(
      localize("teamstoolkit.survey.banner.title"),
      take,
      remind,
      never
    );

    if (this.showSurveyTimeout) {
      clearTimeout(this.showSurveyTimeout);
      this.showSurveyTimeout = undefined;
    }

    if (selection) {
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Survey, {
        message: getDefaultString("teamstoolkit.survey.banner.message"),
      });
      await selection.run();
    } else {
      ExtTelemetry.sendTelemetryEvent(TelemetryEvent.Survey, {
        message: getDefaultString("teamstoolkit.survey.cancelMessage"),
      });
      const disableSurveyForTime = Date.now() + this.timeToRemindMeLater;
      await globalStateUpdate(ExtensionSurveyStateKeys.RemindMeLater, disableSurveyForTime);
    }

    this.timeToShowSurvey = V3PREVIEW_TIME_TO_REMIND_ME_LATER;
  }
}
