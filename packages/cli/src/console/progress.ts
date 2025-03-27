// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { IProgressHandler } from "@microsoft/teamsfx-api";
import figures from "figures";
import { TextType, colorize } from "../colorize";
import ScreenManager, { Row } from "./screen";

export default class Progress implements IProgressHandler {
  private static instances: Progress[] = [];
  private static rows: Row[] = [];
  private static finishedRows: Row[] = [];

  static readonly barSize = 20;

  private static add(instance: Progress) {
    this.instances.push(instance);
    this.rows.push(ScreenManager.addProgress(instance.wholeMessage.bind(instance)));
  }

  private static finish(instance: Progress, hideAfterFinish = false) {
    const idx = this.findIndex(instance);
    if (idx > -1) {
      this.rows[idx].update();
      this.rows[idx].removeCB();
      if (!hideAfterFinish) this.finishedRows.push(this.rows[idx]);
      else {
        ScreenManager.delete(this.rows[idx]);
      }
      this.instances.splice(idx, 1);
      this.rows.splice(idx, 1);
    }

    if (this.instances.length === 0 && this.finishedRows.length > 0) {
      for (const row of this.finishedRows) {
        row.freeze();
      }
      this.finishedRows = [];
    }
  }

  static end(success: boolean) {
    const instances = [...this.instances];
    instances.forEach((i) => i.end(success));
  }

  private static findIndex(instance: Progress) {
    return this.instances.findIndex((i) => i === instance);
  }

  private readonly title: string;
  private totalSteps: number;
  private currentStep: number;
  private currentPercentage: number;
  private detail?: string;
  private status?: "done" | "error" | "running";

  constructor(title: string, totalSteps: number) {
    this.totalSteps = totalSteps;
    this.title = title;
    this.currentStep = 0;
    this.currentPercentage = 0;
  }

  start(detail?: string) {
    this.status = "running";
    this.detail = detail;
    this.currentStep = 0;
    if (Progress.findIndex(this) < 0) {
      Progress.add(this);
    }
  }

  end(success: boolean, hideAfterFinish = false) {
    this.status = success ? "done" : "error";
    if (success) this.currentPercentage = 100;
    if (Progress.findIndex(this) > -1) {
      Progress.finish(this, hideAfterFinish);
    }
  }

  next(detail?: string) {
    this.detail = detail;
    this.currentStep++;
    if (this.totalSteps < this.currentStep) this.totalSteps = this.currentStep;
  }
  text(detail?: string) {
    this.detail = detail;
  }
  private updatePercentage() {
    const needArrivedPercentage = ((this.currentStep - 1) / this.totalSteps) * 100;
    const nextArrivedPercentage = (this.currentStep / this.totalSteps) * 100 - 1;
    if (this.currentPercentage < needArrivedPercentage) {
      const diff = needArrivedPercentage - this.currentPercentage;
      this.currentPercentage += diff / ScreenManager.fps >= 5 ? diff / ScreenManager.fps : 5;
    } else if (this.currentPercentage < nextArrivedPercentage) {
      const diff = nextArrivedPercentage - this.currentPercentage;
      this.currentPercentage += diff / ScreenManager.fps / 20;
    }
    this.currentPercentage = Math.min(this.currentPercentage, 100);
  }

  wholeMessage(): string {
    this.updatePercentage();
    const message =
      this.status === "done"
        ? this.doneMessage
        : this.status === "error"
        ? this.errorMessage
        : this.message;
    return colorize(
      `${this.barStatus}  ${Math.round(this.currentPercentage)}% ${this.runningChar} ${message}`,
      TextType.Info
    );
  }

  get barStatus(): string {
    const completeSize = Math.round((this.currentPercentage / 100) * Progress.barSize);
    return "█".repeat(completeSize) + "▒".repeat(Progress.barSize - completeSize);
  }

  get runningChar() {
    const chars = ["|", "/", "-", "\\"];
    return chars[this.status === "running" ? Math.floor(Date.now() / 1000) % 4 : 0];
  }

  get doneMessage(): string {
    return (
      colorize(`[${this.totalSteps}/${this.totalSteps}] ${this.title} `, TextType.Info) +
      colorize(` (${figures.tick}) Done.`, TextType.Success)
    );
  }

  get errorMessage(): string {
    return (
      colorize(
        `[${this.currentStep}/${this.totalSteps}] ${this.title}: ${this.detail || "starting."}`,
        TextType.Info
      ) + colorize(` (${figures.cross}) Failed.`, TextType.Error)
    );
  }

  get message(): string {
    return colorize(
      `[${this.currentStep}/${this.totalSteps}] ${this.title}: ${this.detail || "starting."}`,
      TextType.Info
    );
  }
}
