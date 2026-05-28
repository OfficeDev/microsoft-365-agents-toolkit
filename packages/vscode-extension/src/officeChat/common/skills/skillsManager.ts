// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { OfficeChatCommand } from "../../consts";
import { Explainer } from "./codeExplainer";
import { CodeGenerator } from "./codeGenerator";
import { CodeIssueCorrector } from "./codeIssueCorrector";
import { ISkill } from "./iSkill"; // Replace this import statement
import { Printer } from "./printer";
import { projectCreator } from "./projectCreator";
import { SkillSet } from "./skillset";

export class SkillsManager {
  private static instance: SkillsManager;
  private projectCreator: ISkill;
  private codeGenerator: ISkill;
  private codeExplainer: ISkill;
  private printer: ISkill;
  private codeIssueCorrector: ISkill;

  private constructor() {
    // Private constructor to prevent direct instantiation
    this.codeGenerator = new CodeGenerator();
    this.printer = new Printer();
    this.codeExplainer = new Explainer();
    this.projectCreator = new projectCreator();
    this.codeIssueCorrector = new CodeIssueCorrector();
  }

  public static getInstance(): SkillsManager {
    if (!SkillsManager.instance) {
      SkillsManager.instance = new SkillsManager();
    }
    return SkillsManager.instance;
  }

  public getCapableSkills(capability: OfficeChatCommand): ISkill[] {
    const capableSkills: ISkill[] = [];
    switch (capability) {
      case OfficeChatCommand.GenerateCode:
        capableSkills.push(new SkillSet([this.codeGenerator, this.codeIssueCorrector], 2));
        capableSkills.push(this.printer);
        break;
      case OfficeChatCommand.Create:
        capableSkills.push(new SkillSet([this.codeGenerator, this.codeIssueCorrector], 2));
        capableSkills.push(this.printer);
        capableSkills.push(this.projectCreator);
        break;
      default:
        break;
    }

    return capableSkills;
  }
}
