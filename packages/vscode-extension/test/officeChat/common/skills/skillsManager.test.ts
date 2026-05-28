import * as chai from "chai";
import sinon from "ts-sinon";

import { SkillsManager } from "../../../../src/officeChat/common/skills/skillsManager";
import { OfficeChatCommand } from "../../../../src/officeChat/consts";

describe("skillsManager", () => {
  const sandbox = sinon.createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  it("getInstance create instance", () => {
    const skillsManager = SkillsManager.getInstance();

    chai.assert.isNotNull(skillsManager);
  });

  it("getInstance return same instance", () => {
    const skillsManager1 = SkillsManager.getInstance();
    const skillsManager2 = SkillsManager.getInstance();

    chai.assert.equal(skillsManager1, skillsManager2);
  });

  it("getCapableSkills GenerateCode", () => {
    const skillsManager = SkillsManager.getInstance();
    const skills = skillsManager.getCapableSkills(OfficeChatCommand.GenerateCode);
    chai.expect(skills).to.have.lengthOf(2);
  });

  it("getCapableSkills Create", () => {
    const skillsManager = SkillsManager.getInstance();
    const skills = skillsManager.getCapableSkills(OfficeChatCommand.Create);
    chai.expect(skills).to.have.lengthOf(3);
  });

  it("getCapableSkills other commands", () => {
    const skillsManager = SkillsManager.getInstance();
    const skills = skillsManager.getCapableSkills("other" as OfficeChatCommand);
    chai.expect(skills).to.have.lengthOf(0);
  });
});
