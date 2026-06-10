import { assert } from "chai";
import "mocha";
import { createSandbox } from "sinon";
import {
  DeclarativeAgentManifestWrapper,
  AgentSkillElement,
} from "../../src/wrappers/DeclarativeAgentManifestWrapper";

describe("DeclarativeAgentManifestWrapper - Agent Skills", () => {
  const sandbox = createSandbox();

  afterEach(() => {
    sandbox.restore();
  });

  function createAgent(): DeclarativeAgentManifestWrapper {
    return DeclarativeAgentManifestWrapper.create({
      version: "v1.6",
      name: "Test Agent",
      description: "A test agent",
    });
  }

  describe("addSkill", () => {
    it("should add a skill with folder only", () => {
      const agent = createAgent();

      agent.addSkill("skills/my-skill");

      assert.equal(agent.skills.length, 1);
      assert.equal(agent.skills[0].folder, "skills/my-skill");
    });

    it("should prevent duplicate folder paths", () => {
      const agent = createAgent();

      agent.addSkill("skills/my-skill");
      agent.addSkill("skills/my-skill");

      assert.equal(agent.skills.length, 1);
    });

    it("should respect max 10 limit", () => {
      const agent = createAgent();

      for (let i = 0; i < 10; i++) {
        agent.addSkill(`skills/skill-${i}`);
      }
      assert.equal(agent.skills.length, 10);

      // 11th skill should be ignored
      agent.addSkill("skills/skill-overflow");

      assert.equal(agent.skills.length, 10);
      assert.isFalse(agent.hasSkill("skills/skill-overflow"));
    });

    it("should return this for fluent chaining", () => {
      const agent = createAgent();

      const result = agent.addSkill("skills/a").addSkill("skills/b");

      assert.strictEqual(result, agent);
      assert.equal(agent.skills.length, 2);
    });
  });

  describe("removeSkill", () => {
    it("should remove an existing skill by folder", () => {
      const agent = createAgent();

      agent.addSkill("skills/a").addSkill("skills/b");
      agent.removeSkill("skills/a");

      assert.isFalse(agent.hasSkill("skills/a"));
      assert.isTrue(agent.hasSkill("skills/b"));
      assert.equal(agent.skills.length, 1);
    });

    it("should be a no-op when skill does not exist", () => {
      const agent = createAgent();

      agent.addSkill("skills/a");
      agent.removeSkill("skills/nonexistent");

      assert.equal(agent.skills.length, 1);
      assert.isTrue(agent.hasSkill("skills/a"));
    });
  });

  describe("hasSkill", () => {
    it("should return true for existing folder", () => {
      const agent = createAgent();

      agent.addSkill("skills/exists");

      assert.isTrue(agent.hasSkill("skills/exists"));
    });

    it("should return false for non-existing folder", () => {
      const agent = createAgent();

      assert.isFalse(agent.hasSkill("skills/nope"));
    });
  });

  describe("getSkill", () => {
    it("should return skill for existing folder", () => {
      const agent = createAgent();

      agent.addSkill("skills/my-skill");

      const skill = agent.getSkill("skills/my-skill");
      assert.isDefined(skill);
      assert.equal(skill!.folder, "skills/my-skill");
    });

    it("should return undefined for non-existing folder", () => {
      const agent = createAgent();

      const skill = agent.getSkill("skills/nope");
      assert.isUndefined(skill);
    });
  });

  describe("getSkillFolders", () => {
    it("should return empty array when no skills", () => {
      const agent = createAgent();

      assert.deepEqual(agent.getSkillFolders(), []);
    });

    it("should return all folder paths", () => {
      const agent = createAgent();

      agent.addSkill("skills/a").addSkill("skills/b").addSkill("skills/c");

      assert.deepEqual(agent.getSkillFolders(), ["skills/a", "skills/b", "skills/c"]);
    });
  });

  describe("skills getter", () => {
    it("should read from agent_skills property", () => {
      const json = JSON.stringify({
        version: "v1.6",
        name: "Agent",
        description: "Test",
        agent_skills: [{ folder: "skills/from-property" }],
      });

      const agent = DeclarativeAgentManifestWrapper.fromJSON(json);

      assert.equal(agent.skills.length, 1);
      assert.equal(agent.skills[0].folder, "skills/from-property");
    });

    it("should return empty array when neither property exists", () => {
      const agent = createAgent();

      assert.deepEqual(agent.skills, []);
      assert.equal(agent.skills.length, 0);
    });
  });

  describe("round-trip", () => {
    it("should create from JSON with agent_skills, add skill, and include all in toJSON", () => {
      const json = JSON.stringify({
        version: "v1.6",
        name: "Agent",
        description: "Test",
        agent_skills: [{ folder: "skills/original" }],
      });

      const agent = DeclarativeAgentManifestWrapper.fromJSON(json);

      assert.equal(agent.skills.length, 1);
      assert.equal(agent.skills[0].folder, "skills/original");

      agent.addSkill("skills/added");

      assert.equal(agent.skills.length, 2);
      assert.isTrue(agent.hasSkill("skills/original"));
      assert.isTrue(agent.hasSkill("skills/added"));

      const output = JSON.parse(agent.toJSON());
      const outputSkills = output["agent_skills"] as AgentSkillElement[];
      assert.isDefined(outputSkills);
      assert.equal(outputSkills.length, 2);
      assert.equal(outputSkills[0].folder, "skills/original");
      assert.equal(outputSkills[1].folder, "skills/added");
    });
  });

  describe("isDirty", () => {
    it("should mark manifest as dirty when addSkill is called", () => {
      const agent = createAgent();

      assert.isFalse(agent.isDirty);

      agent.addSkill("skills/new-skill");

      assert.isTrue(agent.isDirty);
    });

    it("should mark manifest as dirty when removeSkill is called", () => {
      const agent = createAgent();
      agent.addSkill("skills/to-remove");

      // Reset dirty state by cloning
      const fresh = agent.clone();
      assert.isFalse(fresh.isDirty);

      fresh.removeSkill("skills/to-remove");

      assert.isTrue(fresh.isDirty);
    });
  });
});
