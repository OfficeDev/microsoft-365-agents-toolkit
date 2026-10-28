// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import "mocha";
import {
  AppManifestUtils,
  DeclarativeAgentManifestConverter,
  DeclarativeAgentManifestWrapper,
} from "../src";

const schemaUrl =
  "https://developer.microsoft.com/json-schemas/copilot/declarative-agent/v1.9/schema.json";

function skillManifest(count: number) {
  return {
    $schema: schemaUrl,
    version: "v1.9",
    name: "Skill agent",
    description: "An agent with skills",
    agent_skills: Array.from({ length: count }, (_, index) => ({
      folder: `skills/skill-${index}`,
    })),
  };
}

describe("Declarative agent v1.9", () => {
  it("SCN-CREATE-SKILL-09: round-trips skills through the registered converter", async () => {
    const input = skillManifest(8);
    const manifest = DeclarativeAgentManifestConverter.jsonToManifest(JSON.stringify(input));
    const output = DeclarativeAgentManifestConverter.manifestToJson(manifest);

    assert.deepEqual(JSON.parse(output), input);
    assert.deepEqual(await AppManifestUtils.validateAgainstSchema(manifest), []);
  });

  it("SCN-CREATE-SKILL-09: rejects malformed skill entries without unchecked fallback", () => {
    for (const agentSkills of [{ folder: "skills/test" }, [{}], [{ folder: 42 }]]) {
      assert.throws(() =>
        DeclarativeAgentManifestConverter.jsonToManifest(
          JSON.stringify({ ...skillManifest(1), agent_skills: agentSkills })
        )
      );
    }
  });

  it("SCN-CREATE-SKILL-09: rejects a ninth skill during schema validation", async () => {
    const manifest = DeclarativeAgentManifestConverter.jsonToManifest(
      JSON.stringify(skillManifest(9))
    );
    const errors = await AppManifestUtils.validateAgainstSchema(manifest);

    assert.isTrue(errors.some((error) => error.includes("/agent_skills")));
  });

  it("SCN-CREATE-SKILL-09: rejects an empty skill folder during schema validation", async () => {
    const manifest = DeclarativeAgentManifestConverter.jsonToManifest(
      JSON.stringify({ ...skillManifest(1), agent_skills: [{ folder: "" }] })
    );
    const errors = await AppManifestUtils.validateAgainstSchema(manifest);

    assert.isTrue(errors.some((error) => error.includes("/agent_skills/0/folder")));
  });

  it("SCN-CREATE-SKILL-10: the v1.9 wrapper round-trips skills", () => {
    const agent = DeclarativeAgentManifestWrapper.fromJSON(JSON.stringify(skillManifest(1)));
    agent.addSkill("skills/another");
    const restored = DeclarativeAgentManifestWrapper.fromJSON(agent.toJSON());

    assert.equal(restored.version, "v1.9");
    assert.deepEqual(restored.getSkillFolders(), ["skills/skill-0", "skills/another"]);
  });
});
