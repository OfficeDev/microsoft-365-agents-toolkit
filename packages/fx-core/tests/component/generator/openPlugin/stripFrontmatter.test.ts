// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { expect } from "chai";
import { stripDisallowedFrontmatter } from "../../../../src/component/generator/openPlugin/importer";

describe("stripDisallowedFrontmatter", () => {
  it("removes Claude-specific keys and preserves allowed ones + body", () => {
    const src = `---
name: memory-management
description: A skill.
user-invocable: false
argument-hint: "[--flag]"
---

# Body

Hello.
`;
    const { content, removedKeys } = stripDisallowedFrontmatter(src);
    expect(removedKeys).to.deep.equal(["user-invocable", "argument-hint"]);
    expect(content).to.include("name: memory-management");
    expect(content).to.include("description: A skill.");
    expect(content).to.not.include("user-invocable");
    expect(content).to.not.include("argument-hint");
    expect(content).to.include("# Body");
    expect(content).to.include("Hello.");
  });

  it("returns input unchanged when there is no frontmatter", () => {
    const src = "# Just markdown\n\nNo frontmatter here.\n";
    const { content, removedKeys } = stripDisallowedFrontmatter(src);
    expect(removedKeys).to.deep.equal([]);
    expect(content).to.equal(src);
  });

  it("returns input unchanged when every key is allowed", () => {
    const src = `---
name: ok
description: ok
---
body
`;
    const { content, removedKeys } = stripDisallowedFrontmatter(src);
    expect(removedKeys).to.deep.equal([]);
    expect(content).to.equal(src);
  });

  it("preserves nested objects under an allowed key (metadata)", () => {
    const src = `---
name: ok
metadata:
  version: 1
  tags:
    - a
    - b
forbidden: drop
---
body
`;
    const { content, removedKeys } = stripDisallowedFrontmatter(src);
    expect(removedKeys).to.deep.equal(["forbidden"]);
    expect(content).to.include("metadata:");
    expect(content).to.include("version: 1");
    expect(content).to.include("- a");
    expect(content).to.not.include("forbidden");
  });
});
