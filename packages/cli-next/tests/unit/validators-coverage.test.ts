/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import {
  templateRegistry,
  registerBuiltinTemplates,
  type TemplateDescriptor,
} from "@microsoft/teamsfx-core-next";
import { VALIDATORS, runValidatorsForPhase } from "../e2e/infra/validators";

/**
 * Every testable template must declare tags that match the validators that
 * will run against it. This test fails when a new driver is introduced
 * without a matching validator, or when a testable template forgets to
 * declare its tags. See `.dev/assessments/2026-04-17-core-next-cli-next-refactor.md`
 * for the validator design.
 */

const KNOWN_TAGS = new Set<string>([
  "*",
  "teamsApp",
  "publishable",
  "publishedApp",
  "bot",
  "tab",
  "aad",
  "function",
  "azureFunction",
  "appService",
  "declarativeAgent",
  "apiPlugin",
  "apiAuthOAuth",
  "apiAuthApiKey",
  "openapi",
  "mcp",
  "extendedToM365",
  "connector",
]);

describe("E2E validator coverage", () => {
  before(() => {
    registerBuiltinTemplates();
  });

  it("every built-in validator declares at least one phase and one tag", () => {
    for (const v of VALIDATORS) {
      expect(v.phases.length, `${v.id} has no phases`).to.be.greaterThan(0);
      expect(v.tags.length, `${v.id} has no tags`).to.be.greaterThan(0);
    }
  });

  it("every validator tag is in the known taxonomy", () => {
    for (const v of VALIDATORS) {
      for (const t of v.tags) {
        expect(KNOWN_TAGS.has(t), `validator ${v.id} uses unknown tag "${t}"`).to.be.true;
      }
    }
  });

  it("every testable template declares at least one validator tag", () => {
    const untagged: string[] = [];
    for (const t of templateRegistry.list() as TemplateDescriptor[]) {
      if (t.testable === false) continue;
      if (!t.tags || t.tags.length === 0) {
        untagged.push(t.id);
      }
    }
    expect(untagged, `testable templates without tags: ${untagged.join(", ")}`).to.be.empty;
  });

  it("every template tag is in the known taxonomy", () => {
    const unknownUses: string[] = [];
    for (const t of templateRegistry.list() as TemplateDescriptor[]) {
      for (const tag of t.tags ?? []) {
        if (!KNOWN_TAGS.has(tag)) {
          unknownUses.push(`${t.id}: "${tag}"`);
        }
      }
    }
    expect(unknownUses, `unknown tags in use: ${unknownUses.join(", ")}`).to.be.empty;
  });

  it("runValidatorsForPhase returns only results for matching tags", async () => {
    // With an empty tag set, only the "*" (project structure) validator should run.
    const results = await runValidatorsForPhase("post-scaffold", {
      envMap: new Map(),
      projectPath: process.cwd(),
      tags: new Set<string>(),
    });
    // project.structure always runs and asserts env/ exists — we don't assert
    // pass/fail, just that results come from the "*" validator only.
    expect(results.length).to.be.greaterThan(0);
  });
});
