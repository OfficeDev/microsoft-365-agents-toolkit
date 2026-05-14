/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 *
 * Guard-rail test: auto-verifies that every `uses:` driver ID referenced
 * in template YAML files is registered in the core-next DriverRegistry.
 *
 * This test scans ALL m365agents.yml.tpl files under templates/vsc/ and
 * fails if any template references an unregistered driver. Adding a new
 * template with a new driver automatically triggers this test.
 */

import { expect } from "chai";
import { describe, it, before } from "mocha";
import {
  registerBuiltinDrivers,
  builtinDrivers,
  driverRegistry,
} from "@microsoft/teamsfx-core-next";
import { scanAllTemplateDrivers } from "./helpers/templateYamlLoader";

/**
 * Known driver IDs that are intentionally unregistered in core-next.
 * These are skipped in the coverage check with an explanatory comment.
 */
const KNOWN_GAPS: Record<string, string> = {
  "typeSpec/compile": "TypeSpec compiler driver — not yet implemented in core-next",
};

describe("Driver Coverage: all template YAML drivers must be registered", () => {
  before(() => {
    registerBuiltinDrivers();
  });

  it("should have all builtin drivers registered", () => {
    expect(builtinDrivers.length).to.be.greaterThan(0);
    for (const d of builtinDrivers) {
      expect(driverRegistry.has(d.id), `${d.id} should be registered`).to.be.true;
    }
  });

  it("should cover every uses: driver ID across all template YAML files", () => {
    const templateDriverMap = scanAllTemplateDrivers();
    expect(templateDriverMap.size).to.be.greaterThan(0);

    const missing: Array<{ template: string; driver: string }> = [];
    const skipped: Array<{ template: string; driver: string; reason: string }> = [];

    for (const [templatePath, driverIds] of templateDriverMap) {
      for (const driverId of driverIds) {
        if (KNOWN_GAPS[driverId]) {
          skipped.push({ template: templatePath, driver: driverId, reason: KNOWN_GAPS[driverId] });
          continue;
        }
        if (!driverRegistry.has(driverId)) {
          missing.push({ template: templatePath, driver: driverId });
        }
      }
    }

    if (skipped.length > 0) {
      console.log(
        `\n  ⚠ Skipped ${skipped.length} known gap(s):\n` +
          skipped.map((s) => `    - ${s.driver} in ${s.template}: ${s.reason}`).join("\n")
      );
    }

    if (missing.length > 0) {
      const details = missing
        .map((m) => `    - ${m.driver} (referenced in ${m.template})`)
        .join("\n");
      expect.fail(
        `${missing.length} unregistered driver(s) found in template YAML files:\n${details}\n\n` +
          "Register the missing driver(s) in packages/core-next/src/drivers/builtin/index.ts,\n" +
          "or add them to KNOWN_GAPS in this test file if intentionally deferred."
      );
    }
  });

  it("should report the total number of unique driver IDs across all templates", () => {
    const templateDriverMap = scanAllTemplateDrivers();
    const allDriverIds = new Set<string>();
    for (const driverIds of templateDriverMap.values()) {
      for (const id of driverIds) {
        allDriverIds.add(id);
      }
    }
    console.log(
      `\n  📊 ${templateDriverMap.size} templates, ${allDriverIds.size} unique driver IDs, ` +
        `${builtinDrivers.length} registered drivers`
    );
    // At minimum, the registered drivers should cover most template drivers
    const knownGapCount = Object.keys(KNOWN_GAPS).length;
    const covered = [...allDriverIds].filter((id) => driverRegistry.has(id)).length;
    expect(covered).to.be.greaterThanOrEqual(allDriverIds.size - knownGapCount);
  });
});
