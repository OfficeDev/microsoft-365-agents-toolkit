/**
 * Copyright (c) Microsoft Corporation.
 * Licensed under the MIT license.
 */

import { expect } from "chai";
import { describe, it } from "mocha";
import { analyzeSteps } from "../../../src/lifecycle/analyze";
import type { DriverStep } from "../../../src/lifecycle/types";

function step(uses: string, config?: Record<string, unknown>): DriverStep {
  return { uses, with: config ?? {} };
}

describe("analyzeSteps", () => {
  it("should detect M365 drivers", () => {
    const steps = [step("teamsApp/create"), step("teamsApp/configure")];
    const result = analyzeSteps(steps);
    expect(result.needsM365).to.be.true;
    expect(result.needsAzure).to.be.false;
    expect(result.driverIds).to.deep.equal(["teamsApp/create", "teamsApp/configure"]);
  });

  it("should detect Azure drivers", () => {
    const steps = [step("arm/deploy"), step("azureFunctions/zipDeploy")];
    const result = analyzeSteps(steps);
    expect(result.needsM365).to.be.false;
    expect(result.needsAzure).to.be.true;
  });

  it("should detect both M365 and Azure", () => {
    const steps = [step("teamsApp/create"), step("arm/deploy"), step("aadApp/create")];
    const result = analyzeSteps(steps);
    expect(result.needsM365).to.be.true;
    expect(result.needsAzure).to.be.true;
  });

  it("should detect neither for file-only drivers", () => {
    const steps = [step("file/createOrUpdateEnvironmentFile"), step("script")];
    const result = analyzeSteps(steps);
    expect(result.needsM365).to.be.false;
    expect(result.needsAzure).to.be.false;
  });

  it("should collect unresolved placeholders without envMap", () => {
    const steps = [
      step("arm/deploy", {
        subscriptionId: "${{AZURE_SUBSCRIPTION_ID}}",
        resourceGroupName: "${{AZURE_RESOURCE_GROUP_NAME}}",
      }),
    ];
    const result = analyzeSteps(steps);
    expect(result.unresolvedVars).to.include("AZURE_SUBSCRIPTION_ID");
    expect(result.unresolvedVars).to.include("AZURE_RESOURCE_GROUP_NAME");
  });

  it("should detect resolved placeholders with envMap", () => {
    const steps = [
      step("arm/deploy", {
        subscriptionId: "${{AZURE_SUBSCRIPTION_ID}}",
        resourceGroupName: "${{AZURE_RESOURCE_GROUP_NAME}}",
      }),
    ];
    const envMap = new Map([["AZURE_SUBSCRIPTION_ID", "sub-123"]]);
    const result = analyzeSteps(steps, envMap);
    // Only RG is unresolved
    expect(result.unresolvedVars).to.not.include("AZURE_SUBSCRIPTION_ID");
    expect(result.unresolvedVars).to.include("AZURE_RESOURCE_GROUP_NAME");
  });

  it("should deduplicate unresolved vars", () => {
    const steps = [
      step("arm/deploy", { sub: "${{SUB_ID}}" }),
      step("azureFunctions/zipDeploy", { sub: "${{SUB_ID}}" }),
    ];
    const result = analyzeSteps(steps);
    const count = result.unresolvedVars.filter((v) => v === "SUB_ID").length;
    expect(count).to.equal(1);
  });

  it("should handle empty steps", () => {
    const result = analyzeSteps([]);
    expect(result.needsM365).to.be.false;
    expect(result.needsAzure).to.be.false;
    expect(result.driverIds).to.be.empty;
    expect(result.unresolvedVars).to.be.empty;
  });

  it("should collect nested placeholders", () => {
    const steps = [
      step("file/createOrUpdateJsonFile", {
        target: "config.json",
        content: {
          nested: {
            value: "${{DEEP_VAR}}",
          },
          list: ["${{LIST_VAR}}"],
        },
      }),
    ];
    const result = analyzeSteps(steps);
    expect(result.unresolvedVars).to.include("DEEP_VAR");
    expect(result.unresolvedVars).to.include("LIST_VAR");
  });
});
