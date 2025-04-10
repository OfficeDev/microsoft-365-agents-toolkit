// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import { expect } from "chai";
import "mocha";
import { isValidProject, validateProjectSettings } from "../../src/common/projectSettingsHelper";
describe("tools", () => {
  it("is not valid project", () => {
    expect(isValidProject()).is.false;
  });

  it("validateProjectSettings()", () => {
    const projectSettings: any = {
      appName: "my app",
      projectId: "123234",
      solutionSettings: {
        name: "test",
        version: "3.0.0",
        capabilities: ["Tab"],
        hostType: "Azure",
        azureResources: [],
        activeResourcePlugins: [],
      },
    };
    expect(validateProjectSettings(projectSettings)).is.undefined;
  });

  it("validateProjectSettings() - migrateFromV1", () => {
    const projectSettings: any = {
      appName: "my app",
      projectId: "123234",
      solutionSettings: {
        name: "test",
        version: "3.0.0",
        capabilities: ["Tab"],
        hostType: "Azure",
        azureResources: [],
        activeResourcePlugins: [],
        migrateFromV1: true,
      },
    };
    expect(validateProjectSettings(projectSettings)).is.not.undefined;
  });
});

describe("redactObject", () => {
  const testCases: [unknown, unknown, unknown][] = [
    // happy path: redact unknown key and all values
    [
      { name: "user content 1", userContent2: "user content 3" },
      { type: "object", properties: { name: { type: "string" } } },
      { name: null },
    ],
    // no known keys
    [
      { userContent2: "user content 3" },
      { type: "object", properties: { name: { type: "string" } } },
      {},
    ],
    // second level
    [
      {
        appName: { short: "short name", long: "long name", other: "other name" },
        userContent2: "user content 3",
      },
      {
        type: "object",
        properties: {
          appName: {
            type: "object",
            properties: { short: { type: "string" }, long: { type: "string" } },
          },
        },
      },
      { appName: { short: null, long: null } },
    ],
    // user specified a wrong type
    [
      { appName: "", userContent2: "user content 3" },
      { type: "object", properties: { appName: { type: "object" } } },
      { appName: null },
    ],
    // null input
    [null, { type: "object", properties: { appName: { type: "object" } } }, null],
    [{ name: "name" }, null, null],
    // invalid JSON schema, though not likely because these are checked compile time and not changed frequently
    [{ name: "name" }, { type: "unkown" }, null],
    [{ name: "name" }, { type: "object", properties: "invalid" }, null],
    [{ name: "name" }, { type: "object", properties: { name: "name" } }, { name: null }],
    [{ name: "name" }, { type: "object", properties: { name: {} } }, { name: null }],
  ];
});
