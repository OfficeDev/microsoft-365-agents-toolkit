// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";
import mockedEnv, { RestoreFn } from "mocked-env";
import {
  loadStateFromEnv,
  mapStateToEnv,
  updateVersionForTeamsAppYamlFile,
} from "../../../../src/component/driver/util/utils";
import fs from "fs-extra";
import { expect } from "chai";
import sinon from "sinon";

describe("loadStateFromEnv", () => {
  let envRestore: RestoreFn | undefined;

  afterEach(() => {
    if (envRestore) {
      envRestore();
      envRestore = undefined;
    }
  });

  it("should return empty object when outputEnvVarNames is empty", () => {
    const outputEnvVarNames: Map<string, string> = new Map<string, string>();
    const result = loadStateFromEnv(outputEnvVarNames);
    expect(Object.entries(result).length).to.equal(0);
  });

  it("should return state object with value from env", () => {
    envRestore = mockedEnv({
      ENV_A: "ENV_A value",
      ENV_B: "ENV_B value",
    });
    const outputEnvVarNames: Map<string, string> = new Map(
      Object.entries({
        envA: "ENV_A",
        envB: "ENV_B",
      })
    );

    const result = loadStateFromEnv(outputEnvVarNames);
    expect(Object.entries(result).length).to.equal(2);
    expect(result.envA).to.equal("ENV_A value");
    expect(result.envB).to.equal("ENV_B value");
  });

  it("should return state object with undefined property if env does not exist", () => {
    envRestore = mockedEnv({
      ENV_A: "ENV_A value",
    });
    const outputEnvVarNames: Map<string, string> = new Map(
      Object.entries({
        envA: "ENV_A",
        envB: "ENV_B",
      })
    );

    const result = loadStateFromEnv(outputEnvVarNames);
    expect(Object.entries(result).length).to.equal(2);
    expect(result.envA).to.equal("ENV_A value");
    expect(result.envB).to.be.undefined;
  });
});

describe("mapStateToEnv", async () => {
  it("should convert state to env based on outputEnvVarNames", () => {
    const state: Record<string, string> = {
      envA: "ENV_A value",
      envB: "ENV_B value",
    };
    let outputEnvVarNames: Map<string, string> = new Map(
      Object.entries({
        envA: "ENV_A",
      })
    );
    let result = mapStateToEnv(state, outputEnvVarNames);
    expect(result.size).to.equal(1);
    expect(result.get("ENV_A")).to.equal("ENV_A value");

    outputEnvVarNames = new Map(
      Object.entries({
        envA: "ENV_A",
        envB: "ENV_B",
      })
    );
    result = mapStateToEnv(state, outputEnvVarNames);
    expect(result.size).to.equal(2);
    expect(result.get("ENV_A")).to.equal("ENV_A value");
    expect(result.get("ENV_B")).to.equal("ENV_B value");

    outputEnvVarNames = new Map();
    result = mapStateToEnv(state, outputEnvVarNames);
    expect(result.size).to.equal(0);
  });

  it("should convert state to env and exclude given properties", () => {
    const state: Record<string, string> = {
      envA: "ENV_A value",
      envB: "ENV_B value",
    };
    const outputEnvVarNames: Map<string, string> = new Map(
      Object.entries({
        envA: "ENV_A",
        envB: "ENV_B",
      })
    );
    const result = mapStateToEnv(state, outputEnvVarNames, ["envB"]);
    expect(result.size).to.equal(1);
    expect(result.get("ENV_A")).to.equal("ENV_A value");
  });
});

describe("updateVersionForTeamsAppYamlFile", async () => {
  afterEach(() => {
    sinon.restore();
  });
  it("updateVersionForTeamsAppYamlFile should works fine", async () => {
    const teamsAppYaml = "version: v1.7";
    const expectedTeamsAppYaml = "version: v1.8";

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(teamsAppYaml as any);
    const writeFileStub = sinon.stub(fs, "writeFile");

    await updateVersionForTeamsAppYamlFile("fake-project-path");

    const writtenContent = writeFileStub.getCall(0).args[1];
    // use epect instead
    expect(writtenContent).to.include(expectedTeamsAppYaml);
  });

  it("updateVersionForTeamsAppYamlFile should works fine when yaml contains schema url", async () => {
    const teamsAppYaml = `# yaml-language-server: $schema=https://aka.ms/teams-toolkit/v1.7/yaml.schema.json
# Visit https://aka.ms/teamsfx-v5.0-guide for details on this file
# Visit https://aka.ms/teamsfx-actions for details on actions
version: v1.7`;
    const expectedTeamsAppYaml = `# yaml-language-server: $schema=https://aka.ms/teams-toolkit/v1.8/yaml.schema.json
# Visit https://aka.ms/teamsfx-v5.0-guide for details on this file
# Visit https://aka.ms/teamsfx-actions for details on actions
version: v1.8`;

    sinon.stub(fs, "pathExists").resolves(true);
    sinon.stub(fs, "readFile").resolves(teamsAppYaml as any);
    const writeFileStub = sinon.stub(fs, "writeFile");

    await updateVersionForTeamsAppYamlFile("fake-project-path");

    const writtenContent = writeFileStub.getCall(0).args[1];
    expect(writtenContent).to.include(expectedTeamsAppYaml);
  });
});
