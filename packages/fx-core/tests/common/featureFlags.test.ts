// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import "mocha";

import * as chai from "chai";
import chaiAsPromised from "chai-as-promised";
import mockedEnv, { RestoreFn } from "mocked-env";

import { FeatureFlags, featureFlagManager } from "../../src/common/featureFlags";
chai.use(chaiAsPromised);

describe("FeatureFlagManager", () => {
  let mockedEnvRestore: RestoreFn = () => {};
  afterEach(() => {
    mockedEnvRestore();
  });
  it("getBooleanValue, getStringValue is true", async () => {
    mockedEnvRestore = mockedEnv({ TEAMSFX_CLI_DOTNET: "true" });
    const booleanRes = featureFlagManager.getBooleanValue(FeatureFlags.CLIDotNet);
    chai.assert.isTrue(booleanRes);
    const stringRes = featureFlagManager.getStringValue(FeatureFlags.CLIDotNet);
    chai.assert.equal(stringRes, "true");
  });
  it("setBooleanValue", async () => {
    mockedEnvRestore = mockedEnv({ TEAMSFX_CLI_DOTNET: "false" });
    featureFlagManager.setBooleanValue(FeatureFlags.CLIDotNet, true);
    const booleanRes = featureFlagManager.getBooleanValue(FeatureFlags.CLIDotNet);
    chai.assert.isTrue(booleanRes);
  });
  it("getBooleanValue, getStringValue is false", async () => {
    mockedEnvRestore = mockedEnv({ TEAMSFX_CLI_DOTNET: "false" });
    const booleanRes = featureFlagManager.getBooleanValue(FeatureFlags.CLIDotNet);
    chai.assert.isFalse(booleanRes);
    const stringRes = featureFlagManager.getStringValue(FeatureFlags.CLIDotNet);
    chai.assert.equal(stringRes, "false");
  });
  it("list", async () => {
    const list = featureFlagManager.list();
    chai.assert.deepEqual(list, Object.values(FeatureFlags));
  });
  it("listEnabled", async () => {
    mockedEnvRestore = mockedEnv({ TEAMSFX_CLI_DOTNET: "true", SME_OAUTH: "true" });
    const list = featureFlagManager.listEnabled();
    chai.assert.deepEqual(list, ["TEAMSFX_CLI_DOTNET", "SME_OAUTH"]);
  });
});
