// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { Inputs, Platform } from "@microsoft/teamsfx-api";
import { assert } from "chai";
import "mocha";
import sinon from "sinon";
import { FxCore } from "../../src";
import { FeatureFlags, featureFlagManager } from "../../src/common/featureFlags";
import { setTools } from "../../src/common/globalVars";
import * as generator from "../../src/component/generator/openPlugin/generator";
import { MockTools } from "./utils";

describe("FxCore.convertOpenPlugin", () => {
  const tools = new MockTools();
  const sandbox = sinon.createSandbox();
  let savedFlag: boolean;

  beforeEach(() => {
    setTools(tools);
    savedFlag = featureFlagManager.getBooleanValue(FeatureFlags.OpenPluginConvert);
  });

  afterEach(() => {
    sandbox.restore();
    featureFlagManager.setBooleanValue(FeatureFlags.OpenPluginConvert, savedFlag);
  });

  it("returns FeatureFlagDisabled when feature flag is off", async () => {
    featureFlagManager.setBooleanValue(FeatureFlags.OpenPluginConvert, false);
    const core = new FxCore(tools);
    const inputs: Inputs = {
      platform: Platform.CLI,
      path: "/tmp/some-plugin",
      "privacy-url": "https://example.com/privacy",
      "terms-url": "https://example.com/terms",
    };
    const res = await core.convertOpenPlugin(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.name, "FeatureFlagDisabled");
    }
  });

  it("returns MissingRequiredInput when path is absent", async () => {
    featureFlagManager.setBooleanValue(FeatureFlags.OpenPluginConvert, true);
    const core = new FxCore(tools);
    const inputs: Inputs = {
      platform: Platform.CLI,
      "privacy-url": "https://example.com/privacy",
      "terms-url": "https://example.com/terms",
    };
    const res = await core.convertOpenPlugin(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.name, "MissingRequiredInput");
      assert.include(res.error.message, "path");
    }
  });

  it("returns MissingRequiredInput when privacy-url is absent", async () => {
    featureFlagManager.setBooleanValue(FeatureFlags.OpenPluginConvert, true);
    const core = new FxCore(tools);
    const inputs: Inputs = {
      platform: Platform.CLI,
      path: "/tmp/some-plugin",
      "terms-url": "https://example.com/terms",
    };
    const res = await core.convertOpenPlugin(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.name, "MissingRequiredInput");
      assert.include(res.error.message, "privacy-url");
    }
  });

  it("returns MissingRequiredInput when terms-url is absent", async () => {
    featureFlagManager.setBooleanValue(FeatureFlags.OpenPluginConvert, true);
    const core = new FxCore(tools);
    const inputs: Inputs = {
      platform: Platform.CLI,
      path: "/tmp/some-plugin",
      "privacy-url": "https://example.com/privacy",
    };
    const res = await core.convertOpenPlugin(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.name, "MissingRequiredInput");
      assert.include(res.error.message, "terms-url");
    }
  });

  it("returns InvalidDefaultAuthType for unrecognized auth type", async () => {
    featureFlagManager.setBooleanValue(FeatureFlags.OpenPluginConvert, true);
    const core = new FxCore(tools);
    const inputs: Inputs = {
      platform: Platform.CLI,
      path: "/tmp/some-plugin",
      "privacy-url": "https://example.com/privacy",
      "terms-url": "https://example.com/terms",
      "default-auth-type": "BadAuth",
    };
    const res = await core.convertOpenPlugin(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.name, "InvalidDefaultAuthType");
    }
  });

  it("delegates to convertOpenPlugin on valid inputs", async () => {
    featureFlagManager.setBooleanValue(FeatureFlags.OpenPluginConvert, true);
    const stub = sandbox
      .stub(generator, "convertOpenPlugin")
      .resolves({ isOk: () => true, isErr: () => false, value: { projectPath: "/tmp/out", warnings: ["w1"] } } as any);
    const core = new FxCore(tools);
    const inputs: Inputs = {
      platform: Platform.CLI,
      path: "/tmp/some-plugin",
      "privacy-url": "https://example.com/privacy",
      "terms-url": "https://example.com/terms",
      "website-url": "https://example.com",
      "app-id": "00000000-0000-0000-0000-000000000001",
      "default-auth-type": "None",
      output: "/tmp/out",
      "package-name": "com.example",
    };
    const res = await core.convertOpenPlugin(inputs);
    assert.isTrue(res.isOk());
    assert.isTrue(stub.calledOnce);
    const passedInputs = stub.firstCall.args[0];
    assert.equal(passedInputs.path, "/tmp/some-plugin");
    assert.equal(passedInputs.privacyUrl, "https://example.com/privacy");
    assert.equal(passedInputs.termsUrl, "https://example.com/terms");
    assert.equal(passedInputs.websiteUrl, "https://example.com");
    assert.equal(passedInputs.appId, "00000000-0000-0000-0000-000000000001");
    assert.equal(passedInputs.defaultAuthType, "None");
    assert.equal(passedInputs.output, "/tmp/out");
    assert.equal(passedInputs.packageName, "com.example");
  });

  it("forwards error from convertOpenPlugin generator", async () => {
    featureFlagManager.setBooleanValue(FeatureFlags.OpenPluginConvert, true);
    const { err: makeErr, UserError } = await import("@microsoft/teamsfx-api");
    sandbox
      .stub(generator, "convertOpenPlugin")
      .resolves(makeErr(new UserError("OpenPluginConvert", "SomeError", "boom")));
    const core = new FxCore(tools);
    const inputs: Inputs = {
      platform: Platform.CLI,
      path: "/tmp/some-plugin",
      "privacy-url": "https://example.com/privacy",
      "terms-url": "https://example.com/terms",
    };
    const res = await core.convertOpenPlugin(inputs);
    assert.isTrue(res.isErr());
    if (res.isErr()) {
      assert.equal(res.error.name, "SomeError");
    }
  });
});
