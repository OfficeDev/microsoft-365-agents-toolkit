// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { assert } from "chai";
import path from "path";
import { createSandbox, SinonSandbox, SinonStub } from "sinon";
import { featureFlagManager, FeatureFlags } from "../../../src/common/featureFlags";
import { pathUtils } from "../../../src/component/utils/pathUtils";
import { getProjectSettingsPath } from "../../../src/core/middleware/projectSettingsLoader";

describe("projectSettingsLoader - getProjectSettingsPath", () => {
  let sandbox: SinonSandbox;
  let flagStub: SinonStub;
  let availablePathStub: SinonStub;
  let ymlPathStub: SinonStub;

  const projectPath = "/tmp/project";

  beforeEach(() => {
    sandbox = createSandbox();
    flagStub = sandbox
      .stub(featureFlagManager, "getBooleanValue")
      .withArgs(FeatureFlags.GenerateConfigFiles)
      .returns(true);
    availablePathStub = sandbox.stub(pathUtils, "getAvailableYmlFilePath");
    ymlPathStub = sandbox.stub(pathUtils, "getYmlFilePath");
  });

  afterEach(() => {
    sandbox.restore();
  });

  it("returns playground config path when flag enabled and playground exists", () => {
    const playgroundPath = path.join(projectPath, "m365agents.playground.yml");
    availablePathStub.returns(playgroundPath);

    const result = getProjectSettingsPath(projectPath);

    assert.equal(result, playgroundPath);
    assert.isTrue(availablePathStub.calledOnce);
    assert.isTrue(flagStub.calledOnce);
    assert.isTrue(ymlPathStub.notCalled);
  });
});
